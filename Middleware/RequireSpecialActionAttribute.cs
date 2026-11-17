using System;
using System.IdentityModel.Tokens.Jwt;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Filters;
using Microsoft.Extensions.DependencyInjection;
using Tools.Models;
using Tools.Services;

namespace Tools.Middleware
{
    [AttributeUsage(AttributeTargets.Class | AttributeTargets.Method, AllowMultiple = true, Inherited = true)]
    public class RequireSpecialActionAttribute : Attribute, IAsyncActionFilter
    {
        public string ActionCode { get; }
        public string ParentModule { get; }
        public PermissionAction RequiredStandardAction { get; }

        public RequireSpecialActionAttribute(string actionCode, string parentModule, PermissionAction requiredStandardAction)
        {
            ActionCode = actionCode;
            ParentModule = parentModule;
            RequiredStandardAction = requiredStandardAction;
        }

        public async Task OnActionExecutionAsync(ActionExecutingContext context, ActionExecutionDelegate next)
        {
            var httpContext = context.HttpContext;
            var permissionService = httpContext.RequestServices.GetService<IPermissionService>();

            if (permissionService == null)
            {
                await next();
                return;
            }

            int userId = LogHelper.GetTriggeredBy(httpContext.User, httpContext.Request);
            if (userId <= 0)
            {
                var authHeader = httpContext.Request.Headers["Authorization"].ToString();
                if (!string.IsNullOrWhiteSpace(authHeader))
                {
                    var token = authHeader.Replace("Bearer ", "", StringComparison.OrdinalIgnoreCase).Trim();
                    try
                    {
                        var handler = new JwtSecurityTokenHandler();
                        if (handler.CanReadToken(token))
                        {
                            var jwt = handler.ReadToken(token) as JwtSecurityToken;
                            var nameClaim = jwt?.Claims.FirstOrDefault(c =>
                                c.Type == ClaimTypes.Name ||
                                c.Type == "sub" ||
                                c.Type == "userid" ||
                                c.Type == "userId")?.Value;
                            if (int.TryParse(nameClaim, out int parsed))
                            {
                                userId = parsed;
                            }
                        }
                    }
                    catch { }
                }
            }

            if (userId <= 0)
            {
                context.Result = new ObjectResult(new { message = "User authentication required." }) { StatusCode = 401 };
                return;
            }

            int projectId = 0;
            var matchedArg = context.ActionArguments.FirstOrDefault(kvp => string.Equals(kvp.Key, "projectId", StringComparison.OrdinalIgnoreCase));
            if (matchedArg.Value is int pId)
            {
                projectId = pId;
            }
            else if (httpContext.Request.Query.TryGetValue("projectId", out var qpId) && int.TryParse(qpId, out int qpIdVal))
            {
                projectId = qpIdVal;
            }
            else if (context.RouteData.Values.TryGetValue("projectId", out var rpId) && int.TryParse(rpId?.ToString(), out int rpIdVal))
            {
                projectId = rpIdVal;
            }

            bool isAllowed = await permissionService.HasSpecialActionAsync(userId, ActionCode, ParentModule, RequiredStandardAction, projectId, httpContext.Request);
            if (!isAllowed)
            {
                var logger = httpContext.RequestServices.GetService<ILoggerService>();
                if (logger != null)
                {
                    await logger.LogEventAsync(
                        $"SPECIAL ACTION DENIED: UserId {userId} attempted '{ActionCode}' on module '{ParentModule}' (ProjectId: {projectId}).",
                        "Permissions",
                        userId,
                        projectId
                    );
                }

                context.Result = new ObjectResult(new
                {
                    message = $"Access Denied: You do not have authorization to perform special action '{ActionCode}'. Requires {ParentModule} View, {RequiredStandardAction}, and explicit {ActionCode} permission."
                })
                { StatusCode = 403 };
                return;
            }

            await next();
        }
    }
}
