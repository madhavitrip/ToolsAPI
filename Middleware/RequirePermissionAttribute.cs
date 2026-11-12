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
    public class RequirePermissionAttribute : Attribute, IAsyncActionFilter
    {
        public string ModuleCode { get; }
        public PermissionAction Action { get; }

        public RequirePermissionAttribute(string moduleCode, PermissionAction action)
        {
            ModuleCode = moduleCode;
            Action = action;
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
                // Fallback: Check standard authorization headers
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

            // Extract projectId if available in arguments, query, or route
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
            else
            {
                // Try finding any argument object with a ProjectId property
                foreach (var arg in context.ActionArguments.Values)
                {
                    if (arg == null) continue;
                    var prop = arg.GetType().GetProperty("ProjectId", System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.IgnoreCase);
                    if (prop != null && (prop.PropertyType == typeof(int) || prop.PropertyType == typeof(int?)))
                    {
                        var val = (int?)prop.GetValue(arg);
                        if (val.HasValue && val.Value > 0)
                        {
                            projectId = val.Value;
                            break;
                        }
                    }
                }
            }

            bool hasPermission = await permissionService.HasPermissionAsync(userId, ModuleCode, Action, projectId, httpContext.Request);
            if (!hasPermission)
            {
                var logger = httpContext.RequestServices.GetService<ILoggerService>();
                if (logger != null)
                {
                    await logger.LogEventAsync(
                        $"AUTHORIZATION DENIED: UserId {userId} attempted {Action} on module '{ModuleCode}' (ProjectId: {projectId}) without permission.",
                        "Permissions",
                        userId,
                        projectId
                    );
                }

                context.Result = new ObjectResult(new { 
                    message = $"Access Denied: You do not have permission to {Action.ToString().ToLower()} in module '{ModuleCode}'." 
                }) { StatusCode = 403 };
                return;
            }

            await next();
        }
    }
}
