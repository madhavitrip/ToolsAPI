using System;
using System.Linq;
using System.Reflection;
using System.Security.Claims;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Filters;
using Microsoft.Extensions.DependencyInjection;
using Tools.Services;

namespace Tools.Middleware
{
    [AttributeUsage(AttributeTargets.Class | AttributeTargets.Method, AllowMultiple = false, Inherited = true)]
    public class RequireMasterAuthAttribute : Attribute, IAsyncActionFilter
    {
        public string Module { get; set; } = "Master";
        public string Operation { get; set; } = "MODIFY";

        public async Task OnActionExecutionAsync(ActionExecutingContext context, ActionExecutionDelegate next)
        {
            var httpContext = context.HttpContext;
            var masterAuthService = httpContext.RequestServices.GetService<IMasterAuthService>();

            if (masterAuthService == null)
            {
                context.Result = new ObjectResult(new { message = "Master authorization service is not configured." })
                {
                    StatusCode = 500
                };
                return;
            }

            // Extract passcode from X-Master-Auth-Passcode header
            if (!httpContext.Request.Headers.TryGetValue("X-Master-Auth-Passcode", out var passcodeHeader) ||
                string.IsNullOrWhiteSpace(passcodeHeader.ToString()))
            {
                context.Result = new ObjectResult(new { 
                    message = "Master authorization passcode is required for this operation.",
                    requiresMasterAuth = true 
                })
                {
                    StatusCode = 401
                };
                return;
            }

            string passcode = passcodeHeader.ToString().Trim();

            // Resolve GroupId
            int groupId = ResolveGroupId(context, masterAuthService);

            // Extract User ID from Claims or Header
            int userId = 0;
            var userIdClaim = httpContext.User?.Claims.FirstOrDefault(c => c.Type == "userid" || c.Type == ClaimTypes.NameIdentifier);
            if (userIdClaim != null && int.TryParse(userIdClaim.Value, out var parsedId))
            {
                userId = parsedId;
            }
            else if (httpContext.Request.Headers.TryGetValue("X-User-Id", out var headerUserId) && int.TryParse(headerUserId.ToString(), out var parsedHeaderId))
            {
                userId = parsedHeaderId;
            }

            // Extract IP Address
            string ipAddress = httpContext.Connection.RemoteIpAddress?.ToString() ?? "0.0.0.0";
            if (httpContext.Request.Headers.TryGetValue("X-Forwarded-For", out var forwardedFor))
            {
                ipAddress = forwardedFor.ToString().Split(',')[0].Trim();
            }

            // Verify passcode via MasterAuthService for this GroupId
            var (isValid, errorMessage) = await masterAuthService.VerifyPasscodeAsync(passcode, groupId, userId, Operation, Module, ipAddress);

            if (!isValid)
            {
                context.Result = new ObjectResult(new { message = errorMessage, success = false, groupId })
                {
                    StatusCode = 403
                };
                return;
            }

            // Passcode is valid -> execute controller action
            await next();
        }

        private static int ResolveGroupId(ActionExecutingContext context, IMasterAuthService masterAuthService)
        {
            var req = context.HttpContext.Request;

            // 1. Header X-Group-Id
            if (req.Headers.TryGetValue("X-Group-Id", out var groupHeader) && int.TryParse(groupHeader.ToString(), out var headerGroupId) && headerGroupId > 0)
            {
                return headerGroupId;
            }

            // 2. Query string groupId or GroupId
            if (req.Query.TryGetValue("groupId", out var qGroup) && int.TryParse(qGroup.ToString(), out var queryGroupId) && queryGroupId > 0)
            {
                return queryGroupId;
            }
            if (req.Query.TryGetValue("GroupId", out var qGroupCap) && int.TryParse(qGroupCap.ToString(), out var queryGroupCapId) && queryGroupCapId > 0)
            {
                return queryGroupCapId;
            }

            // 3. Route values
            if (context.RouteData.Values.TryGetValue("groupId", out var rGroup) && int.TryParse(rGroup?.ToString(), out var routeGroupId) && routeGroupId > 0)
            {
                return routeGroupId;
            }

            // 4. Action arguments (inspecting DTO objects and primitive values for GroupId or ProjectId)
            int projectId = 0;
            if (context.ActionArguments.TryGetValue("groupId", out var argGroupId) && argGroupId is int gid && gid > 0)
                return gid;
            if (context.ActionArguments.TryGetValue("TargetGroupId", out var argTargetGroupId) && argTargetGroupId is int tgid && tgid > 0)
                return tgid;

            if (context.ActionArguments.TryGetValue("projectId", out var argProjectId) && argProjectId is int pid && pid > 0)
                projectId = pid;
            else if (context.ActionArguments.TryGetValue("TargetProjectId", out var argTargetProjectId) && argTargetProjectId is int tpid && tpid > 0)
                projectId = tpid;

            foreach (var arg in context.ActionArguments.Values)
            {
                if (arg == null) continue;

                var propGroup = arg.GetType().GetProperty("GroupId", BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase)
                             ?? arg.GetType().GetProperty("TargetGroupId", BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);
                if (propGroup != null && propGroup.PropertyType == typeof(int))
                {
                    if (propGroup.GetValue(arg) is int val && val > 0) return val;
                }

                var propProject = arg.GetType().GetProperty("ProjectId", BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase)
                               ?? arg.GetType().GetProperty("TargetProjectId", BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);
                if (propProject != null && propProject.PropertyType == typeof(int))
                {
                    if (propProject.GetValue(arg) is int pVal && pVal > 0) projectId = pVal;
                }
            }

            // 5. Query string or header ProjectId -> resolve GroupId
            if (projectId <= 0)
            {
                if (req.Headers.TryGetValue("X-Project-Id", out var headerProj) && int.TryParse(headerProj.ToString(), out var parsedHeaderProj))
                {
                    projectId = parsedHeaderProj;
                }
                else if (req.Query.TryGetValue("projectId", out var qProj) && int.TryParse(qProj.ToString(), out var parsedQueryProj))
                {
                    projectId = parsedQueryProj;
                }
            }

            if (projectId > 0)
            {
                return masterAuthService.ResolveGroupIdFromProjectId(projectId);
            }

            return 0;
        }
    }
}
