using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ERPToolsAPI.Data;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Filters;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Tools.Services;

namespace Tools.Middleware
{
    public class ApiAuditLoggingFilter : IAsyncActionFilter
    {
        private readonly ILoggerService _loggerService;
        private readonly IServiceScopeFactory _scopeFactory;
        private readonly ILogger<ApiAuditLoggingFilter> _logger;

        public ApiAuditLoggingFilter(
            ILoggerService loggerService,
            IServiceScopeFactory scopeFactory,
            ILogger<ApiAuditLoggingFilter> logger)
        {
            _loggerService = loggerService;
            _scopeFactory = scopeFactory;
            _logger = logger;
        }

        public async Task OnActionExecutionAsync(ActionExecutingContext context, ActionExecutionDelegate next)
        {
            var httpMethod = context.HttpContext.Request.Method.ToUpperInvariant();

            // Only audit POST and PUT requests
            if (httpMethod != "POST" && httpMethod != "PUT")
            {
                await next();
                return;
            }

            // Check if endpoint is annotated with [NoAuditLog]
            var endpoint = context.HttpContext.GetEndpoint();
            if (endpoint?.Metadata.GetMetadata<NoAuditLogAttribute>() != null)
            {
                await next();
                return;
            }

            // Determine category (attribute or controller name)
            var categoryAttr = endpoint?.Metadata.GetMetadata<AuditCategoryAttribute>();
            var controllerName = context.Controller.GetType().Name.Replace("Controller", "");
            var category = categoryAttr?.Category ?? controllerName;

            // Extract user ID using LogHelper
            int userId = LogHelper.GetTriggeredBy(context.HttpContext.User, context.HttpContext.Request);

            // Extract project ID from route, query, or arguments
            int projectId = ExtractProjectId(context);

            // Capture old value for PUT operations
            string? oldValue = null;
            if (httpMethod == "PUT")
            {
                oldValue = await TryFetchOldValueAsync(context);
            }

            // Identify candidate incoming payload for new value
            object? incomingPayload = ExtractIncomingPayload(context);

            ActionExecutedContext executedContext;
            try
            {
                executedContext = await next();
            }
            catch (Exception ex)
            {
                var failedPath = context.HttpContext.Request.Path + context.HttpContext.Request.QueryString;
                var payloadJson = AuditJsonHelper.SerializeSafe(incomingPayload);

                await _loggerService.LogApiActivityAsync(
                    httpMethod: httpMethod,
                    endpoint: failedPath,
                    category: category,
                    triggeredBy: userId,
                    projectId: projectId,
                    oldValue: oldValue,
                    newValue: payloadJson,
                    statusCode: 500,
                    details: $"Exception: {ex.Message}"
                );

                throw;
            }

            // Determine HTTP status code
            int statusCode = executedContext.HttpContext.Response.StatusCode;
            if (executedContext.Result is ObjectResult objResult && objResult.StatusCode.HasValue)
            {
                statusCode = objResult.StatusCode.Value;
            }
            else if (executedContext.Result is StatusCodeResult statusResult)
            {
                statusCode = statusResult.StatusCode;
            }

            // Prioritize returned object if successful, else use incoming payload
            object? finalNewValue = incomingPayload;
            if (executedContext.Result is ObjectResult returnedObj && returnedObj.Value != null && statusCode < 400)
            {
                finalNewValue = returnedObj.Value;
            }

            string? newValueJson = AuditJsonHelper.SerializeSafe(finalNewValue);
            var path = context.HttpContext.Request.Path + context.HttpContext.Request.QueryString;
            string? actionName = context.ActionDescriptor.RouteValues.TryGetValue("action", out var act) ? act : "";
            var details = actionName;

            try
            {
                await _loggerService.LogApiActivityAsync(
                    httpMethod: httpMethod,
                    endpoint: path,
                    category: category,
                    triggeredBy: userId,
                    projectId: projectId,
                    oldValue: oldValue,
                    newValue: newValueJson,
                    statusCode: statusCode,
                    details: actionName
                );
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error while saving centralized audit log for {Path}", path);
            }
        }

        private async Task<string?> TryFetchOldValueAsync(ActionExecutingContext context)
        {
            try
            {
                using var scope = _scopeFactory.CreateScope();
                var db = scope.ServiceProvider.GetRequiredService<ERPToolsDbContext>();

                // Attempt to resolve primary key value from route or action arguments
                object? primaryKeyValue = null;
                if (context.ActionArguments.TryGetValue("id", out var idArg) && idArg != null)
                {
                    primaryKeyValue = idArg;
                }
                else if (context.RouteData.Values.TryGetValue("id", out var routeId) && routeId != null)
                {
                    primaryKeyValue = routeId;
                }

                foreach (var arg in context.ActionArguments.Values)
                {
                    if (arg == null) continue;

                    var argType = arg.GetType();
                    var entityType = db.Model.FindEntityType(argType);
                    if (entityType != null)
                    {
                        var pk = entityType.FindPrimaryKey();
                        if (pk != null && pk.Properties.Count == 1)
                        {
                            var pkProp = pk.Properties[0];
                            var keyVal = primaryKeyValue;
                            if (keyVal == null && pkProp.PropertyInfo != null)
                            {
                                keyVal = pkProp.PropertyInfo.GetValue(arg);
                            }

                            if (keyVal != null)
                            {
                                try
                                {
                                    var convertedKey = Convert.ChangeType(keyVal, pkProp.ClrType);
                                    var existing = await db.FindAsync(argType, convertedKey);
                                    if (existing != null)
                                    {
                                        return AuditJsonHelper.SerializeSafe(existing);
                                    }
                                }
                                catch
                                {
                                    // Primary key conversion fallback
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogDebug(ex, "Could not fetch pre-mutation old value for PUT request.");
            }

            return null;
        }

        private static object? ExtractIncomingPayload(ActionExecutingContext context)
        {
            // First look for any complex argument that is not the URL id
            foreach (var kvp in context.ActionArguments)
            {
                if (!string.Equals(kvp.Key, "id", StringComparison.OrdinalIgnoreCase) && kvp.Value != null)
                {
                    var type = kvp.Value.GetType();
                    if (!type.IsPrimitive && type != typeof(string) && type != typeof(decimal))
                    {
                        return kvp.Value;
                    }
                }
            }

            // If only primitive arguments or single argument, return all arguments
            if (context.ActionArguments.Count == 1)
            {
                return context.ActionArguments.Values.First();
            }

            return context.ActionArguments.Count > 0 ? context.ActionArguments : null;
        }

        private static int ExtractProjectId(ActionExecutingContext context)
        {
            // 1. From RouteData
            if (context.RouteData.Values.TryGetValue("projectId", out var routeProjId) &&
                int.TryParse(routeProjId?.ToString(), out var parsedRouteProj) && parsedRouteProj > 0)
            {
                return parsedRouteProj;
            }

            // 2. From QueryString
            if (context.HttpContext.Request.Query.TryGetValue("projectId", out var queryProjId) &&
                int.TryParse(queryProjId.FirstOrDefault(), out var parsedQueryProj) && parsedQueryProj > 0)
            {
                return parsedQueryProj;
            }

            // 3. Inspect ActionArguments for ProjectId or GroupId
            foreach (var arg in context.ActionArguments.Values)
            {
                var extracted = AuditJsonHelper.ExtractProjectId(arg);
                if (extracted > 0)
                {
                    return extracted;
                }
            }

            return 0;
        }
    }
}
