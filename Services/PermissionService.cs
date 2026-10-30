using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using ERPToolsAPI.Data;
using Microsoft.AspNetCore.Http;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Caching.Memory;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Primitives;
using Tools.Models;

namespace Tools.Services
{
    public interface IPermissionService
    {
        Task<bool> HasPermissionAsync(int userId, string moduleCode, PermissionAction action, int projectId = 0, HttpRequest? request = null);
        Task<UserPermissionsDto> GetUserPermissionsAsync(int userId, HttpRequest? request = null, int? explicitErpRoleId = null);
        void InvalidateUserCache(int userId);
        void InvalidateAllCaches();
        Task EnsureDatabaseInitializedAsync();
    }

    public class PermissionService : IPermissionService
    {
        private readonly ERPToolsDbContext _context;
        private readonly IMemoryCache _cache;
        private readonly IConfiguration _configuration;
        private readonly ILogger<PermissionService> _logger;
        private readonly IHttpContextAccessor _httpContextAccessor;
        private readonly string _baseUrl;

        private const string CachePrefix = "user_perm_";
        private static readonly TimeSpan CacheTtl = TimeSpan.FromMinutes(15);
        private static CancellationTokenSource _resetToken = new();
        private static bool _initialized = false;
        private static readonly object _initLock = new();

        public PermissionService(
            ERPToolsDbContext context,
            IMemoryCache cache,
            IConfiguration configuration,
            ILogger<PermissionService> logger,
            IHttpContextAccessor httpContextAccessor)
        {
            _context = context;
            _cache = cache;
            _configuration = configuration;
            _logger = logger;
            _httpContextAccessor = httpContextAccessor;
            _baseUrl = configuration["ApiSettings:BaseUrl"] ?? "http://192.168.10.208:82/API/api";
        }

        public async Task EnsureDatabaseInitializedAsync()
        {
            if (_initialized) return;

            lock (_initLock)
            {
                if (_initialized) return;
            }

            try
            {
                // Seed Default System Roles if table is empty
                if (!await _context.ToolsRoles.AnyAsync())
                {
                    await SeedDefaultRolesAsync();
                }

                _initialized = true;
                _logger.LogInformation("Roles and Permissions verified.");
            }
            catch (Exception ex)
            {
                _logger.LogWarning("Permissions verification check: {Message}", ex.Message);
            }
        }

        public async Task<bool> HasPermissionAsync(
            int userId,
            string moduleCode,
            PermissionAction action,
            int projectId = 0,
            HttpRequest? request = null)
        {
            if (userId <= 0) return false;

            // 1. Project-level boundary check
            if (projectId > 0)
            {
                var projectAuthorized = await CheckProjectAssignmentAsync(userId, projectId, request);
                if (!projectAuthorized) return false;
            }

            // 2. Compute effective permissions (cached)
            var perms = await GetUserPermissionsAsync(userId, request);

            // Fail-safe: Always allow Lead Developer (Role 1) to access and manage the Roles & Permissions module
            if (perms.RoleId == 1 && string.Equals(moduleCode, ModuleCodes.RolesPermissions, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            // 3. Check action flag based on effective computed permissions
            if (perms.Modules.TryGetValue(moduleCode, out var flags) && flags != null)
            {
                return action switch
                {
                    PermissionAction.View => flags.CanView ?? false,
                    PermissionAction.Create => flags.CanCreate ?? false,
                    PermissionAction.Edit => flags.CanEdit ?? false,
                    PermissionAction.Delete => flags.CanDelete ?? false,
                    PermissionAction.Execute => flags.CanExecute ?? false,
                    PermissionAction.Export => flags.CanExport ?? false,
                    _ => false
                };
            }

            return false;
        }

        public async Task<UserPermissionsDto> GetUserPermissionsAsync(int userId, HttpRequest? request = null, int? explicitErpRoleId = null)
        {
            if (userId <= 0) return new UserPermissionsDto { UserId = userId };

            string cacheKey = $"{CachePrefix}{userId}";
            if (!explicitErpRoleId.HasValue && _cache.TryGetValue(cacheKey, out UserPermissionsDto? cached) && cached != null)
            {
                return cached;
            }

            await EnsureDatabaseInitializedAsync();

            var httpReq = request ?? _httpContextAccessor.HttpContext?.Request;
            var claimsUser = _httpContextAccessor.HttpContext?.User;

            int callerUserId = LogHelper.GetTriggeredBy(claimsUser, httpReq);

            int erpRoleId = 0;
            if (explicitErpRoleId.HasValue && explicitErpRoleId.Value > 0)
            {
                erpRoleId = explicitErpRoleId.Value;
            }
            else if (httpReq?.Query.TryGetValue("erpRoleId", out var qVal) == true && int.TryParse(qVal, out var qRoleId) && qRoleId > 0)
            {
                erpRoleId = qRoleId;
            }
            else if (callerUserId > 0 && callerUserId == userId)
            {
                // Only extract role from request token if the token actually belongs to this target user
                erpRoleId = LogHelper.GetUserRoleId(claimsUser, httpReq);
            }
            else
            {
                // Target user is different from caller: fetch their ERP role from ERP User API
                erpRoleId = await FetchErpUserRoleIdAsync(userId, httpReq);
            }

            // Fetch user permission row if exists
            var userRecord = await _context.UserPermissions.AsNoTracking().FirstOrDefaultAsync(u => u.UserId == userId);

            // Inherit ERP role if no local override; default to Operator (5) if unknown, never Developer (1)!
            int effectiveRoleId = userRecord?.RoleId ?? (erpRoleId > 0 ? erpRoleId : 5);
            var role = await _context.ToolsRoles.AsNoTracking().FirstOrDefaultAsync(r => r.Id == effectiveRoleId);

            string roleName = role?.RoleName ?? effectiveRoleId switch
            {
                1 => "Developer",
                2 => "Admin",
                3 => "Head",
                4 => "Manager",
                5 => "Operator",
                _ => $"Role {effectiveRoleId}"
            };

            bool isDeveloper = (effectiveRoleId == 1);

            // Parse Role Permissions JSON
            var rolePermsMap = ParsePermissionsMap(role?.PermissionsJson);

            // Parse User Overrides JSON
            var overridesMap = ParsePermissionsMap(userRecord?.OverridesJson);

            // Build effective permission dictionary across all module codes
            var effectiveModules = new Dictionary<string, ModuleActionFlags>(StringComparer.OrdinalIgnoreCase);

            foreach (var code in ModuleCodes.All)
            {
                rolePermsMap.TryGetValue(code, out var roleFlags);
                overridesMap.TryGetValue(code, out var overrideFlags);

                // User override takes precedence, then role definition, default to false if unspecified
                effectiveModules[code] = new ModuleActionFlags
                {
                    CanView = overrideFlags?.CanView ?? roleFlags?.CanView ?? false,
                    CanCreate = overrideFlags?.CanCreate ?? roleFlags?.CanCreate ?? false,
                    CanEdit = overrideFlags?.CanEdit ?? roleFlags?.CanEdit ?? false,
                    CanDelete = overrideFlags?.CanDelete ?? roleFlags?.CanDelete ?? false,
                    CanExecute = overrideFlags?.CanExecute ?? roleFlags?.CanExecute ?? false,
                    CanExport = overrideFlags?.CanExport ?? roleFlags?.CanExport ?? false,
                };

                // Fail-safe: Ensure Lead Developer (Role 1) always maintains access to the Roles & Permissions module
                // so an administrator can never lock themselves out of managing permissions.
                if (isDeveloper && string.Equals(code, ModuleCodes.RolesPermissions, StringComparison.OrdinalIgnoreCase))
                {
                    effectiveModules[code] = new ModuleActionFlags
                    {
                        CanView = true,
                        CanCreate = true,
                        CanEdit = true,
                        CanDelete = true,
                        CanExecute = true,
                        CanExport = true
                    };
                }
            }

            var result = new UserPermissionsDto
            {
                UserId = userId,
                RoleId = effectiveRoleId,
                RoleName = roleName,
                IsSuperuser = isDeveloper,
                Modules = effectiveModules
            };

            var cacheOptions = new MemoryCacheEntryOptions
            {
                SlidingExpiration = CacheTtl
            }.AddExpirationToken(new CancellationChangeToken(_resetToken.Token));

            _cache.Set(cacheKey, result, cacheOptions);
            return result;
        }

        public void InvalidateUserCache(int userId)
        {
            _cache.Remove($"{CachePrefix}{userId}");
        }

        public void InvalidateAllCaches()
        {
            _logger.LogInformation("Evicting all user permission caches.");
            var oldToken = Interlocked.Exchange(ref _resetToken, new CancellationTokenSource());
            oldToken.Cancel();
            oldToken.Dispose();
        }

        private async Task<int> FetchErpUserRoleIdAsync(int targetUserId, HttpRequest? request)
        {
            if (targetUserId <= 0) return 0;
            try
            {
                using var client = new HttpClient { Timeout = TimeSpan.FromSeconds(4) };
                var authHeader = request?.Headers["Authorization"].ToString();
                if (!string.IsNullOrWhiteSpace(authHeader))
                {
                    client.DefaultRequestHeaders.TryAddWithoutValidation("Authorization", authHeader);
                }

                string url = $"{_baseUrl}/User";
                var response = await client.GetAsync(url);
                if (response.IsSuccessStatusCode)
                {
                    var json = await response.Content.ReadAsStringAsync();
                    using var doc = JsonDocument.Parse(json);
                    JsonElement root = doc.RootElement;
                    JsonElement usersArray = root.ValueKind == JsonValueKind.Array ? root :
                                             root.TryGetProperty("$values", out var vals) ? vals : default;

                    if (usersArray.ValueKind == JsonValueKind.Array)
                    {
                        foreach (var u in usersArray.EnumerateArray())
                        {
                            int uId = 0;
                            if ((u.TryGetProperty("userId", out var p1) || u.TryGetProperty("UserId", out p1) || u.TryGetProperty("id", out p1)) &&
                                p1.TryGetInt32(out uId) && uId == targetUserId)
                            {
                                if ((u.TryGetProperty("roleId", out var rp) || u.TryGetProperty("RoleId", out rp)) &&
                                    rp.TryGetInt32(out var rId) && rId > 0)
                                {
                                    return rId;
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning("FetchErpUserRoleIdAsync failed for user {UserId}: {Message}", targetUserId, ex.Message);
            }

            return 0;
        }

        private async Task<bool> CheckProjectAssignmentAsync(int userId, int projectId, HttpRequest? request)
        {
            var perms = await GetUserPermissionsAsync(userId, request);
            if (perms.RoleId == 1 || perms.IsSuperuser) return true; // Lead Developer has universal project access

            var project = await _context.Projects.AsNoTracking().FirstOrDefaultAsync(p => p.ProjectId == projectId);
            if (project == null) return false;

            return project.UserAssigned != null && project.UserAssigned.Contains(userId);
        }

        private static Dictionary<string, ModuleActionFlags> ParsePermissionsMap(string? json)
        {
            if (string.IsNullOrWhiteSpace(json)) return new Dictionary<string, ModuleActionFlags>(StringComparer.OrdinalIgnoreCase);

            try
            {
                var options = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
                return JsonSerializer.Deserialize<Dictionary<string, ModuleActionFlags>>(json, options)
                    ?? new Dictionary<string, ModuleActionFlags>(StringComparer.OrdinalIgnoreCase);
            }
            catch
            {
                return new Dictionary<string, ModuleActionFlags>(StringComparer.OrdinalIgnoreCase);
            }
        }

        private async Task SeedDefaultRolesAsync()
        {
            var fullAccessFlags = new ModuleActionFlags { CanView = true, CanCreate = true, CanEdit = true, CanDelete = true, CanExecute = true, CanExport = true };
            var fullAccessJson = JsonSerializer.Serialize(ModuleCodes.All.ToDictionary(m => m, _ => fullAccessFlags));

            var managerMap = new Dictionary<string, ModuleActionFlags>(StringComparer.OrdinalIgnoreCase);
            foreach (var m in ModuleCodes.All)
            {
                bool isMaster = m.StartsWith("MASTER_");
                bool isRoles = m == ModuleCodes.RolesPermissions;

                managerMap[m] = new ModuleActionFlags
                {
                    CanView = !isRoles,
                    CanCreate = !isMaster && !isRoles,
                    CanEdit = !isRoles && m != ModuleCodes.MasterProject,
                    CanDelete = !isMaster && !isRoles,
                    CanExecute = !isRoles,
                    CanExport = !isRoles
                };
            }
            var managerJson = JsonSerializer.Serialize(managerMap);

            var operatorMap = new Dictionary<string, ModuleActionFlags>(StringComparer.OrdinalIgnoreCase);
            foreach (var m in ModuleCodes.All)
            {
                bool canView = m is ModuleCodes.Dashboard or ModuleCodes.DataImport or ModuleCodes.ProcessingPipeline
                                or ModuleCodes.HeaderVerification or ModuleCodes.BatchData or ModuleCodes.Reports;
                operatorMap[m] = new ModuleActionFlags
                {
                    CanView = canView,
                    CanCreate = false,
                    CanEdit = false,
                    CanDelete = false,
                    CanExecute = false,
                    CanExport = m is ModuleCodes.BatchData or ModuleCodes.Reports
                };
            }
            var operatorJson = JsonSerializer.Serialize(operatorMap);

            var roles = new[]
            {
                new ToolsRole { Id = 1, RoleName = "Developer", Description = "Lead Developer (Superuser)", IsSystemRole = true, PermissionsJson = fullAccessJson },
                new ToolsRole { Id = 2, RoleName = "Admin", Description = "System Administrator", IsSystemRole = true, PermissionsJson = fullAccessJson },
                new ToolsRole { Id = 3, RoleName = "Head", Description = "Project / Department Head", IsSystemRole = true, PermissionsJson = fullAccessJson },
                new ToolsRole { Id = 4, RoleName = "Manager", Description = "Project Manager", IsSystemRole = true, PermissionsJson = managerJson },
                new ToolsRole { Id = 5, RoleName = "Operator", Description = "Standard Tool Operator", IsSystemRole = true, PermissionsJson = operatorJson }
            };

            foreach (var r in roles)
            {
                var existing = await _context.ToolsRoles.FirstOrDefaultAsync(x => x.Id == r.Id);
                if (existing == null)
                {
                    _context.ToolsRoles.Add(r);
                }
                else
                {
                    existing.RoleName = r.RoleName;
                    existing.Description = r.Description;
                    existing.IsSystemRole = r.IsSystemRole;
                    existing.UpdatedAt = DateTime.Now;
                }
            }
            await _context.SaveChangesAsync();
        }
    }
}
