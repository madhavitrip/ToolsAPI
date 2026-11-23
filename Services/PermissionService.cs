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
        Task<bool> HasSpecialActionAsync(int userId, string actionCode, string parentModule, PermissionAction requiredStandardAction, int projectId = 0, HttpRequest? request = null);
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

            // Dashboard is accessible to every authenticated user without requiring a permission check
            if (string.Equals(moduleCode, ModuleCodes.Dashboard, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            // Compute effective permissions (cached)
            var perms = await GetUserPermissionsAsync(userId, request);

            // 3. Check action flag based on effective computed permissions
            if (perms.Modules.TryGetValue(moduleCode, out var flags) && flags != null)
            {
                return action switch
                {
                    PermissionAction.View => flags.CanView ?? false,
                    PermissionAction.Create => flags.CanCreate ?? false,
                    PermissionAction.Edit => flags.CanEdit ?? false,
                    PermissionAction.Delete => flags.CanDelete ?? false,
                    _ => false
                };
            }

            return false;
        }

        public async Task<bool> HasSpecialActionAsync(
            int userId,
            string actionCode,
            string parentModule,
            PermissionAction requiredStandardAction,
            int projectId = 0,
            HttpRequest? request = null)
        {
            if (userId <= 0) return false;

            var perms = await GetUserPermissionsAsync(userId, request);

            // Condition 1: User must have View permission on Parent Module
            if (!perms.Modules.TryGetValue(parentModule, out var parentFlags) || parentFlags?.CanView != true)
            {
                return false;
            }

            // Condition 2: User must have the required standard action on Parent Module
            bool hasStandardAction = requiredStandardAction switch
            {
                PermissionAction.View => parentFlags.CanView ?? false,
                PermissionAction.Create => parentFlags.CanCreate ?? false,
                PermissionAction.Edit => parentFlags.CanEdit ?? false,
                PermissionAction.Delete => parentFlags.CanDelete ?? false,
                _ => false
            };
            if (!hasStandardAction) return false;

            // Condition 3: User must have the explicit special action permission
            if (perms.SpecialActions != null && perms.SpecialActions.TryGetValue(actionCode, out bool specialAllowed))
            {
                return specialAllowed;
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

            int effectiveRoleId = 0;

            // 1. If explicit role ID is passed in (for administrative testing/preview)
            if (explicitErpRoleId.HasValue && explicitErpRoleId.Value > 0)
            {
                effectiveRoleId = explicitErpRoleId.Value;
            }
            else
            {
                // 2. Query internal Tools role assignment table: UserToolsRoles
                var assignment = await _context.UserToolsRoles.AsNoTracking().FirstOrDefaultAsync(u => u.UserId == userId);
                if (assignment != null && assignment.RoleId > 0)
                {
                    effectiveRoleId = assignment.RoleId;
                }
                else
                {
                    // 3. Independent fallback: Use protected Read Only role (never ERP JWT roleId!)
                    var readOnlyRole = await _context.ToolsRoles.AsNoTracking().FirstOrDefaultAsync(r => r.RoleName == "Read Only");
                    effectiveRoleId = readOnlyRole?.Id ?? 0;
                }
            }

            // Fail-closed: If RoleId is missing, invalid, or <= 0, return empty permissions.
            if (effectiveRoleId <= 0)
            {
                var emptyResult = new UserPermissionsDto
                {
                    UserId = userId,
                    RoleId = 0,
                    RoleName = string.Empty,
                    IsSuperuser = false,
                    Modules = ModuleCodes.All.ToDictionary(
                        code => code,
                        _ => new ModuleActionFlags
                        {
                            CanView = false,
                            CanCreate = false,
                            CanEdit = false,
                            CanDelete = false
                        },
                        StringComparer.OrdinalIgnoreCase),
                    SpecialActions = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase)
                };

                return emptyResult;
            }

            var role = await _context.ToolsRoles.AsNoTracking().FirstOrDefaultAsync(r => r.Id == effectiveRoleId);

            // Fail-closed: If role does not exist in ToolsRoles, return empty permissions
            if (role == null)
            {
                var emptyResult = new UserPermissionsDto
                {
                    UserId = userId,
                    RoleId = effectiveRoleId,
                    RoleName = $"Role {effectiveRoleId}",
                    IsSuperuser = false,
                    Modules = ModuleCodes.All.ToDictionary(
                        code => code,
                        _ => new ModuleActionFlags
                        {
                            CanView = false,
                            CanCreate = false,
                            CanEdit = false,
                            CanDelete = false
                        },
                        StringComparer.OrdinalIgnoreCase),
                    SpecialActions = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase)
                };

                return emptyResult;
            }

            string roleName = role.RoleName;
            bool isDeveloper = (effectiveRoleId == 1);

            // Parse Role Permissions JSON from ToolsRoles
            var (rolePermsMap, specialActionsMap) = ParseRolePermissionsPayload(role.PermissionsJson);

            // Build effective permission dictionary across all module codes
            var effectiveModules = new Dictionary<string, ModuleActionFlags>(StringComparer.OrdinalIgnoreCase);

            foreach (var code in ModuleCodes.All)
            {
                rolePermsMap.TryGetValue(code, out var roleFlags);

                // Permissions derive exclusively from assigned role definition
                effectiveModules[code] = new ModuleActionFlags
                {
                    CanView = roleFlags?.CanView ?? false,
                    CanCreate = roleFlags?.CanCreate ?? false,
                    CanEdit = roleFlags?.CanEdit ?? false,
                    CanDelete = roleFlags?.CanDelete ?? false
                };
            }

            var result = new UserPermissionsDto
            {
                UserId = userId,
                RoleId = effectiveRoleId,
                RoleName = roleName,
                IsSuperuser = isDeveloper,
                Modules = effectiveModules,
                SpecialActions = specialActionsMap
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

        private static (Dictionary<string, ModuleActionFlags> Modules, Dictionary<string, bool> SpecialActions) ParseRolePermissionsPayload(string? json)
        {
            var modules = new Dictionary<string, ModuleActionFlags>(StringComparer.OrdinalIgnoreCase);
            var special = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);

            if (string.IsNullOrWhiteSpace(json)) return (modules, special);

            try
            {
                using var doc = JsonDocument.Parse(json);
                var root = doc.RootElement;

                if (root.ValueKind == JsonValueKind.Object)
                {
                    // Check if it's new structured format: { "Modules": { ... }, "SpecialActions": { ... } }
                    if (root.TryGetProperty("Modules", out var modProp) || root.TryGetProperty("modules", out modProp))
                    {
                        var opts = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
                        modules = JsonSerializer.Deserialize<Dictionary<string, ModuleActionFlags>>(modProp.GetRawText(), opts)
                                  ?? new Dictionary<string, ModuleActionFlags>(StringComparer.OrdinalIgnoreCase);

                        if (root.TryGetProperty("SpecialActions", out var specProp) || root.TryGetProperty("specialActions", out specProp))
                        {
                            special = JsonSerializer.Deserialize<Dictionary<string, bool>>(specProp.GetRawText(), opts)
                                      ?? new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
                        }
                    }
                    else
                    {
                        // Legacy flat dictionary of module codes
                        var opts = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
                        modules = JsonSerializer.Deserialize<Dictionary<string, ModuleActionFlags>>(json, opts)
                                  ?? new Dictionary<string, ModuleActionFlags>(StringComparer.OrdinalIgnoreCase);
                    }
                }
            }
            catch
            {
                // Return empty safe fallbacks on error
            }

            return (modules, special);
        }

        private static Dictionary<string, ModuleActionFlags> ParsePermissionsMap(string? json)
        {
            var (mods, _) = ParseRolePermissionsPayload(json);
            return mods;
        }

        private async Task SeedDefaultRolesAsync()
        {
            var fullAccessFlags = new ModuleActionFlags { CanView = true, CanCreate = true, CanEdit = true, CanDelete = true };
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
                    CanEdit = !isRoles,
                    CanDelete = !isMaster && !isRoles
                };
            }
            var managerJson = JsonSerializer.Serialize(managerMap);

            var operatorMap = new Dictionary<string, ModuleActionFlags>(StringComparer.OrdinalIgnoreCase);
            foreach (var m in ModuleCodes.All)
            {
                bool canView = m is ModuleCodes.DataImport or ModuleCodes.UploadedData or ModuleCodes.ConflictReport
                                or ModuleCodes.FillMissingData or ModuleCodes.LotBifurcation or ModuleCodes.MergeCatchNumbers
                                or ModuleCodes.ProcessingPipeline or ModuleCodes.ReportTemplateManagement
                                or ModuleCodes.HeaderVerification or ModuleCodes.BatchData or ModuleCodes.Reports;
                operatorMap[m] = new ModuleActionFlags
                {
                    CanView = canView,
                    CanCreate = false,
                    CanEdit = false,
                    CanDelete = false
                };
            }
            var operatorJson = JsonSerializer.Serialize(operatorMap);

            var roles = new[]
            {
                new ToolsRole { Id = 1, RoleName = "Developer", PermissionsJson = fullAccessJson },
                new ToolsRole { Id = 2, RoleName = "Admin", PermissionsJson = fullAccessJson },
                new ToolsRole { Id = 3, RoleName = "Head", PermissionsJson = fullAccessJson },
                new ToolsRole { Id = 4, RoleName = "Manager", PermissionsJson = managerJson },
                new ToolsRole { Id = 5, RoleName = "Operator", PermissionsJson = operatorJson }
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
                    existing.UpdatedAt = DateTime.Now;
                }
            }
            await _context.SaveChangesAsync();
        }
    }
}
