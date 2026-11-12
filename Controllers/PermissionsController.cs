using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using ERPToolsAPI.Data;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Tools.Models;
using Tools.Services;
using Tools.Middleware;

namespace Tools.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class PermissionsController : ControllerBase
    {
        private readonly ERPToolsDbContext _context;
        private readonly IPermissionService _permissionService;
        private readonly ILoggerService _loggerService;

        public PermissionsController(
            ERPToolsDbContext context,
            IPermissionService permissionService,
            ILoggerService loggerService)
        {
            _context = context;
            _permissionService = permissionService;
            _loggerService = loggerService;
        }

        /// <summary>
        /// Get current logged-in user's computed permissions.
        /// </summary>
        [HttpGet("my-permissions")]
        public async Task<IActionResult> GetMyPermissions()
        {
            int userId = LogHelper.GetTriggeredBy(User, Request);
            if (userId <= 0)
            {
                return Unauthorized(new { message = "User not authenticated." });
            }

            var result = await _permissionService.GetUserPermissionsAsync(userId, Request);
            return Ok(result);
        }

        /// <summary>
        /// Get supported system modules.
        /// </summary>
        [HttpGet("modules")]
        public IActionResult GetModules()
        {
            var list = ModuleCodes.All.Select(code => new
            {
                Code = code,
                DisplayName = FormatDisplayName(code),
                SupportedActions = new[] { "View", "Create", "Edit", "Delete", "Execute", "Export" }
            });

            return Ok(list);
        }

        /// <summary>
        /// Get all roles.
        /// </summary>
        [HttpGet("roles")]
        public async Task<IActionResult> GetRoles()
        {
            await _permissionService.EnsureDatabaseInitializedAsync();

            var roles = await _context.ToolsRoles
                .AsNoTracking()
                .OrderBy(r => r.Id)
                .Select(r => new
                {
                    r.Id,
                    r.RoleName,
                    r.Description,
                    r.IsSystemRole,
                    r.UpdatedAt,
                    UserCount = _context.UserPermissions.Count(u => u.RoleId == r.Id)
                })
                .ToListAsync();

            return Ok(roles);
        }

        /// <summary>
        /// Create a custom role.
        /// </summary>
        [HttpPost("roles")]
        [RequirePermission(ModuleCodes.RolesPermissions, PermissionAction.Edit)]
        public async Task<IActionResult> CreateRole([FromBody] CreateRoleRequest request)
        {
            if (string.IsNullOrWhiteSpace(request.RoleName))
            {
                return BadRequest(new { message = "RoleName is required." });
            }

            int userId = LogHelper.GetTriggeredBy(User, Request);
            int roleId = LogHelper.GetUserRoleId(User, Request);
            if (roleId != 1 && roleId != 2)
            {
                return StatusCode(403, new { message = "Only administrators can create custom roles." });
            }

            bool exists = await _context.ToolsRoles.AnyAsync(r => r.RoleName == request.RoleName.Trim());
            if (exists)
            {
                return Conflict(new { message = $"Role '{request.RoleName}' already exists." });
            }

            // Default empty permissions for new role
            var defaultPerms = ModuleCodes.All.ToDictionary(m => m, _ => new ModuleActionFlags
            {
                CanView = false, CanCreate = false, CanEdit = false, CanDelete = false, CanExecute = false, CanExport = false
            });

            var newRole = new ToolsRole
            {
                RoleName = request.RoleName.Trim(),
                Description = request.Description,
                IsSystemRole = false,
                PermissionsJson = JsonSerializer.Serialize(defaultPerms),
                UpdatedAt = DateTime.Now
            };

            _context.ToolsRoles.Add(newRole);
            await _context.SaveChangesAsync();

            await _loggerService.LogEventAsync($"Created custom role '{newRole.RoleName}'", "RolesPermissions", userId, 0);
            return Ok(newRole);
        }

        /// <summary>
        /// Update custom role name or description.
        /// </summary>
        [HttpPut("roles/{id}")]
        [RequirePermission(ModuleCodes.RolesPermissions, PermissionAction.Edit)]
        public async Task<IActionResult> UpdateRole(int id, [FromBody] UpdateRoleRequest request)
        {
            int roleId = LogHelper.GetUserRoleId(User, Request);
            if (roleId != 1 && roleId != 2)
            {
                return StatusCode(403, new { message = "Only administrators can modify roles." });
            }

            var role = await _context.ToolsRoles.FindAsync(id);
            if (role == null) return NotFound(new { message = $"Role with ID {id} not found." });

            if (role.IsSystemRole && !string.Equals(role.RoleName, request.RoleName, StringComparison.OrdinalIgnoreCase))
            {
                return BadRequest(new { message = "Built-in role names cannot be changed." });
            }

            role.RoleName = request.RoleName.Trim();
            role.Description = request.Description;
            role.UpdatedAt = DateTime.Now;

            await _context.SaveChangesAsync();
            return Ok(role);
        }

        /// <summary>
        /// Delete custom role (cannot delete system roles).
        /// </summary>
        [HttpDelete("roles/{id}")]
        [RequirePermission(ModuleCodes.RolesPermissions, PermissionAction.Edit)]
        public async Task<IActionResult> DeleteRole(int id)
        {
            int roleId = LogHelper.GetUserRoleId(User, Request);
            if (roleId != 1 && roleId != 2)
            {
                return StatusCode(403, new { message = "Only administrators can delete roles." });
            }

            var role = await _context.ToolsRoles.FindAsync(id);
            if (role == null) return NotFound(new { message = $"Role with ID {id} not found." });
            if (role.IsSystemRole) return BadRequest(new { message = "Built-in system roles cannot be deleted." });

            _context.ToolsRoles.Remove(role);
            await _context.SaveChangesAsync();

            return NoContent();
        }

        /// <summary>
        /// Get permissions matrix for a role.
        /// </summary>
        [HttpGet("role/{roleId}")]
        public async Task<IActionResult> GetRolePermissions(int roleId)
        {
            await _permissionService.EnsureDatabaseInitializedAsync();

            var role = await _context.ToolsRoles.AsNoTracking().FirstOrDefaultAsync(r => r.Id == roleId);
            if (role == null) return NotFound(new { message = $"Role with ID {roleId} not found." });

            var permsMap = ParseMap(role.PermissionsJson);
            var list = ModuleCodes.All.Select(code =>
            {
                permsMap.TryGetValue(code, out var f);
                return new RolePermissionItemDto
                {
                    ModuleCode = code,
                    CanView = f?.CanView ?? false,
                    CanCreate = f?.CanCreate ?? false,
                    CanEdit = f?.CanEdit ?? false,
                    CanDelete = f?.CanDelete ?? false,
                    CanExecute = f?.CanExecute ?? false,
                    CanExport = f?.CanExport ?? false
                };
            }).ToList();

            return Ok(new { Role = role, Permissions = list });
        }

        /// <summary>
        /// Update permissions matrix for a role.
        /// </summary>
        [HttpPut("role/{roleId}")]
        [RequirePermission(ModuleCodes.RolesPermissions, PermissionAction.Edit)]
        public async Task<IActionResult> UpdateRolePermissions(int roleId, [FromBody] UpdateRolePermissionsRequest request)
        {
            int callerRoleId = LogHelper.GetUserRoleId(User, Request);
            if (callerRoleId != 1 && callerRoleId != 2)
            {
                return StatusCode(403, new { message = "Only administrators can update role permissions." });
            }

            var role = await _context.ToolsRoles.FindAsync(roleId);
            if (role == null) return NotFound(new { message = $"Role with ID {roleId} not found." });

            var map = new Dictionary<string, ModuleActionFlags>(StringComparer.OrdinalIgnoreCase);
            foreach (var item in request.Permissions)
            {
                map[item.ModuleCode] = new ModuleActionFlags
                {
                    CanView = item.CanView,
                    CanCreate = item.CanCreate,
                    CanEdit = item.CanEdit,
                    CanDelete = item.CanDelete,
                    CanExecute = item.CanExecute,
                    CanExport = item.CanExport
                };
            }

            if (role.Id == 1 || string.Equals(role.RoleName, "Developer", StringComparison.OrdinalIgnoreCase))
            {
                // Administrative fail-safe: Developer always retains full access to Roles & Permissions Admin
                map[ModuleCodes.RolesPermissions] = new ModuleActionFlags
                {
                    CanView = true,
                    CanCreate = true,
                    CanEdit = true,
                    CanDelete = true,
                    CanExecute = true,
                    CanExport = true
                };
            }

            role.PermissionsJson = JsonSerializer.Serialize(map);
            role.UpdatedAt = DateTime.Now;

            await _context.SaveChangesAsync();
            _permissionService.InvalidateAllCaches();

            return Ok(new { success = true, message = $"Permissions updated for role '{role.RoleName}'." });
        }

        /// <summary>
        /// Get user's assigned role and overrides.
        /// </summary>
        [HttpGet("user/{userId}")]
        public async Task<IActionResult> GetUserConfig(int userId, [FromQuery] int? erpRoleId = null)
        {
            await _permissionService.EnsureDatabaseInitializedAsync();
            _permissionService.InvalidateUserCache(userId);

            var userRecord = await _context.UserPermissions.AsNoTracking().FirstOrDefaultAsync(u => u.UserId == userId);
            ToolsRole? assignedRole = null;
            if (userRecord?.RoleId != null)
            {
                assignedRole = await _context.ToolsRoles.AsNoTracking().FirstOrDefaultAsync(r => r.Id == userRecord.RoleId.Value);
            }

            var overridesMap = ParseMap(userRecord?.OverridesJson);
            var overridesList = ModuleCodes.All.Select(code =>
            {
                overridesMap.TryGetValue(code, out var f);
                return new UserOverrideItemDto
                {
                    ModuleCode = code,
                    CanView = f?.CanView,
                    CanCreate = f?.CanCreate,
                    CanEdit = f?.CanEdit,
                    CanDelete = f?.CanDelete,
                    CanExecute = f?.CanExecute,
                    CanExport = f?.CanExport
                };
            }).ToList();

            var effective = await _permissionService.GetUserPermissionsAsync(userId, Request, erpRoleId);

            return Ok(new
            {
                UserId = userId,
                AssignedRole = assignedRole != null ? new { RoleId = assignedRole.Id, RoleName = assignedRole.RoleName } : null,
                EffectiveRole = new { effective.RoleId, effective.RoleName, effective.IsSuperuser },
                Overrides = overridesList,
                EffectivePermissions = effective.Modules
            });
        }

        /// <summary>
        /// Assign role to user or update user overrides.
        /// </summary>
        [HttpPut("user/{userId}")]
        [RequirePermission(ModuleCodes.RolesPermissions, PermissionAction.Edit)]
        public async Task<IActionResult> UpdateUserConfig(int userId, [FromBody] UpdateUserPermissionsRequest request)
        {
            int callerUserId = LogHelper.GetTriggeredBy(User, Request);
            int callerRoleId = LogHelper.GetUserRoleId(User, Request);
            if (callerRoleId != 1 && callerRoleId != 2)
            {
                return StatusCode(403, new { message = "Only administrators can configure user permissions." });
            }

            var userRecord = await _context.UserPermissions.FirstOrDefaultAsync(u => u.UserId == userId);

            // Filter overrides: only store modules with at least one non-null override
            var overridesMap = new Dictionary<string, ModuleActionFlags>(StringComparer.OrdinalIgnoreCase);
            if (request.Overrides != null)
            {
                foreach (var item in request.Overrides)
                {
                    bool hasValue = item.CanView.HasValue || item.CanCreate.HasValue || item.CanEdit.HasValue
                                    || item.CanDelete.HasValue || item.CanExecute.HasValue || item.CanExport.HasValue;
                    if (hasValue)
                    {
                        overridesMap[item.ModuleCode] = new ModuleActionFlags
                        {
                            CanView = item.CanView,
                            CanCreate = item.CanCreate,
                            CanEdit = item.CanEdit,
                            CanDelete = item.CanDelete,
                            CanExecute = item.CanExecute,
                            CanExport = item.CanExport
                        };
                    }
                }
            }

            string? overridesJson = overridesMap.Count > 0 ? JsonSerializer.Serialize(overridesMap) : null;

            if (userRecord != null)
            {
                userRecord.RoleId = request.RoleId;
                userRecord.OverridesJson = overridesJson;
                userRecord.UpdatedAt = DateTime.Now;
                userRecord.UpdatedBy = callerUserId;
            }
            else
            {
                _context.UserPermissions.Add(new UserPermission
                {
                    UserId = userId,
                    RoleId = request.RoleId,
                    OverridesJson = overridesJson,
                    UpdatedAt = DateTime.Now,
                    UpdatedBy = callerUserId
                });
            }

            await _context.SaveChangesAsync();
            _permissionService.InvalidateUserCache(userId);

            return Ok(new { success = true, message = $"User {userId} permissions configuration saved." });
        }

        private static Dictionary<string, ModuleActionFlags> ParseMap(string? json)
        {
            if (string.IsNullOrWhiteSpace(json)) return new Dictionary<string, ModuleActionFlags>(StringComparer.OrdinalIgnoreCase);
            try
            {
                var opts = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
                return JsonSerializer.Deserialize<Dictionary<string, ModuleActionFlags>>(json, opts)
                    ?? new Dictionary<string, ModuleActionFlags>(StringComparer.OrdinalIgnoreCase);
            }
            catch
            {
                return new Dictionary<string, ModuleActionFlags>(StringComparer.OrdinalIgnoreCase);
            }
        }

        private static string FormatDisplayName(string code) => code switch
        {
            ModuleCodes.Dashboard => "Dashboard",
            ModuleCodes.MasterProject => "Project Master",
            ModuleCodes.MasterConfig => "Master Configuration",
            ModuleCodes.MasterBoxCapacity => "Box Capacity",
            ModuleCodes.MasterEnvelope => "Envelope Type",
            ModuleCodes.MasterFields => "Fields Master",
            ModuleCodes.MasterModules => "Tools Module",
            ModuleCodes.MasterNodalExtra => "Nodal Univ Extra",
            ModuleCodes.MasterRpt => "RPT Templates",
            ModuleCodes.MasterCrystal => "Crystal Reports",
            ModuleCodes.ProjectConfig => "Project Configuration",
            ModuleCodes.DataImport => "Data Import",
            ModuleCodes.ProcessingPipeline => "Processing Pipeline",
            ModuleCodes.HeaderVerification => "Header Verification",
            ModuleCodes.ChangedNr => "Changed NR Analysis",
            ModuleCodes.BatchData => "Batch Wise Data",
            ModuleCodes.NodalCatch => "Nodal & Catch Lists",
            ModuleCodes.HtovTool => "Horizontal To Vertical",
            ModuleCodes.CorrectionTool => "Correction Tool",
            ModuleCodes.Reports => "Reports",
            ModuleCodes.RolesPermissions => "Roles & Permissions Admin",
            _ => code
        };
    }
}
