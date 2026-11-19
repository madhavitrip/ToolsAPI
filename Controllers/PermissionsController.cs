using Microsoft.AspNetCore.Authorization;
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
    [Authorize]
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
        /// Get supported system modules with their actual page-specific supported standard actions.
        /// </summary>
        [HttpGet("modules")]
        public IActionResult GetModules()
        {
            var list = ModuleCodes.All
                .Where(code => !string.Equals(code, ModuleCodes.Dashboard, StringComparison.OrdinalIgnoreCase))
                .Select(code => new
                {
                    Code = code,
                    DisplayName = FormatDisplayName(code),
                    SupportedActions = GetSupportedActionsForModule(code)
                });

            return Ok(list);
        }

        /// <summary>
        /// Get supported special actions metadata.
        /// </summary>
        [HttpGet("special-actions")]
        public IActionResult GetSpecialActions()
        {
            var list = new[]
            {
                new SpecialActionMetadata
                {
                    Code = SpecialActionCodes.DeleteAllNrData,
                    DisplayName = "Delete All NR Data",
                    ParentModule = ModuleCodes.UploadedData,
                    RequiredStandardAction = PermissionAction.Delete,
                    Description = "Allows purging the entire project or lot NR data from the system."
                },
                new SpecialActionMetadata
                {
                    Code = SpecialActionCodes.BulkVerifyCatches,
                    DisplayName = "Verify All Catches",
                    ParentModule = ModuleCodes.HeaderVerification,
                    RequiredStandardAction = PermissionAction.Edit,
                    Description = "Allows bulk verification of all catch records across the project in one action."
                },
                new SpecialActionMetadata
                {
                    Code = SpecialActionCodes.ApplyChangedNr,
                    DisplayName = "Apply Changed NR Comparisons",
                    ParentModule = ModuleCodes.ChangedNr,
                    RequiredStandardAction = PermissionAction.Edit,
                    Description = "Allows applying differential changes from Changed NR analysis into active project data."
                },
                new SpecialActionMetadata
                {
                    Code = SpecialActionCodes.PushNodalToNr,
                    DisplayName = "Push Temp into Main NR Data",
                    ParentModule = ModuleCodes.NodalCatch,
                    RequiredStandardAction = PermissionAction.Edit,
                    Description = "Allows merging verified Nodal center and catch assignments directly into primary NR records."
                },
                new SpecialActionMetadata
                {
                    Code = SpecialActionCodes.PurgeReportRecord,
                    DisplayName = "Delete Report Records",
                    ParentModule = ModuleCodes.ReportTemplateManagement,
                    RequiredStandardAction = PermissionAction.Delete,
                    Description = "Allows deleting generated report records and archiving history from the repository."
                },
                new SpecialActionMetadata
                {
                    Code = SpecialActionCodes.UploadRptTemplate,
                    DisplayName = "Upload RPT Template",
                    ParentModule = ModuleCodes.ReportTemplateManagement,
                    RequiredStandardAction = PermissionAction.Create,
                    Description = "Allows uploading Crystal Report (.rpt) template files into master and project repositories."
                },
                new SpecialActionMetadata
                {
                    Code = SpecialActionCodes.MapRptTemplate,
                    DisplayName = "Configure Template Mapping",
                    ParentModule = ModuleCodes.ReportTemplateManagement,
                    RequiredStandardAction = PermissionAction.Edit,
                    Description = "Allows configuring database field mappings and table relationships for RPT templates."
                },
                new SpecialActionMetadata
                {
                    Code = SpecialActionCodes.SendToErp,
                    DisplayName = "Transmit Data to ERP",
                    ParentModule = ModuleCodes.ProcessingPipeline,
                    RequiredStandardAction = PermissionAction.Edit,
                    Description = "Allows transmitting compiled lot catch records, quantities, envelopes, and pages directly to ERP."
                }
            };

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
                    r.UpdatedAt
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
                CanView = false, CanCreate = false, CanEdit = false, CanDelete = false
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

            var (permsMap, specialMap) = ParseRolePermissions(role.PermissionsJson);
            var list = ModuleCodes.All
                .Where(code => !string.Equals(code, ModuleCodes.Dashboard, StringComparison.OrdinalIgnoreCase))
                .Select(code =>
            {
                permsMap.TryGetValue(code, out var f);
                return new RolePermissionItemDto
                {
                    ModuleCode = code,
                    CanView = f?.CanView ?? false,
                    CanCreate = f?.CanCreate ?? false,
                    CanEdit = f?.CanEdit ?? false,
                    CanDelete = f?.CanDelete ?? false
                };
            }).ToList();

            return Ok(new { Role = role, Permissions = list, SpecialActions = specialMap });
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
                    CanDelete = item.CanDelete
                };
            }

            var specialMap = request.SpecialActions ?? new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);

            var payload = new
            {
                Modules = map,
                SpecialActions = specialMap
            };

            role.PermissionsJson = JsonSerializer.Serialize(payload);
            role.UpdatedAt = DateTime.Now;

            await _context.SaveChangesAsync();
            _permissionService.InvalidateAllCaches();

            return Ok(new { success = true, message = $"Permissions updated for role '{role.RoleName}'." });
        }

        private static (Dictionary<string, ModuleActionFlags> Modules, Dictionary<string, bool> SpecialActions) ParseRolePermissions(string? json)
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
                        var opts = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
                        modules = JsonSerializer.Deserialize<Dictionary<string, ModuleActionFlags>>(json, opts)
                                  ?? new Dictionary<string, ModuleActionFlags>(StringComparer.OrdinalIgnoreCase);
                    }
                }
            }
            catch { }

            return (modules, special);
        }

        private static string[] GetSupportedActionsForModule(string code) => code switch
        {
            ModuleCodes.MasterProject => new[] { "View", "Create", "Edit" },
            ModuleCodes.MasterConfig => new[] { "View","Edit" },
            ModuleCodes.MasterBoxCapacity => new[] { "View", "Create", "Edit", "Delete" },
            ModuleCodes.MasterEnvelope => new[] { "View", "Create", "Edit", "Delete" },
            ModuleCodes.MasterFields => new[] { "View", "Create", "Edit", "Delete" },
            ModuleCodes.MasterModules => new[] { "View", "Create", "Edit", "Delete" },
            ModuleCodes.MasterNodalExtra => new[] { "View", "Create", "Edit", "Delete" },
            ModuleCodes.MasterRpt => new[] { "View", "Create", "Edit", "Delete" },
            ModuleCodes.MessageManagement => new[] { "View", "Create", "Edit" },
            ModuleCodes.RolesPermissions => new[] { "View", "Create", "Edit", "Delete" },
            ModuleCodes.ProjectConfig => new[] { "View", "Edit", "Delete" },
            ModuleCodes.DataImport => new[] { "View", "Create" },
            ModuleCodes.UploadedData => new[] { "View", "Edit", "Delete" },
            ModuleCodes.ConflictReport => new[] { "View", "Edit" },
            ModuleCodes.FillMissingData => new[] { "View", "Edit" },
            ModuleCodes.LotBifurcation => new[] { "View", "Edit" },
            ModuleCodes.MergeCatchNumbers => new[] { "View", "Edit" },
            ModuleCodes.ProcessingPipeline => new[] { "View", "Edit" },
            ModuleCodes.ReportTemplateManagement => new[] { "View", "Edit", "Delete" },
            ModuleCodes.HeaderVerification => new[] { "View", "Edit" },
            ModuleCodes.ChangedNr => new[] { "View", "Create", "Edit" },
            ModuleCodes.BatchData => new[] { "View" },
            ModuleCodes.NodalCatch => new[] { "View", "Create", "Edit", "Delete" },
            ModuleCodes.HtovTool => new[] { "View", "Edit" },
            ModuleCodes.CorrectionTool => new[] { "View", "Edit" },
            ModuleCodes.Reports => new[] { "View" },
            _ => new[] { "View", "Create", "Edit", "Delete" }
        };

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
            ModuleCodes.ProjectConfig => "Project Configuration",
            ModuleCodes.DataImport => "Data Import (Upload)",
            ModuleCodes.UploadedData => "Data Import - Uploaded Data",
            ModuleCodes.ConflictReport => "Data Import - Conflict Report",
            ModuleCodes.FillMissingData => "Data Import - Fill Missing Data",
            ModuleCodes.LotBifurcation => "Data Import - Lot Bifurcation",
            ModuleCodes.MergeCatchNumbers => "Data Import - Merge Catch Numbers",
            ModuleCodes.ProcessingPipeline => "Processing Pipeline",
            ModuleCodes.ReportTemplateManagement => "Report & Template Management",
            ModuleCodes.HeaderVerification => "Header Verification",
            ModuleCodes.ChangedNr => "Changed NR Analysis",
            ModuleCodes.BatchData => "Batch Wise Data",
            ModuleCodes.NodalCatch => "Nodal & Catch Lists",
            ModuleCodes.HtovTool => "Horizontal To Vertical",
            ModuleCodes.CorrectionTool => "Correction Tool",
            ModuleCodes.Reports => "Reports",
            ModuleCodes.RolesPermissions => "Roles & Permissions",
            ModuleCodes.MessageManagement => "Message Management",
            _ => code
        };
    }
}
