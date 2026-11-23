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
        [RequirePermission(ModuleCodes.RolesPermissions, PermissionAction.View)]
        public async Task<IActionResult> GetRoles()
        {
            await _permissionService.EnsureDatabaseInitializedAsync();

            var assignmentCounts = await _context.UserToolsRoles
                .GroupBy(u => u.RoleId)
                .Select(g => new { RoleId = g.Key, Count = g.Count() })
                .ToDictionaryAsync(g => g.RoleId, g => g.Count);

            var roles = await _context.ToolsRoles
                .AsNoTracking()
                .OrderBy(r => r.Id)
                .Select(r => new
                {
                    r.Id,
                    r.RoleName,
                    r.UpdatedAt,
                    AssignedUsersCount = assignmentCounts.ContainsKey(r.Id) ? assignmentCounts[r.Id] : 0
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

            if (!string.Equals(role.RoleName, request.RoleName, StringComparison.OrdinalIgnoreCase))
            {
                return BadRequest(new { message = "Built-in role names cannot be changed." });
            }

            role.RoleName = request.RoleName.Trim();
            role.UpdatedAt = DateTime.Now;

            await _context.SaveChangesAsync();
            return Ok(role);
        }

        /// <summary>
        /// Delete custom role (cannot delete protected roles or roles with assigned users).
        /// </summary>
        [HttpDelete("roles/{id}")]
        [RequirePermission(ModuleCodes.RolesPermissions, PermissionAction.Edit)]
        public async Task<IActionResult> DeleteRole(int id)
        {
            var role = await _context.ToolsRoles.FindAsync(id);
            if (role == null) return NotFound(new { message = $"Role with ID {id} not found." });

            if (role.RoleName == "Read Only" || id == 1 || id == 2 )
            {
                return BadRequest(new { message = "Protected system roles cannot be deleted." });
            }

            bool hasAssignedUsers = await _context.UserToolsRoles.AnyAsync(u => u.RoleId == id);
            if (hasAssignedUsers)
            {
                return BadRequest(new { message = "Cannot delete a role that is currently assigned to users. Please reassign the users first." });
            }

            _context.ToolsRoles.Remove(role);
            await _context.SaveChangesAsync();

            return NoContent();
        }

        /// <summary>
        /// Get permissions matrix for a role.
        /// </summary>
        [HttpGet("role/{roleId}")]
        [RequirePermission(ModuleCodes.RolesPermissions, PermissionAction.View)]
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
            var role = await _context.ToolsRoles.FindAsync(roleId);
            if (role == null) return NotFound(new { message = $"Role with ID {roleId} not found." });

            // Safeguard Read Only role: Reject any mutating standard permissions or Special Actions
            if (role.RoleName == "Read Only")
            {
                bool hasMutatingPerms = request.Permissions.Any(p => p.CanCreate || p.CanEdit || p.CanDelete)
                    || (request.SpecialActions != null && request.SpecialActions.Values.Any(v => v));
                if (hasMutatingPerms)
                {
                    return BadRequest(new { message = "Mutating permissions and Special Actions cannot be granted to the Read Only role." });
                }
            }

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

        /// <summary>
        /// Get all ERP users via external ERP API: {baseUrl}/User/GetAllUsersExcept
        /// </summary>
        [HttpGet("available-users")]
        [RequirePermission(ModuleCodes.RolesPermissions, PermissionAction.View)]
        public async Task<IActionResult> GetAvailableUsers([FromServices] IConfiguration config)
        {
            string baseUrl = config["ApiSettings:BaseUrl"] ?? "http://192.168.10.208:82/API/api";
            string url = $"{baseUrl}/User/GetAllUsersExcept";

            try
            {
                using var client = new HttpClient { Timeout = TimeSpan.FromSeconds(5) };
                var authHeader = Request.Headers["Authorization"].ToString();
                if (!string.IsNullOrWhiteSpace(authHeader))
                {
                    client.DefaultRequestHeaders.TryAddWithoutValidation("Authorization", authHeader);
                }

                var response = await client.GetAsync(url);
                if (response.IsSuccessStatusCode)
                {
                    var json = await response.Content.ReadAsStringAsync();
                    using var doc = JsonDocument.Parse(json);
                    var root = doc.RootElement;
                    var arr = root.ValueKind == JsonValueKind.Array ? root :
                              root.TryGetProperty("$values", out var vals) ? vals : default;

                    var users = new List<object>();
                    if (arr.ValueKind == JsonValueKind.Array)
                    {
                        foreach (var u in arr.EnumerateArray())
                        {
                            int uId = 0;
                            if (u.TryGetProperty("userId", out var p1) || u.TryGetProperty("UserId", out p1) || u.TryGetProperty("id", out p1))
                                p1.TryGetInt32(out uId);

                            string uName = (u.TryGetProperty("userName", out var p2) || u.TryGetProperty("UserName", out p2)) ? p2.GetString() ?? "" : "";
                            string fName = (u.TryGetProperty("firstName", out var p3) || u.TryGetProperty("FirstName", out p3)) ? p3.GetString() ?? "" : "";
                            string lName = (u.TryGetProperty("lastName", out var p4) || u.TryGetProperty("LastName", out p4)) ? p4.GetString() ?? "" : "";
                            string mobile = (u.TryGetProperty("mobileNo", out var p5) || u.TryGetProperty("MobileNo", out p5)) ? p5.GetString() ?? "" : "";
                            string addr = (u.TryGetProperty("address", out var p6) || u.TryGetProperty("Address", out p6)) ? p6.GetString() ?? "" : "";
                            string pic = (u.TryGetProperty("profilePicturePath", out var p7) || u.TryGetProperty("ProfilePicturePath", out p7)) ? p7.GetString() ?? "" : "";

                            if (uId > 0)
                            {
                                users.Add(new
                                {
                                    UserId = uId,
                                    UserName = uName,
                                    FirstName = fName,
                                    LastName = lName,
                                    FullName = $"{fName} {lName}".Trim(),
                                    MobileNo = mobile,
                                    Address = addr,
                                    ProfilePicturePath = pic
                                });
                            }
                        }
                    }
                    return Ok(users);
                }
                return StatusCode((int)response.StatusCode, new { message = "Failed to fetch users from ERP API." });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { message = "Error contacting ERP User API", error = ex.Message });
            }
        }

        /// <summary>
        /// Get user Tools role assignments combined with ERP users supporting backend search, sorting, filtering, and pagination.
        /// </summary>
        [HttpGet("user-roles")]
        [RequirePermission(ModuleCodes.RolesPermissions, PermissionAction.View)]
        public async Task<IActionResult> GetUserRoles(
            [FromServices] IConfiguration config,
            [FromQuery] string? search = null,
            [FromQuery] string? roleFilter = null,
            [FromQuery] string? sortBy = "userId",
            [FromQuery] string? sortOrder = "ascend",
            [FromQuery] int pageNumber = 1,
            [FromQuery] int pageSize = 15)
        {
            var assignments = await _context.UserToolsRoles
                .AsNoTracking()
                .Include(u => u.Role)
                .ToListAsync();

            var readOnlyRole = await _context.ToolsRoles.AsNoTracking().FirstOrDefaultAsync(r => r.RoleName == "Read Only");
            int readOnlyId = readOnlyRole?.Id ?? 0;
            string readOnlyName = readOnlyRole?.RoleName ?? "Read Only";

            // Fetch ERP Users using {baseUrl}/User/GetAllUsersExcept
            string baseUrl = config["ApiSettings:BaseUrl"] ?? "http://192.168.10.208:82/API/api";
            string url = $"{baseUrl}/User/GetAllUsersExcept";
            var erpUsers = new List<(int UserId, string UserName, string FirstName, string LastName, string MobileNo, string Address, string ProfilePicturePath)>();

            try
            {
                using var client = new HttpClient { Timeout = TimeSpan.FromSeconds(5) };
                var authHeader = Request.Headers["Authorization"].ToString();
                if (!string.IsNullOrWhiteSpace(authHeader))
                {
                    client.DefaultRequestHeaders.TryAddWithoutValidation("Authorization", authHeader);
                }

                var response = await client.GetAsync(url);
                if (response.IsSuccessStatusCode)
                {
                    var json = await response.Content.ReadAsStringAsync();
                    using var doc = JsonDocument.Parse(json);
                    var root = doc.RootElement;
                    var arr = root.ValueKind == JsonValueKind.Array ? root :
                              root.TryGetProperty("$values", out var vals) ? vals : default;

                    if (arr.ValueKind == JsonValueKind.Array)
                    {
                        foreach (var u in arr.EnumerateArray())
                        {
                            int uId = 0;
                            if (u.TryGetProperty("userId", out var p1) || u.TryGetProperty("UserId", out p1) || u.TryGetProperty("id", out p1))
                                p1.TryGetInt32(out uId);

                            string uName = (u.TryGetProperty("userName", out var p2) || u.TryGetProperty("UserName", out p2)) ? p2.GetString() ?? "" : "";
                            string fName = (u.TryGetProperty("firstName", out var p3) || u.TryGetProperty("FirstName", out p3)) ? p3.GetString() ?? "" : "";
                            string lName = (u.TryGetProperty("lastName", out var p4) || u.TryGetProperty("LastName", out p4)) ? p4.GetString() ?? "" : "";
                            string mobile = (u.TryGetProperty("mobileNo", out var p5) || u.TryGetProperty("MobileNo", out p5)) ? p5.GetString() ?? "" : "";
                            string addr = (u.TryGetProperty("address", out var p6) || u.TryGetProperty("Address", out p6)) ? p6.GetString() ?? "" : "";
                            string pic = (u.TryGetProperty("profilePicturePath", out var p7) || u.TryGetProperty("ProfilePicturePath", out p7)) ? p7.GetString() ?? "" : "";

                            if (uId > 0) erpUsers.Add((uId, uName, fName, lName, mobile, addr, pic));
                        }
                    }
                }
            }
            catch { }

            var assignmentMap = assignments.ToDictionary(a => a.UserId, a => a);
            var results = new List<UserRoleAssignmentDto>();

            // 1. Process all known ERP users
            var processedUserIds = new HashSet<int>();
            foreach (var eu in erpUsers)
            {
                processedUserIds.Add(eu.UserId);
                if (assignmentMap.TryGetValue(eu.UserId, out var a))
                {
                    results.Add(new UserRoleAssignmentDto
                    {
                        UserId = eu.UserId,
                        UserName = eu.UserName,
                        FirstName = eu.FirstName,
                        LastName = eu.LastName,
                        MobileNo = eu.MobileNo,
                        Address = eu.Address,
                        ProfilePicturePath = eu.ProfilePicturePath,
                        RoleId = a.RoleId,
                        RoleName = a.Role?.RoleName ?? $"Role {a.RoleId}",
                        IsExplicit = true
                    });
                }
                else
                {
                    results.Add(new UserRoleAssignmentDto
                    {
                        UserId = eu.UserId,
                        UserName = eu.UserName,
                        FirstName = eu.FirstName,
                        LastName = eu.LastName,
                        MobileNo = eu.MobileNo,
                        Address = eu.Address,
                        ProfilePicturePath = eu.ProfilePicturePath,
                        RoleId = readOnlyId,
                        RoleName = readOnlyName,
                        IsExplicit = false
                    });
                }
            }

            // 2. Include any assigned users not in the ERP list
            foreach (var a in assignments)
            {
                if (!processedUserIds.Contains(a.UserId))
                {
                    results.Add(new UserRoleAssignmentDto
                    {
                        UserId = a.UserId,
                        UserName = $"User {a.UserId}",
                        FirstName = $"User",
                        LastName = $"{a.UserId}",
                        RoleId = a.RoleId,
                        RoleName = a.Role?.RoleName ?? $"Role {a.RoleId}",
                        IsExplicit = true
                    });
                }
            }

            // 3. Backend Filtering: Search
            IEnumerable<UserRoleAssignmentDto> query = results;
            if (!string.IsNullOrWhiteSpace(search))
            {
                string s = search.Trim();
                query = query.Where(u =>
                    u.UserId.ToString().Contains(s, StringComparison.OrdinalIgnoreCase) ||
                    (!string.IsNullOrEmpty(u.UserName) && u.UserName.Contains(s, StringComparison.OrdinalIgnoreCase)) ||
                    (!string.IsNullOrEmpty(u.FirstName) && u.FirstName.Contains(s, StringComparison.OrdinalIgnoreCase)) ||
                    (!string.IsNullOrEmpty(u.LastName) && u.LastName.Contains(s, StringComparison.OrdinalIgnoreCase)) ||
                    (!string.IsNullOrEmpty(u.MobileNo) && u.MobileNo.Contains(s, StringComparison.OrdinalIgnoreCase)) ||
                    (!string.IsNullOrEmpty(u.Address) && u.Address.Contains(s, StringComparison.OrdinalIgnoreCase)) ||
                    (!string.IsNullOrEmpty(u.RoleName) && u.RoleName.Contains(s, StringComparison.OrdinalIgnoreCase))
                );
            }

            // 4. Backend Filtering: Role Filter
            if (!string.IsNullOrWhiteSpace(roleFilter) && !string.Equals(roleFilter, "all", StringComparison.OrdinalIgnoreCase))
            {
                if (string.Equals(roleFilter, "explicit", StringComparison.OrdinalIgnoreCase))
                {
                    query = query.Where(u => u.IsExplicit);
                }
                else if (string.Equals(roleFilter, "default", StringComparison.OrdinalIgnoreCase))
                {
                    query = query.Where(u => !u.IsExplicit);
                }
                else if (int.TryParse(roleFilter, out int filterRoleId))
                {
                    query = query.Where(u => u.RoleId == filterRoleId);
                }
                else
                {
                    query = query.Where(u => string.Equals(u.RoleName, roleFilter.Trim(), StringComparison.OrdinalIgnoreCase));
                }
            }

            // 5. Backend Sorting
            bool isDesc = string.Equals(sortOrder, "descend", StringComparison.OrdinalIgnoreCase) || string.Equals(sortOrder, "desc", StringComparison.OrdinalIgnoreCase);
            query = (sortBy?.ToLower()) switch
            {
                "username" or "name" => isDesc ? query.OrderByDescending(u => $"{u.FirstName} {u.LastName}".Trim()) : query.OrderBy(u => $"{u.FirstName} {u.LastName}".Trim()),
                "rolename" or "role" => isDesc ? query.OrderByDescending(u => u.RoleName).ThenBy(u => u.UserId) : query.OrderBy(u => u.RoleName).ThenBy(u => u.UserId),
                "isexplicit" or "status" => isDesc ? query.OrderByDescending(u => u.IsExplicit).ThenBy(u => u.UserId) : query.OrderBy(u => u.IsExplicit).ThenBy(u => u.UserId),
                "contact" or "mobileno" => isDesc ? query.OrderByDescending(u => u.MobileNo).ThenBy(u => u.UserId) : query.OrderBy(u => u.MobileNo).ThenBy(u => u.UserId),
                _ => isDesc ? query.OrderByDescending(u => u.UserId) : query.OrderBy(u => u.UserId)
            };

            int totalCount = query.Count();

            // 6. Backend Pagination (Safe defaults)
            if (pageNumber < 1) pageNumber = 1;
            if (pageSize < 1) pageSize = 15;

            var pagedItems = query
                .Skip((pageNumber - 1) * pageSize)
                .Take(pageSize)
                .ToList();

            var responseDto = new UserRolesPagedResponse
            {
                Items = pagedItems,
                TotalCount = totalCount,
                PageNumber = pageNumber,
                PageSize = pageSize
            };

            return Ok(responseDto);
        }

        /// <summary>
        /// Assign or change Tools role for a single user or multiple users at a time.
        /// </summary>
        [HttpPost("user-roles")]
        [RequirePermission(ModuleCodes.RolesPermissions, PermissionAction.Edit)]
        public async Task<IActionResult> AssignUserRole([FromBody] AssignUserRoleRequest request)
        {
            if (request.UserId <= 0) return BadRequest(new { message = "Valid UserId is required." });
            if (request.RoleId <= 0) return BadRequest(new { message = "Valid RoleId is required." });

            var roleExists = await _context.ToolsRoles.AnyAsync(r => r.Id == request.RoleId);
            if (!roleExists) return NotFound(new { message = $"Tools Role with ID {request.RoleId} not found." });

            var existing = await _context.UserToolsRoles.FirstOrDefaultAsync(u => u.UserId == request.UserId);
            if (existing != null)
            {
                existing.RoleId = request.RoleId;
            }
            else
            {
                _context.UserToolsRoles.Add(new UserToolsRole
                {
                    UserId = request.UserId,
                    RoleId = request.RoleId
                });
            }

            await _context.SaveChangesAsync();
            _permissionService.InvalidateUserCache(request.UserId);

            int actorId = LogHelper.GetTriggeredBy(User, Request);
            await _loggerService.LogEventAsync($"Assigned Tools Role {request.RoleId} to User {request.UserId}", "RolesPermissions", actorId, 0);

            return Ok(new { success = true, message = $"Tools Role {request.RoleId} assigned successfully to User {request.UserId}." });
        }

        /// <summary>
        /// Bulk assign Tools role to multiple users at a time.
        /// </summary>
        [HttpPost("user-roles/bulk")]
        [RequirePermission(ModuleCodes.RolesPermissions, PermissionAction.Edit)]
        public async Task<IActionResult> BulkAssignUserRoles([FromBody] BulkAssignUserRolesRequest request)
        {
            if (request.UserIds == null || !request.UserIds.Any())
                return BadRequest(new { message = "At least one UserId is required." });

            var validUserIds = request.UserIds.Where(id => id > 0).Distinct().ToList();
            if (!validUserIds.Any())
                return BadRequest(new { message = "No valid UserIds provided." });

            if (request.RoleId <= 0)
                return BadRequest(new { message = "Valid RoleId is required." });

            var role = await _context.ToolsRoles.FirstOrDefaultAsync(r => r.Id == request.RoleId);
            if (role == null)
                return NotFound(new { message = $"Tools Role with ID {request.RoleId} not found." });

            var existingAssignments = await _context.UserToolsRoles
                .Where(u => validUserIds.Contains(u.UserId))
                .ToListAsync();

            var existingDict = existingAssignments.ToDictionary(u => u.UserId);

            foreach (var userId in validUserIds)
            {
                if (existingDict.TryGetValue(userId, out var existing))
                {
                    existing.RoleId = request.RoleId;
                }
                else
                {
                    _context.UserToolsRoles.Add(new UserToolsRole
                    {
                        UserId = userId,
                        RoleId = request.RoleId
                    });
                }
            }

            await _context.SaveChangesAsync();

            foreach (var userId in validUserIds)
            {
                _permissionService.InvalidateUserCache(userId);
            }

            int actorId = LogHelper.GetTriggeredBy(User, Request);
            await _loggerService.LogEventAsync($"Bulk assigned Tools Role {request.RoleId} ({role.RoleName}) to {validUserIds.Count} user(s)", "RolesPermissions", actorId, 0);

            return Ok(new
            {
                success = true,
                message = $"Successfully assigned role '{role.RoleName}' to {validUserIds.Count} user(s).",
                count = validUserIds.Count
            });
        }

        /// <summary>
        /// Remove explicit Tools role assignment (user falls back to Read Only).
        /// </summary>
        [HttpDelete("user-roles/{userId}")]
        [RequirePermission(ModuleCodes.RolesPermissions, PermissionAction.Edit)]
        public async Task<IActionResult> RemoveUserRole(int userId)
        {
            var assignment = await _context.UserToolsRoles.FirstOrDefaultAsync(u => u.UserId == userId);
            if (assignment == null)
            {
                return NotFound(new { message = $"No explicit role assignment found for User ID {userId}." });
            }

            _context.UserToolsRoles.Remove(assignment);
            await _context.SaveChangesAsync();
            _permissionService.InvalidateUserCache(userId);

            int actorId = LogHelper.GetTriggeredBy(User, Request);
            await _loggerService.LogEventAsync($"Removed explicit Tools Role assignment for User {userId}; user reverted to Read Only.", "RolesPermissions", actorId, 0);

            return Ok(new { success = true, message = $"Explicit role assignment removed for User {userId}. User has reverted to default Read Only role." });
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
            ModuleCodes.MasterProject => new[] { "View", "Create", "Edit" , "Delete"},
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
