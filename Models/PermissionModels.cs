using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Tools.Models
{
    // ==========================================
    // 1. DATABASE ENTITIES (Only 2 Tables)
    // ==========================================

    [Table("ToolsRoles")]
    public class ToolsRole
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }

        [Required]
        [MaxLength(50)]
        public string RoleName { get; set; } = string.Empty;

        [MaxLength(255)]
        public string? Description { get; set; }

        public bool IsSystemRole { get; set; } = false;

        /// <summary>
        /// JSON map: { "MODULE_CODE": { "view": true, "create": true, "edit": true, "delete": true, "execute": true, "export": true } }
        /// </summary>
        public string PermissionsJson { get; set; } = "{}";

        public DateTime UpdatedAt { get; set; } = DateTime.Now;
    }

    [Table("UserPermissions")]
    public class UserPermission
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }

        [Required]
        public int UserId { get; set; }

        /// <summary>
        /// Optional local Tools Role ID. If null, user inherits their ERP JWT roleId.
        /// </summary>
        public int? RoleId { get; set; }

        /// <summary>
        /// JSON map of user overrides: { "MODULE_CODE": { "view": true, "delete": false } }
        /// </summary>
        public string? OverridesJson { get; set; }

        public DateTime UpdatedAt { get; set; } = DateTime.Now;
        public int UpdatedBy { get; set; }
    }

    // ==========================================
    // 2. MODULE CODES & ACTIONS
    // ==========================================

    public static class ModuleCodes
    {
        public const string Dashboard = "DASHBOARD";
        public const string MasterProject = "MASTER_PROJECT";
        public const string MasterConfig = "MASTER_CONFIG";
        public const string MasterBoxCapacity = "MASTER_BOX_CAPACITY";
        public const string MasterEnvelope = "MASTER_ENVELOPE";
        public const string MasterFields = "MASTER_FIELDS";
        public const string MasterModules = "MASTER_MODULES";
        public const string MasterNodalExtra = "MASTER_NODAL_EXTRA";
        public const string MasterRpt = "MASTER_RPT";
        public const string ProjectConfig = "PROJECT_CONFIG";

        // Data Import Sub-tabs
        public const string DataImport = "DATA_IMPORT";
        public const string UploadedData = "UPLOADED_DATA";
        public const string ConflictReport = "CONFLICT_REPORT";
        public const string FillMissingData = "FILL_MISSING_DATA";
        public const string LotBifurcation = "LOT_BIFURCATION";
        public const string MergeCatchNumbers = "MERGE_CATCH_NUMBERS";

        // Independent Processing & Report/Template Section
        public const string ProcessingPipeline = "PROCESSING_PIPELINE";
        public const string ReportTemplateManagement = "REPORT_TEMPLATE_MANAGEMENT";

        public const string HeaderVerification = "HEADER_VERIFICATION";
        public const string ChangedNr = "CHANGED_NR";
        public const string BatchData = "BATCH_DATA";
        public const string NodalCatch = "NODAL_CATCH";
        public const string HtovTool = "HTOV_TOOL";
        public const string CorrectionTool = "CORRECTION_TOOL";
        public const string Reports = "REPORTS";
        public const string RolesPermissions = "ROLES_PERMISSIONS";
        public const string MessageManagement = "MESSAGE_MANAGEMENT";

        public static readonly string[] All = new[]
        {
            MasterProject, ProjectConfig, MasterConfig, MasterBoxCapacity, MasterEnvelope,
            MasterFields, MasterModules, MasterNodalExtra, MasterRpt,
            DataImport, UploadedData, ConflictReport, FillMissingData, LotBifurcation, MergeCatchNumbers,
            ProcessingPipeline, ReportTemplateManagement,
            HeaderVerification, ChangedNr, BatchData, NodalCatch, HtovTool, CorrectionTool,
            Reports, RolesPermissions, MessageManagement
        };
    }

    public static class SpecialActionCodes
    {
        public const string DeleteAllNrData = "ACTION_DELETE_ALL_NR_DATA";
        public const string BulkVerifyCatches = "ACTION_BULK_VERIFY_CATCHES";
        public const string ApplyChangedNr = "ACTION_APPLY_CHANGED_NR";
        public const string PushNodalToNr = "ACTION_PUSH_NODAL_TO_NR";
        public const string PurgeReportRecord = "ACTION_PURGE_REPORT_RECORD";
        public const string UploadRptTemplate = "ACTION_UPLOAD_RPT_TEMPLATE";
        public const string MapRptTemplate = "ACTION_MAP_RPT_TEMPLATE";
        public const string SendToErp = "ACTION_SEND_TO_ERP";

        public static readonly string[] All = new[]
        {
            DeleteAllNrData,
            BulkVerifyCatches,
            ApplyChangedNr,
            PushNodalToNr,
            PurgeReportRecord,
            UploadRptTemplate,
            MapRptTemplate,
            SendToErp
        };
    }

    public class SpecialActionMetadata
    {
        public string Code { get; set; } = string.Empty;
        public string DisplayName { get; set; } = string.Empty;
        public string ParentModule { get; set; } = string.Empty;
        public PermissionAction RequiredStandardAction { get; set; }
        public string Description { get; set; } = string.Empty;
    }

    public enum PermissionAction
    {
        View,
        Create,
        Edit,
        Delete,
        Execute,
        Export
    }

    // ==========================================
    // 3. CLEAN DTOs (Request & Response)
    // ==========================================

    public class ModuleActionFlags
    {
        public bool? CanView { get; set; }
        public bool? CanCreate { get; set; }
        public bool? CanEdit { get; set; }
        public bool? CanDelete { get; set; }
        public bool? CanExecute { get; set; }
        public bool? CanExport { get; set; }
    }

    public class UserPermissionsDto
    {
        public int UserId { get; set; }
        public int RoleId { get; set; }
        public string RoleName { get; set; } = string.Empty;
        public bool IsSuperuser { get; set; }
        public Dictionary<string, ModuleActionFlags> Modules { get; set; } = new(StringComparer.OrdinalIgnoreCase);
        public Dictionary<string, bool> SpecialActions { get; set; } = new(StringComparer.OrdinalIgnoreCase);
    }

    public class RolePermissionItemDto
    {
        public string ModuleCode { get; set; } = string.Empty;
        public bool CanView { get; set; }
        public bool CanCreate { get; set; }
        public bool CanEdit { get; set; }
        public bool CanDelete { get; set; }
        public bool CanExecute { get; set; }
        public bool CanExport { get; set; }
    }

    public class UpdateRolePermissionsRequest
    {
        public List<RolePermissionItemDto> Permissions { get; set; } = new();
        public Dictionary<string, bool>? SpecialActions { get; set; }
    }

    public class UserOverrideItemDto
    {
        public string ModuleCode { get; set; } = string.Empty;
        public bool? CanView { get; set; }
        public bool? CanCreate { get; set; }
        public bool? CanEdit { get; set; }
        public bool? CanDelete { get; set; }
        public bool? CanExecute { get; set; }
        public bool? CanExport { get; set; }
    }

    public class UpdateUserPermissionsRequest
    {
        public int? RoleId { get; set; }
        public List<UserOverrideItemDto> Overrides { get; set; } = new();
    }

    public class CreateRoleRequest
    {
        public string RoleName { get; set; } = string.Empty;
        public string? Description { get; set; }
    }

    public class UpdateRoleRequest
    {
        public string RoleName { get; set; } = string.Empty;
        public string? Description { get; set; }
    }
}
