using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Tools.Services
{
    public class ErpCatchItemDto
    {
        public string CatchNo { get; set; } = string.Empty;
        public int Quantity { get; set; }
        public int Pages { get; set; }
        public string InnerEnv { get; set; } = "{}";
        public string OuterEnv { get; set; } = "{}";
        public Dictionary<string, int> InnerEnvelopeBreakdown { get; set; } = new();
        public Dictionary<string, int> OuterEnvelopeBreakdown { get; set; } = new();
        public string? CourseName { get; set; }
        public string? SubjectName { get; set; }
        public string? ExamDate { get; set; }
        public string? ExamTime { get; set; }
        public int? CourseId { get; set; }
        public int? SubjectId { get; set; }
    }

    public class ErpQuantitySheetPayloadItem
    {
        public int ProjectId { get; set; }
        public string LotNo { get; set; } = string.Empty;
        public string CatchNo { get; set; } = string.Empty;
        public double Quantity { get; set; }
        public int Pages { get; set; }
        public string? InnerEnvelope { get; set; }
        public int? OuterEnvelope { get; set; }
        public string? ExamDate { get; set; }
        public string? ExamTime { get; set; }
        public int? TTFStatus { get; set; }
        public int CourseId { get; set; }
        public int SubjectId { get; set; }
        public int ExamTypeId { get; set; }
        public int Status { get; set; } = 1;
        public string? PaperTitle { get; set; }
        public string? PaperNumber { get; set; }
        public int QPId { get; set; }
    }

    public class ErpLotCatchesPayloadDto
    {
        public int ProjectId { get; set; }
        public int LotNo { get; set; }
        public int TotalCatches { get; set; }
        public int TotalQuantity { get; set; }
        public bool IsFinalized { get; set; }
        public DateTime? FinalizedAt { get; set; }
        public int? FinalizedBy { get; set; }
        public DateTime? LastSyncedAt { get; set; }
        public List<ErpCatchItemDto> Catches { get; set; } = new();
    }

    public class FinalizationResultDto
    {
        public bool Success { get; set; }
        public string Message { get; set; } = string.Empty;
        public int ProjectId { get; set; }
        public int LotNo { get; set; }
        public bool IsFinalized { get; set; }
        public DateTime? FinalizedAt { get; set; }
    }

    public class SendLotCatchesRequest
    {
        public int ProjectId { get; set; }
        public int LotNo { get; set; }
        public int? UserId { get; set; }
    }

    public class FinalizeLotRequest
    {
        public int ProjectId { get; set; }
        public int LotNo { get; set; }
        public int? UserId { get; set; }
    }

    public class ErpTransmissionResultDto
    {
        public bool Success { get; set; }
        public string Message { get; set; } = string.Empty;
        public int StatusCode { get; set; }
        public bool Transmitted { get; set; }
        public string TransmissionStatus { get; set; } = string.Empty;
        public string? TargetUrl { get; set; }
        public DateTime? SyncedAt { get; set; }
        public ErpLotCatchesPayloadDto? Payload { get; set; }
    }

    public class LotSyncStatusDto
    {
        public int LotNo { get; set; }
        public int CatchCount { get; set; }
        public int TotalQuantity { get; set; }
        public bool IsFinalized { get; set; }
        public DateTime? FinalizedAt { get; set; }
        public DateTime? LastSyncedAt { get; set; }
        public bool IsSynced { get; set; }
    }

    public interface IErpSyncService
    {
        /// <summary>
        /// Compiles all catches for a specific project and lot with quantity, pages, inner and outer envelope breakdowns.
        /// </summary>
        Task<ErpLotCatchesPayloadDto> CompileLotCatchesAsync(int projectId, int lotNo);

        /// <summary>
        /// Checks if a lot has been finalized by inspecting the audit trail.
        /// </summary>
        Task<bool> IsLotFinalizedAsync(int projectId, int lotNo);

        /// <summary>
        /// Marks a lot as finalized in the audit log, locking its details for ERP transmission.
        /// </summary>
        Task<FinalizationResultDto> FinalizeLotAsync(int projectId, int lotNo, int? userId);

        /// <summary>
        /// Reopens/unfinalizes a lot if adjustments are required before ERP transmission.
        /// </summary>
        Task<FinalizationResultDto> UnfinalizeLotAsync(int projectId, int lotNo, int? userId);

        /// <summary>
        /// Validates that the lot is finalized, and sends all catch details to ERP.
        /// </summary>
        Task<ErpTransmissionResultDto> SendLotCatchesToErpAsync(int projectId, int lotNo, int? userId);

        /// <summary>
        /// Returns finalization and ERP sync status for all lots in a project.
        /// </summary>
        Task<List<LotSyncStatusDto>> GetLotsSyncStatusAsync(int projectId);
    }
}
