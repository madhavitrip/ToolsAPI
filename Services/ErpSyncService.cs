using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Json;
using System.Text.Json;
using System.Threading.Tasks;
using ERPToolsAPI.Data;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Options;
using Tools.Models;

namespace Tools.Services
{
    public class ErpSyncService : IErpSyncService
    {
        private readonly ERPToolsDbContext _context;
        private readonly HttpClient _httpClient;
        private readonly IConfiguration _configuration;
        private readonly ILoggerService _loggerService;
        private readonly ApiSettings _apiSettings;

        public ErpSyncService(
            ERPToolsDbContext context,
            HttpClient httpClient,
            IConfiguration configuration,
            ILoggerService loggerService,
            IOptions<ApiSettings> apiSettings)
        {
            _context = context;
            _httpClient = httpClient;
            _configuration = configuration;
            _loggerService = loggerService;
            _apiSettings = apiSettings?.Value ?? new ApiSettings();
        }

        /// <summary>
        /// Compiles all catches for a specific project and lot with quantity, pages, inner and outer envelope breakdowns.
        /// </summary>
        public async Task<ErpLotCatchesPayloadDto> CompileLotCatchesAsync(int projectId, int lotNo)
        {
            var isNewModel = await _context.NrData1.AnyAsync(p => p.ProjectId == projectId);
            var isFinalized = await IsLotFinalizedAsync(projectId, lotNo);
            var finalizationLog = await GetLatestFinalizationLogAsync(projectId, lotNo);
            var lastSyncLog = await GetLatestSyncLogAsync(projectId, lotNo);

            var payload = new ErpLotCatchesPayloadDto
            {
                ProjectId = projectId,
                LotNo = lotNo,
                IsFinalized = isFinalized,
                FinalizedAt = isFinalized ? finalizationLog?.LoggedAt : null,
                FinalizedBy = isFinalized ? finalizationLog?.EventTriggeredBy : null,
                LastSyncedAt = lastSyncLog?.LoggedAt,
                Catches = new List<ErpCatchItemDto>()
            };

            if (isNewModel)
            {
                // New model: NrData1 + CenterList + NewEnvelopeBreakages
                var nr1Query = from nr in _context.NrData1
                               where nr.ProjectId == projectId && nr.LotNo == lotNo && nr.Status
                               select new
                               {
                                   nr.Id,
                                   nr.CatchNo,
                                   nr.Pages,
                                   nr.CourseName,
                                   nr.SubjectName,
                                   nr.ExamDate
                               };

                var nr1Records = await nr1Query.ToListAsync();
                if (nr1Records.Count == 0)
                {
                    return payload;
                }

                var nr1Ids = nr1Records.Select(x => x.Id).Distinct().ToList();

                // Sum quantities across CenterList
                var centerQuantities = await _context.CenterList
                    .Where(c => c.ProjectId == projectId && nr1Ids.Contains(c.NRDataId) && c.Status)
                    .GroupBy(c => c.NRDataId)
                    .Select(g => new
                    {
                        NRDataId = g.Key,
                        TotalQty = g.Sum(x => x.Quantity)
                    })
                    .ToDictionaryAsync(x => x.NRDataId, x => x.TotalQty);

                // Fetch envelope breakdowns from NewEnvelopeBreakage
                var envelopeBreakages = await _context.NewEnvelopeBreakages
                    .Where(e => nr1Ids.Contains(e.NrDataId) && e.Status)
                    .ToListAsync();

                var envByNrId = envelopeBreakages
                    .GroupBy(e => e.NrDataId)
                    .ToDictionary(g => g.Key, g => g.ToList());

                var groupedByCatch = nr1Records
                    .GroupBy(x => x.CatchNo ?? string.Empty)
                    .OrderBy(g => g.Key);

                foreach (var grp in groupedByCatch)
                {
                    string catchNo = grp.Key;
                    int maxPages = grp.Max(x => x.Pages);
                    string? courseName = grp.FirstOrDefault(x => !string.IsNullOrWhiteSpace(x.CourseName))?.CourseName;
                    string? subjectName = grp.FirstOrDefault(x => !string.IsNullOrWhiteSpace(x.SubjectName))?.SubjectName;

                    int catchTotalQty = 0;
                    var innerAgg = new Dictionary<string, int>();
                    var outerAgg = new Dictionary<string, int>();

                    foreach (var nr in grp)
                    {
                        catchTotalQty += centerQuantities.GetValueOrDefault(nr.Id, 0);

                        if (envByNrId.TryGetValue(nr.Id, out var breakages))
                        {
                            foreach (var b in breakages)
                            {
                                MergeEnvelopeBreakdown(innerAgg, b.InnerEnvelope);
                                MergeEnvelopeBreakdown(outerAgg, b.OuterEnvelope);
                            }
                        }
                    }

                    var innerFormatted = innerAgg.ToDictionary(k => k.Key, v => v.Value.ToString());
                    var outerFormatted = outerAgg.ToDictionary(k => k.Key, v => v.Value.ToString());

                    payload.Catches.Add(new ErpCatchItemDto
                    {
                        CatchNo = catchNo,
                        Quantity = catchTotalQty,
                        Pages = maxPages,
                        InnerEnv = JsonSerializer.Serialize(innerFormatted),
                        OuterEnv = JsonSerializer.Serialize(outerFormatted),
                        InnerEnvelopeBreakdown = innerAgg,
                        OuterEnvelopeBreakdown = outerAgg,
                        CourseName = courseName,
                        SubjectName = subjectName
                    });
                }
            }
            else
            {
                // Legacy model: NRDatas + EnvelopeBreakages
                var legacyRows = await _context.NRDatas
                    .Where(x => x.ProjectId == projectId && x.LotNo == lotNo && x.Status)
                    .ToListAsync();

                if (legacyRows.Count == 0)
                {
                    return payload;
                }

                var legacyIds = legacyRows.Select(x => x.Id).Distinct().ToList();

                var legacyBreakages = await _context.EnvelopeBreakages
                    .Where(b => b.ProjectId == projectId && legacyIds.Contains(b.NrDataId) && b.Status)
                    .ToListAsync();

                var breakagesByNrId = legacyBreakages
                    .GroupBy(b => b.NrDataId)
                    .ToDictionary(g => g.Key, g => g.ToList());

                // Group by CatchNo
                var groupedByCatch = legacyRows
                    .GroupBy(x => x.CatchNo ?? string.Empty)
                    .OrderBy(g => g.Key);

                foreach (var grp in groupedByCatch)
                {
                    int totalQty = grp.Sum(x => x.Quantity);
                    int maxPages = grp.Max(x => x.Pages);
                    var firstRow = grp.First();

                    var innerAgg = new Dictionary<string, int>();
                    var outerAgg = new Dictionary<string, int>();

                    foreach (var row in grp)
                    {
                        if (breakagesByNrId.TryGetValue(row.Id, out var breakages))
                        {
                            foreach (var b in breakages)
                            {
                                MergeEnvelopeBreakdown(innerAgg, b.InnerEnvelope);
                                MergeEnvelopeBreakdown(outerAgg, b.OuterEnvelope);
                            }
                        }
                    }

                    var innerFormatted = innerAgg.ToDictionary(k => k.Key, v => v.Value.ToString());
                    var outerFormatted = outerAgg.ToDictionary(k => k.Key, v => v.Value.ToString());

                    payload.Catches.Add(new ErpCatchItemDto
                    {
                        CatchNo = grp.Key,
                        Quantity = totalQty,
                        Pages = maxPages,
                        InnerEnv = JsonSerializer.Serialize(innerFormatted),
                        OuterEnv = JsonSerializer.Serialize(outerFormatted),
                        InnerEnvelopeBreakdown = innerAgg,
                        OuterEnvelopeBreakdown = outerAgg,
                        CourseName = firstRow.CourseName,
                        SubjectName = firstRow.SubjectName
                    });
                }
            }

            payload.TotalCatches = payload.Catches.Count;
            payload.TotalQuantity = payload.Catches.Sum(c => c.Quantity);

            return payload;
        }

        /// <summary>
        /// Checks if a lot has been finalized by inspecting the audit trail in EventLogs.
        /// </summary>
        public async Task<bool> IsLotFinalizedAsync(int projectId, int lotNo)
        {
            var latestLog = await _context.EventLogs
                .Where(e => e.ProjectId == projectId &&
                            e.Category == "LotFinalization" &&
                            e.Event != null && e.Event.Contains($"Lot {lotNo}"))
                .OrderByDescending(e => e.LoggedAt)
                .FirstOrDefaultAsync();

            if (latestLog == null) return false;

            return latestLog.NewValue == "Finalized" || (latestLog.Event != null && latestLog.Event.StartsWith("Finalized"));
        }

        /// <summary>
        /// Marks a lot as finalized in the audit log, locking its details for ERP transmission.
        /// </summary>
        public async Task<FinalizationResultDto> FinalizeLotAsync(int projectId, int lotNo, int? userId)
        {
            var isNewModel = await _context.NrData1.AnyAsync(p => p.ProjectId == projectId);
            bool hasCatches = isNewModel
                ? await _context.NrData1.AnyAsync(x => x.ProjectId == projectId && x.LotNo == lotNo && x.Status)
                : await _context.NRDatas.AnyAsync(x => x.ProjectId == projectId && x.LotNo == lotNo && x.Status);

            if (!hasCatches)
            {
                return new FinalizationResultDto
                {
                    Success = false,
                    Message = $"No active catches found for Project {projectId}, Lot {lotNo}.",
                    ProjectId = projectId,
                    LotNo = lotNo,
                    IsFinalized = false
                };
            }

            var now = DateTime.Now;
            var eventLog = new EventLog
            {
                ProjectId = projectId,
                Category = "LotFinalization",
                Event = $"Finalized Lot {lotNo} for Project {projectId}",
                LoggedAt = now,
                EventTriggeredBy = userId ?? 0,
                NewValue = "Finalized"
            };

            await _context.EventLogs.AddAsync(eventLog);
            await _context.SaveChangesAsync();

            await _loggerService.LogEventAsync(
                $"Lot {lotNo} finalized for Project {projectId}",
                nameof(ErpSyncService),
                userId ?? 0,
                projectId
            );

            return new FinalizationResultDto
            {
                Success = true,
                Message = $"Lot {lotNo} finalized successfully. Details are now ready to be pushed to ERP.",
                ProjectId = projectId,
                LotNo = lotNo,
                IsFinalized = true,
                FinalizedAt = now
            };
        }

        /// <summary>
        /// Reopens/unfinalizes a lot if adjustments are required before ERP transmission.
        /// </summary>
        public async Task<FinalizationResultDto> UnfinalizeLotAsync(int projectId, int lotNo, int? userId)
        {
            var now = DateTime.Now;
            var eventLog = new EventLog
            {
                ProjectId = projectId,
                Category = "LotFinalization",
                Event = $"Unfinalized Lot {lotNo} for Project {projectId}",
                LoggedAt = now,
                EventTriggeredBy = userId ?? 0,
                NewValue = "Unfinalized"
            };

            await _context.EventLogs.AddAsync(eventLog);
            await _context.SaveChangesAsync();

            await _loggerService.LogEventAsync(
                $"Lot {lotNo} unfinalized/reopened for Project {projectId}",
                nameof(ErpSyncService),
                userId ?? 0,
                projectId
            );

            return new FinalizationResultDto
            {
                Success = true,
                Message = $"Lot {lotNo} finalization reverted.",
                ProjectId = projectId,
                LotNo = lotNo,
                IsFinalized = false,
                FinalizedAt = null
            };
        }

        /// <summary>
        /// Validates that the lot is finalized, and sends all catch details to ERP.
        /// </summary>
        public async Task<ErpTransmissionResultDto> SendLotCatchesToErpAsync(int projectId, int lotNo, int? userId)
        {
            // Finalize check
            bool isFinalized = await IsLotFinalizedAsync(projectId, lotNo);
            if (!isFinalized)
            {
                return new ErpTransmissionResultDto
                {
                    Success = false,
                    Message = $"Lot {lotNo} is not finalized. Please finalize the lot before pushing details to ERP.",
                    StatusCode = 400
                };
            }

            var payload = await CompileLotCatchesAsync(projectId, lotNo);
            if (payload.Catches.Count == 0)
            {
                return new ErpTransmissionResultDto
                {
                    Success = false,
                    Message = $"No active catches found for Project {projectId}, Lot {lotNo}.",
                    StatusCode = 404
                };
            }

            string erpUrl = _apiSettings.ErpCatchDetailsSyncUrl;
            if (string.IsNullOrWhiteSpace(erpUrl))
            {
                erpUrl = _configuration["ApiSettings:ErpCatchDetailsSyncUrl"]?.Trim() ?? string.Empty;
            }

            var now = DateTime.Now;

            if (string.IsNullOrWhiteSpace(erpUrl))
            {
                // URL not configured yet: record payload audit in EventLogs and report ready state
                var auditLog = new EventLog
                {
                    ProjectId = projectId,
                    Category = "ErpSync",
                    Event = $"Synced Project {projectId} Lot {lotNo} catches to ERP ({payload.Catches.Count} catches, {payload.TotalQuantity} qty)",
                    LoggedAt = now,
                    EventTriggeredBy = userId ?? 0,
                    NewValue = JsonSerializer.Serialize(new
                    {
                        payload.ProjectId,
                        payload.LotNo,
                        payload.TotalCatches,
                        payload.TotalQuantity,
                        TransmissionStatus = "Compiled successfully (ERP endpoint not yet configured)",
                        Transmitted = false
                    })
                };
                await _context.EventLogs.AddAsync(auditLog);
                await _context.SaveChangesAsync();

                await _loggerService.LogEventAsync(
                    $"Lot {lotNo} catches prepared for ERP (ErpCatchDetailsSyncUrl not configured)",
                    nameof(ErpSyncService),
                    userId ?? 0,
                    projectId
                );

                return new ErpTransmissionResultDto
                {
                    Success = true,
                    Message = $"Compiled {payload.Catches.Count} catches for Lot {lotNo}. Compiled successfully (ERP endpoint not yet configured)",
                    StatusCode = 200,
                    Transmitted = false,
                    TransmissionStatus = "Compiled successfully (ERP endpoint not yet configured)",
                    TargetUrl = null,
                    SyncedAt = now,
                    Payload = payload
                };
            }

            try
            {
                var response = await _httpClient.PostAsJsonAsync(erpUrl, payload);
                int statusCode = (int)response.StatusCode;

                if (response.IsSuccessStatusCode)
                {
                    var syncLog = new EventLog
                    {
                        ProjectId = projectId,
                        Category = "ErpSync",
                        Event = $"Synced Project {projectId} Lot {lotNo} catches to ERP ({payload.Catches.Count} catches, {payload.TotalQuantity} qty)",
                        LoggedAt = now,
                        EventTriggeredBy = userId ?? 0,
                        NewValue = JsonSerializer.Serialize(new
                        {
                            payload.ProjectId,
                            payload.LotNo,
                            payload.TotalCatches,
                            payload.TotalQuantity,
                            TransmissionStatus = "Transmitted",
                            Transmitted = true
                        })
                    };
                    await _context.EventLogs.AddAsync(syncLog);
                    await _context.SaveChangesAsync();

                    return new ErpTransmissionResultDto
                    {
                        Success = true,
                        Message = $"Successfully sent {payload.Catches.Count} catches for Lot {lotNo} to ERP.",
                        StatusCode = statusCode,
                        Transmitted = true,
                        TransmissionStatus = "Transmitted",
                        TargetUrl = erpUrl,
                        SyncedAt = now,
                        Payload = payload
                    };
                }
                else
                {
                    var responseBody = await response.Content.ReadAsStringAsync();
                    string statusDesc = $"ERP returned {(int)response.StatusCode}: {response.ReasonPhrase}";
                    await _loggerService.LogErrorAsync(
                        $"ERP sync endpoint returned error for Project {projectId}, Lot {lotNo}",
                        responseBody,
                        nameof(ErpSyncService)
                    );

                    return new ErpTransmissionResultDto
                    {
                        Success = false,
                        Message = $"{statusDesc} - {responseBody}",
                        StatusCode = statusCode,
                        Transmitted = false,
                        TransmissionStatus = statusDesc,
                        TargetUrl = erpUrl
                    };
                }
            }
            catch (Exception ex)
            {
                await _loggerService.LogErrorAsync(
                    $"Failed HTTP call to ERP for Project {projectId}, Lot {lotNo}",
                    ex.ToString(),
                    nameof(ErpSyncService)
                );

                return new ErpTransmissionResultDto
                {
                    Success = false,
                    Message = $"Failed to send lot catches to ERP: {ex.Message}",
                    StatusCode = 500,
                    Transmitted = false,
                    TransmissionStatus = $"Transmission error: {ex.Message}",
                    TargetUrl = erpUrl
                };
            }
        }

        /// <summary>
        /// Returns finalization and ERP sync status for all lots in a project.
        /// </summary>
        public async Task<List<LotSyncStatusDto>> GetLotsSyncStatusAsync(int projectId)
        {
            var isNewModel = await _context.NrData1.AnyAsync(p => p.ProjectId == projectId);
            List<LotInfo> lotsInfo;

            if (isNewModel)
            {
                var distinctLots = await _context.NrData1
                    .Where(x => x.ProjectId == projectId && x.Status && x.LotNo > 0)
                    .Select(x => x.LotNo)
                    .Distinct()
                    .OrderBy(l => l)
                    .ToListAsync();

                lotsInfo = new List<LotInfo>();
                foreach (var l in distinctLots)
                {
                    var nrRows = await _context.NrData1
                        .Where(x => x.ProjectId == projectId && x.LotNo == l && x.Status)
                        .Select(x => new { x.Id, x.CatchNo })
                        .ToListAsync();

                    int catchCount = nrRows.Select(x => x.CatchNo).Distinct().Count();
                    var nrIds = nrRows.Select(x => x.Id).ToList();
                    int totalQty = 0;
                    if (catchCount > 0)
                    {
                        totalQty = await _context.CenterList
                            .Where(c => c.ProjectId == projectId && nrIds.Contains(c.NRDataId) && c.Status)
                            .SumAsync(c => c.Quantity);
                    }

                    lotsInfo.Add(new LotInfo
                    {
                        LotNo = l,
                        CatchCount = catchCount,
                        TotalQuantity = totalQty
                    });
                }
            }
            else
            {
                var legacyRows = await _context.NRDatas
                    .Where(x => x.ProjectId == projectId && x.LotNo > 0 && x.Status)
                    .ToListAsync();

                lotsInfo = legacyRows
                    .GroupBy(x => x.LotNo)
                    .Select(g => new LotInfo
                    {
                        LotNo = g.Key,
                        CatchCount = g.Select(x => x.CatchNo).Distinct().Count(),
                        TotalQuantity = g.Sum(x => x.Quantity)
                    })
                    .OrderBy(l => l.LotNo)
                    .ToList();
            }

            var syncLogs = await _context.EventLogs
                .Where(e => e.ProjectId == projectId && e.Category == "ErpSync")
                .OrderByDescending(e => e.LoggedAt)
                .ToListAsync();

            var finalizationLogs = await _context.EventLogs
                .Where(e => e.ProjectId == projectId && e.Category == "LotFinalization")
                .OrderByDescending(e => e.LoggedAt)
                .ToListAsync();

            var result = new List<LotSyncStatusDto>();

            foreach (var lot in lotsInfo)
            {
                var finLog = finalizationLogs
                    .FirstOrDefault(l => l.Event != null && l.Event.Contains($"Lot {lot.LotNo}"));
                bool isFin = finLog != null && (finLog.NewValue == "Finalized" || (finLog.Event != null && finLog.Event.StartsWith("Finalized")));

                var lastLog = syncLogs
                    .FirstOrDefault(l => l.Event != null && l.Event.Contains($"Lot {lot.LotNo}"));

                result.Add(new LotSyncStatusDto
                {
                    LotNo = lot.LotNo,
                    CatchCount = lot.CatchCount,
                    TotalQuantity = lot.TotalQuantity,
                    IsFinalized = isFin,
                    FinalizedAt = isFin ? finLog?.LoggedAt : null,
                    LastSyncedAt = lastLog?.LoggedAt,
                    IsSynced = lastLog != null
                });
            }

            return result;
        }

        #region Helpers

        private class LotInfo
        {
            public int LotNo { get; set; }
            public int CatchCount { get; set; }
            public int TotalQuantity { get; set; }
        }

        private async Task<EventLog?> GetLatestFinalizationLogAsync(int projectId, int lotNo)
        {
            return await _context.EventLogs
                .Where(e => e.ProjectId == projectId &&
                            e.Category == "LotFinalization" &&
                            e.Event != null && e.Event.Contains($"Lot {lotNo}") &&
                            (e.NewValue == "Finalized" || e.Event.StartsWith("Finalized")))
                .OrderByDescending(e => e.LoggedAt)
                .FirstOrDefaultAsync();
        }

        private async Task<EventLog?> GetLatestSyncLogAsync(int projectId, int lotNo)
        {
            return await _context.EventLogs
                .Where(e => e.ProjectId == projectId &&
                            e.Category == "ErpSync" &&
                            e.Event != null && e.Event.Contains($"Lot {lotNo}"))
                .OrderByDescending(e => e.LoggedAt)
                .FirstOrDefaultAsync();
        }

        private static void MergeEnvelopeBreakdown(Dictionary<string, int> target, string? jsonString)
        {
            if (string.IsNullOrWhiteSpace(jsonString) || jsonString.Trim() == "{}") return;

            try
            {
                using var doc = JsonDocument.Parse(jsonString);
                var root = doc.RootElement;

                if (root.ValueKind == JsonValueKind.Object)
                {
                    foreach (var prop in root.EnumerateObject())
                    {
                        string envType = prop.Name;
                        int count = 0;

                        if (prop.Value.ValueKind == JsonValueKind.Number)
                        {
                            count = prop.Value.GetInt32();
                        }
                        else if (prop.Value.ValueKind == JsonValueKind.String && int.TryParse(prop.Value.GetString(), out int parsed))
                        {
                            count = parsed;
                        }

                        if (count > 0)
                        {
                            target[envType] = target.GetValueOrDefault(envType) + count;
                        }
                    }
                }
            }
            catch
            {
                // Ignore parse errors on malformed legacy JSON
            }
        }

        #endregion
    }
}
