using ERPToolsAPI.Data;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Tools.Services;

namespace Tools.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class NRDataLotsController : ControllerBase
    {
        private readonly ERPToolsDbContext _context;
        private readonly ILoggerService _loggerService;
        private readonly IDispatchService _dispatchService;

        public NRDataLotsController(ERPToolsDbContext context, ILoggerService loggerService, IDispatchService dispatchService)
        {
            _context = context;
            _loggerService = loggerService;
            _dispatchService = dispatchService;
        }

        public class LotAssignmentItem
        {
            public string CatchNo { get; set; } = string.Empty;
            public int LotNo { get; set; }
        }

        public class LotAssignmentRequest
        {
            public int ProjectId { get; set; }
            public List<LotAssignmentItem> Updates { get; set; } = new();
        }

        [HttpGet("by-project/{projectId}")]
        public async Task<IActionResult> GetLotsByProject(int projectId)
        {
            if (projectId <= 0)
                return BadRequest("ProjectId is required.");

            var rows = await _context.NRDatas
                .Where(x => x.ProjectId == projectId)
                .Select(x => new
                {
                    x.CatchNo,
                    x.LotNo
                })
                .ToListAsync();

            return Ok(rows);
        }

        public class EnvLotAssignmentRequest
        {
            public int ProjectId { get; set; }
            public List<string> CatchNos { get; set; } = new();
        }

        [HttpGet("GetMissingEnvLotCatches/{projectId}")]
        public async Task<IActionResult> GetMissingEnvLotCatches(int projectId)
        {
            if (projectId <= 0)
                return BadRequest("ProjectId is required.");

            var catchItems = await _context.NRDatas
                .Where(x => x.ProjectId == projectId && x.Status == true && x.EnvLotNo == 0 && !string.IsNullOrWhiteSpace(x.CatchNo))
                .GroupBy(x => x.CatchNo)
                .Select(g => new
                {
                    CatchNo = g.Key,
                    Count = g.Count()
                })
                .OrderBy(x => x.CatchNo)
                .ToListAsync();

            return Ok(catchItems);
        }

        [HttpGet("GetAssignedEnvLotCatches/{projectId}")]
        public async Task<IActionResult> GetAssignedEnvLotCatches(int projectId)
        {
            if (projectId <= 0)
                return BadRequest("ProjectId is required.");

            var catchItems = await _context.NRDatas
                .Where(x => x.ProjectId == projectId && x.Status == true && x.EnvLotNo > 0 && !string.IsNullOrWhiteSpace(x.CatchNo))
                .GroupBy(x => new { x.EnvLotNo, x.CatchNo })
                .Select(g => new
                {
                    EnvLotNo = g.Key.EnvLotNo,
                    CatchNo = g.Key.CatchNo
                })
                .OrderBy(x => x.EnvLotNo)
                .ThenBy(x => x.CatchNo)
                .ToListAsync();

            return Ok(catchItems);
        }

        [HttpPost("AssignEnvLotByCatches")]
        public async Task<IActionResult> AssignEnvLotByCatches([FromBody] EnvLotAssignmentRequest request)
        {
            if (!ModelState.IsValid)
                return BadRequest(ModelState);

            if (request.ProjectId <= 0)
                return BadRequest("ProjectId is required.");

            if (request.CatchNos == null || request.CatchNos.Count == 0)
                return BadRequest("No catch numbers provided.");

            var catchNos = request.CatchNos
                .Where(c => !string.IsNullOrWhiteSpace(c))
                .Select(c => c.Trim())
                .Distinct()
                .ToList();

            if (catchNos.Count == 0)
                return BadRequest("No valid catch numbers provided.");

            var maxEnvLotNo = await _context.NRDatas
                .Where(x => x.ProjectId == request.ProjectId && x.Status == true && x.EnvLotNo > 0)
                .MaxAsync(x => (int?)x.EnvLotNo) ?? 0;

            var nextEnvLotNo = maxEnvLotNo + 1;

            var rows = await _context.NRDatas
                .Where(x => x.ProjectId == request.ProjectId && x.EnvLotNo == 0 && catchNos.Contains(x.CatchNo))
                .ToListAsync();

            if (rows.Count == 0)
                return NotFound("No matching unassigned NRData rows found.");

            foreach (var row in rows)
            {
                row.EnvLotNo = nextEnvLotNo;
            }

            await _context.SaveChangesAsync();
            _loggerService.LogEventAsync(
                $"Assigned EnvLotNo {nextEnvLotNo} to {rows.Count} NRData rows (ProjectId {request.ProjectId})",
                "NRDataLots",
                LogHelper.GetTriggeredBy(User),
                request.ProjectId
            );

            return Ok(new
            {
                message = "Envelope lot created and assigned successfully.",
                assignedEnvLotNo = nextEnvLotNo
            });
        }

        [HttpPost("UnassignEnvLotByCatches")]
        public async Task<IActionResult> UnassignEnvLotByCatches([FromBody] EnvLotAssignmentRequest request)
        {
            if (!ModelState.IsValid)
                return BadRequest(ModelState);

            if (request.ProjectId <= 0)
                return BadRequest("ProjectId is required.");

            if (request.CatchNos == null || request.CatchNos.Count == 0)
                return BadRequest("No catch numbers provided.");

            var catchNos = request.CatchNos
                .Where(c => !string.IsNullOrWhiteSpace(c))
                .Select(c => c.Trim())
                .Distinct()
                .ToList();

            if (catchNos.Count == 0)
                return BadRequest("No valid catch numbers provided.");

            var rows = await _context.NRDatas
                .Where(x => x.ProjectId == request.ProjectId && catchNos.Contains(x.CatchNo) && x.EnvLotNo > 0)
                .ToListAsync();

            if (rows.Count == 0)
                return NotFound("No matching NRData rows found for unassignment.");

            foreach (var row in rows)
            {
                row.EnvLotNo = 0;
            }

            await _context.SaveChangesAsync();
            _loggerService.LogEventAsync(
                $"Unassigned EnvLotNo from {rows.Count} NRData rows (ProjectId {request.ProjectId})",
                "NRDataLots",
                LogHelper.GetTriggeredBy(User),
                request.ProjectId
            );

            return Ok(new
            {
                message = "Envelope lot assignment reverted successfully.",
                revertedCount = rows.Count
            });
        }

        [HttpGet("GetByProjectId/{projectId}")]
        public async Task<IActionResult> GetUniqueLotsByProject(int projectId)
        {
            if (projectId <= 0)
                return BadRequest("ProjectId is required.");

            try
            {
                // Get unique lots with catch counts
                var lots = await _context.NRDatas
                    .Where(x => x.ProjectId == projectId && x.Status == true)
                    .GroupBy(x => x.LotNo)
                    .Select(g => new
                    {
                        lotNo = g.Key,
                        catchCount = g.Select(x => x.CatchNo).Distinct().Count(),
                        minStep = g.Min(x => x.Steps)
                    })
                    .OrderBy(x => x.lotNo)
                    .ToListAsync();

                if (lots.Count == 0)
                {
                    return Ok(new List<object>());
                }

                return Ok(lots);
            }
            catch (Exception ex)
            {
                await _loggerService.LogErrorAsync(
                    "Error fetching unique lots",
                    ex.Message,
                    nameof(NRDataLotsController)
                );
                return StatusCode(500, new { message = "Failed to fetch lots", error = ex.Message });
            }
        }

        [HttpGet("GetLotsWithDispatchInfo/{projectId}")]
        public async Task<IActionResult> GetLotsWithDispatchInfo(int projectId)
        {
            if (projectId <= 0)
                return BadRequest("ProjectId is required.");

            try
            {
                // Get unique lots with catch counts
                var lots = await _context.NRDatas
                    .Where(x => x.ProjectId == projectId && x.Status == true)
                    .GroupBy(x => x.LotNo)
                    .Select(g => new
                    {
                        lotNo = g.Key,
                        catchCount = g.Select(x => x.CatchNo).Distinct().Count(),
                        minStep = g.Min(x => x.Steps)
                    })
                    .OrderBy(x => x.lotNo)
                    .ToListAsync();

                if (lots.Count == 0)
                {
                    return Ok(new List<object>());
                }

                // Fetch dispatch info for all lots in parallel
                var lotNos = lots.Select(l => l.lotNo).ToList();
                var dispatchInfoDict = await _dispatchService.GetDispatchDatesAsync(projectId, lotNos);

                // Merge dispatch info with lot data
                var lotsWithDispatch = lots.Select(lot => new
                {
                    lot.lotNo,
                    lot.catchCount,
                    lot.minStep,
                    isDispatched = dispatchInfoDict.ContainsKey(lot.lotNo) && dispatchInfoDict[lot.lotNo].IsDispatched,
                    dispatchDate = dispatchInfoDict.ContainsKey(lot.lotNo) ? dispatchInfoDict[lot.lotNo].DispatchDate : null
                }).ToList();

                return Ok(lotsWithDispatch);
            }
            catch (Exception ex)
            {
                await _loggerService.LogErrorAsync(
                    "Error fetching lots with dispatch info",
                    ex.Message,
                    nameof(NRDataLotsController)
                );
                return StatusCode(500, new { message = "Failed to fetch lots with dispatch info", error = ex.Message });
            }
        }

        [HttpGet("GetByProject/{projectId}")]
        public async Task<IActionResult> GetUniqueEnvLotsByProject(int projectId)
        {
            if (projectId <= 0)
                return BadRequest("ProjectId is required.");

            try
            {
                // Get unique lots with catch counts
                var lots = await _context.NRDatas
                    .Where(x => x.ProjectId == projectId && x.Status == true)
                    .GroupBy(x => x.EnvLotNo)
                    .Select(g => new
                    {
                        EnvLotno = g.Key,
                        catchCount = g.Select(x => x.CatchNo).Distinct().Count()
                    })
                    .OrderBy(x => x.EnvLotno)
                    .ToListAsync();

                if (lots.Count == 0)
                {
                    return Ok(new List<object>());
                }

                return Ok(lots);
            }
            catch (Exception ex)
            {
                await _loggerService.LogErrorAsync(
                    "Error fetching unique lots",
                    ex.Message,
                    nameof(NRDataLotsController)
                );
                return StatusCode(500, new { message = "Failed to fetch lots", error = ex.Message });
            }
        }

        [HttpPut("assign")]
        public async Task<IActionResult> AssignLots([FromBody] LotAssignmentRequest request)
        {
            if (!ModelState.IsValid)
                return BadRequest(ModelState);

            if (request.ProjectId <= 0)
                return BadRequest("ProjectId is required.");

            if (request.Updates == null || request.Updates.Count == 0)
                return BadRequest("No updates provided.");

            var catchNos = request.Updates
                .Select(u => u.CatchNo)
                .Where(c => !string.IsNullOrWhiteSpace(c))
                .Distinct()
                .ToList();

            if (catchNos.Count == 0)
                return BadRequest("No valid CatchNo values provided.");

            var rows = await _context.NRDatas
                .Where(x => x.ProjectId == request.ProjectId && catchNos.Contains(x.CatchNo))
                .ToListAsync();

            if (rows.Count == 0)
                return NotFound("No matching NRData rows found.");

            // 🔹 OMR/Booklet serial validation
            var projectConfig = await _context.ProjectConfigs
                .FirstOrDefaultAsync(x => x.ProjectId == request.ProjectId);

            int resetStep = 5;
            bool isBothSerialsPositive = projectConfig != null && projectConfig.OmrSerialNumber > 0 && (projectConfig.BookletSerialNumber ?? 0) > 0;
            bool resetBookletSerial = projectConfig?.ResetBookletSerialOnCatchChange ?? false;

            if (isBothSerialsPositive && !resetBookletSerial)
            {
                resetStep = 4;
            }

            // 🔹 Detect changed or moved catches
            var originalLotByCatch = rows
                .Where(r => r.CatchNo != null)
                .GroupBy(r => r.CatchNo)
                .ToDictionary(g => g.Key, g => g.First().LotNo);

            var changedCatchNos = new List<string>();
            var affectedLotNos = new HashSet<int>();

            foreach (var update in request.Updates)
            {
                if (string.IsNullOrWhiteSpace(update.CatchNo)) continue;

                if (originalLotByCatch.TryGetValue(update.CatchNo, out var oldLotNo))
                {
                    if (oldLotNo != update.LotNo)
                    {
                        changedCatchNos.Add(update.CatchNo);
                        affectedLotNos.Add(oldLotNo);
                        affectedLotNos.Add(update.LotNo);
                    }
                }
            }

            // 🔹 Apply lot updates
            var updatesByCatch = request.Updates
                .Where(u => !string.IsNullOrWhiteSpace(u.CatchNo))
                .GroupBy(u => u.CatchNo)
                .ToDictionary(g => g.Key, g => g.Last().LotNo);

            foreach (var row in rows)
            {
                if (row.CatchNo == null) continue;
                if (updatesByCatch.TryGetValue(row.CatchNo, out var lotNo))
                {
                    row.LotNo = lotNo;
                }
            }

            // 🔹 Reset steps if there were changes/moves
            if (changedCatchNos.Any())
            {
                // Reset all records of the moved/changed catches
                foreach (var row in rows)
                {
                    if (row.CatchNo != null && changedCatchNos.Contains(row.CatchNo) && row.Steps > resetStep)
                    {
                        row.Steps = resetStep;
                    }
                }

                // Reset other records in the affected lots
                var lotRecordsToReset = await _context.NRDatas
                    .Where(x => x.ProjectId == request.ProjectId &&
                                x.LotNo >= 0 &&
                                affectedLotNos.Contains(x.LotNo) &&
                                (x.CatchNo == null || !changedCatchNos.Contains(x.CatchNo)) &&
                                x.Status &&
                                x.Steps > resetStep)
                    .ToListAsync();

                foreach (var row in lotRecordsToReset)
                {
                    row.Steps = resetStep;
                }
            }

            await _context.SaveChangesAsync();
            await _loggerService.LogEventAsync(
                $"Updated LotNo for {rows.Count} NRData rows (ProjectId {request.ProjectId}). Moved catches: {string.Join(", ", changedCatchNos)}.",
                "NRDataLots",
                LogHelper.GetTriggeredBy(User),
                request.ProjectId
            );

            return Ok(new
            {
                message = "Lot numbers updated successfully.",
                updatedCount = rows.Count,
                movedCatchesCount = changedCatchNos.Count
            });
        }
    }
}

