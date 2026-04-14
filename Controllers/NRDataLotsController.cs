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

        public NRDataLotsController(ERPToolsDbContext context, ILoggerService loggerService)
        {
            _context = context;
            _loggerService = loggerService;
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

            await _context.SaveChangesAsync();
            _loggerService.LogEvent(
                $"Updated LotNo for {rows.Count} NRData rows (ProjectId {request.ProjectId})",
                "NRDataLots",
                LogHelper.GetTriggeredBy(User),
                request.ProjectId
            );

            return Ok(new
            {
                message = "Lot numbers updated successfully.",
                updatedCount = rows.Count
            });
        }
    }
}

