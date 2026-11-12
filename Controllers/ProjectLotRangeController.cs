using ERPToolsAPI.Data;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Tools.Models;
using System.Globalization;
using Microsoft.EntityFrameworkCore;
using Tools.Services;

namespace Tools.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ProjectLotRangeController : ControllerBase
    {
        private readonly ERPToolsDbContext _context;
        private readonly IDispatchService _dispatchService;

        public ProjectLotRangeController(ERPToolsDbContext context, IDispatchService dispatchService)
        {
            _context = context;
            _dispatchService = dispatchService;
        }

        [HttpGet("{projectId}")]
        [HttpGet("by-project/{projectId}")]
        public async Task<IActionResult> GetProjectLotRanges(int projectId)
        {
            if (projectId <= 0)
                return BadRequest("ProjectId is required.");

            try
            {
                // Automatically sync any dispatched lots so status=0 in ProjectLotRanges
                await _dispatchService.SyncDispatchedLotRangesForProjectAsync(projectId);

                var ranges = await _context.ProjectLotRanges
                    .Where(x => x.ProjectId == projectId)
                    .OrderBy(x => x.LotNo)
                    .Select(x => new
                    {
                        x.LotRangeId,
                        x.ProjectId,
                        startDate = x.StartDate,
                        endDate = x.EndDate,
                        x.LotNo,
                        status = x.Status ? 1 : 0
                    })
                    .ToListAsync();

                return Ok(ranges);
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { message = "Failed to fetch project lot ranges", error = ex.Message });
            }
        }

        [HttpPost("sync-dispatch/{projectId}")]
        public async Task<IActionResult> SyncDispatch(int projectId)
        {
            if (projectId <= 0)
                return BadRequest("ProjectId is required.");

            try
            {
                var updatedLots = await _dispatchService.SyncDispatchedLotRangesForProjectAsync(projectId);
                return Ok(new
                {
                    message = "Dispatch status synced with ProjectLotRanges successfully.",
                    updatedLotsCount = updatedLots.Count,
                    updatedLots = updatedLots
                });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { message = "Failed to sync dispatch status", error = ex.Message });
            }
        }

        [HttpPost]
        public async Task<IActionResult> PostProjectLotRange([FromBody] ProjectLotRangeDto model)
        {
            try
            {
                // Validate input
                if (model == null)
                    return BadRequest(new { message = "Request body cannot be null." });

                if (model.ProjectId <= 0)
                    return BadRequest(new { message = "ProjectId must be greater than 0." });

                if (string.IsNullOrWhiteSpace(model.StartDate))
                    return BadRequest(new { message = "StartDate is required." });

                if (string.IsNullOrWhiteSpace(model.EndDate))
                    return BadRequest(new { message = "EndDate is required." });

                if (model.LotNo <= 0)
                    return BadRequest(new { message = "LotNo must be greater than 0." });

                // Check if this lot's dispatch date is exceeded or dispatch is done
                var isDispatched = await _dispatchService.SyncAndValidateLotDispatchStatusAsync(model.ProjectId, model.LotNo);
                if (isDispatched)
                {
                    return BadRequest(new
                    {
                        message = $"Lot {model.LotNo} has already been dispatched and is read-only. Modification is not allowed."
                    });
                }

                // Parse dates from string format "DD-MM-YYYY"
                if (!DateTime.TryParseExact(model.StartDate.Trim(), "dd-MM-yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime startDate))
                {
                    return BadRequest(new { message = "Invalid StartDate format. Expected dd-MM-yyyy." });
                }

                if (!DateTime.TryParseExact(model.EndDate.Trim(), "dd-MM-yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime endDate))
                {
                    return BadRequest(new { message = "Invalid EndDate format. Expected dd-MM-yyyy." });
                }

                // Validate date range
                if (startDate > endDate)
                {
                    return BadRequest(new { message = "StartDate cannot be greater than EndDate." });
                }

                // Convert DateTime back to DD-MM-YYYY string format for storage
                string startDateString = startDate.ToString("dd-MM-yyyy");
                string endDateString = endDate.ToString("dd-MM-yyyy");

                Console.WriteLine($"[ProjectLotRangeController] Saving - ProjectId: {model.ProjectId}, StartDate: {startDateString}, EndDate: {endDateString}, LotNo: {model.LotNo}");

                // Check if a record with the same ProjectId and LotNo exists
                var existingLotRange = await _context.ProjectLotRanges
                    .FirstOrDefaultAsync(x => x.ProjectId == model.ProjectId && x.LotNo == model.LotNo);

                if (existingLotRange != null)
                {
                    if (!existingLotRange.Status)
                    {
                        return BadRequest(new
                        {
                            message = $"Lot {model.LotNo} has been marked as inactive/dispatched and cannot be modified."
                        });
                    }

                    // Update existing record
                    existingLotRange.StartDate = startDateString;
                    existingLotRange.EndDate = endDateString;
                    if (model.Status.HasValue)
                    {
                        existingLotRange.Status = model.Status.Value;
                    }
                    
                    _context.ProjectLotRanges.Update(existingLotRange);
                    await _context.SaveChangesAsync();
                    
                    Console.WriteLine($"[ProjectLotRangeController] Successfully updated LotRangeId: {existingLotRange.LotRangeId}");
                    
                    return Ok(new
                    {
                        message = "Project Lot Range updated successfully.",
                        data = new
                        {
                            existingLotRange.LotRangeId,
                            existingLotRange.ProjectId,
                            startDate = existingLotRange.StartDate,
                            endDate = existingLotRange.EndDate,
                            existingLotRange.LotNo,
                            status = existingLotRange.Status
                        }
                    });
                }
                else
                {
                    // Create the model with formatted date strings
                    var projectLotRange = new ProjectLotRange
                    {
                        ProjectId = model.ProjectId,
                        StartDate = startDateString,
                        EndDate = endDateString,
                        LotNo = model.LotNo,
                        Status = model.Status ?? true
                    };

                    _context.ProjectLotRanges.Add(projectLotRange);
                    await _context.SaveChangesAsync();

                    Console.WriteLine($"[ProjectLotRangeController] Successfully saved with LotRangeId: {projectLotRange.LotRangeId}");

                    return Ok(new
                    {
                        message = "Project Lot Range added successfully.",
                        data = new
                        {
                            projectLotRange.LotRangeId,
                            projectLotRange.ProjectId,
                            startDate = projectLotRange.StartDate,
                            endDate = projectLotRange.EndDate,
                            projectLotRange.LotNo,
                            status = projectLotRange.Status
                        }
                    });
                }
            }
            catch (DbUpdateException dbEx)
            {
                Console.WriteLine($"[ProjectLotRangeController] DbUpdateException: {dbEx}");
                Console.WriteLine($"[ProjectLotRangeController] Inner Exception: {dbEx.InnerException}");
                
                return StatusCode(500, new
                {
                    message = "Database error while saving the Project Lot Range.",
                    error = dbEx.InnerException?.Message ?? dbEx.Message
                });
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ProjectLotRangeController] Exception: {ex}");
                Console.WriteLine($"[ProjectLotRangeController] Inner Exception: {ex.InnerException}");
                
                return StatusCode(500, new
                {
                    message = "An error occurred while saving the Project Lot Range.",
                    error = ex.Message,
                    innerError = ex.InnerException?.Message
                });
            }
        }
    }

    // DTO for receiving the request
    public class ProjectLotRangeDto
    {
        public int ProjectId { get; set; }
        public string StartDate { get; set; }  // String format "DD-MM-YYYY"
        public string EndDate { get; set; }    // String format "DD-MM-YYYY"
        public int LotNo { get; set; }
        public bool? Status { get; set; }
    }
}
