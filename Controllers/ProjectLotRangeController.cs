using ERPToolsAPI.Data;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Tools.Models;
using System.Globalization;
using Microsoft.EntityFrameworkCore;

namespace Tools.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ProjectLotRangeController : ControllerBase
    {
        private readonly ERPToolsDbContext _context;

        public ProjectLotRangeController(ERPToolsDbContext context)
        {
            _context = context;
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

                // Create the model with formatted date strings
                var projectLotRange = new ProjectLotRange
                {
                    ProjectId = model.ProjectId,
                    StartDate = startDateString,
                    EndDate = endDateString,
                    LotNo = model.LotNo
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
                        projectLotRange.LotNo
                    }
                });
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
    }
}
