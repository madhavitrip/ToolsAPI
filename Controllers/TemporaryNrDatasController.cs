using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Text.Json;
using Tools.Models;
using ERPToolsAPI.Data;

namespace Tools.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class TemporaryNrDatasController : ControllerBase
    {
        private readonly ERPToolsDbContext _context;

        public TemporaryNrDatasController(ERPToolsDbContext context)
        {
            _context = context;
        }

        [HttpGet("{projectId}")]
        public async Task<IActionResult> GetTemporaryData(int projectId, [FromQuery] int pageNo = 1, [FromQuery] int pageSize = 10, [FromQuery] string? search = null, [FromQuery] string? sortField = null, [FromQuery] string? sortOrder = null, [FromQuery] string? columnFilters = null)
        {
            try
            {
                var query = _context.TemporaryNrDatas.Where(x => x.ProjectId == projectId).AsQueryable();

                if (!string.IsNullOrEmpty(search))
                {
                    search = search.ToLower();
                    query = query.Where(x => 
                        (x.CatchNo != null && x.CatchNo.ToLower().Contains(search)) ||
                        (x.CourseName != null && x.CourseName.ToLower().Contains(search)) ||
                        (x.SubjectName != null && x.SubjectName.ToLower().Contains(search)) ||
                        (x.CenterCode != null && x.CenterCode.ToLower().Contains(search)) ||
                        (x.NodalCode != null && x.NodalCode.ToLower().Contains(search)) ||
                        x.CollegeCode.ToString().Contains(search) ||
                        (x.CollegeName != null && x.CollegeName.ToLower().Contains(search))
                    );
                }

                if (!string.IsNullOrEmpty(columnFilters))
                {
                    try
                    {
                        var filters = JsonSerializer.Deserialize<Dictionary<string, string>>(columnFilters);
                        if (filters != null)
                        {
                            foreach (var filter in filters)
                            {
                                var key = filter.Key.ToLower();
                                var val = filter.Value.ToLower();

                                query = key switch
                                {
                                    "catchno" => query.Where(x => x.CatchNo != null && x.CatchNo.ToLower().Contains(val)),
                                    "centercode" => query.Where(x => x.CenterCode != null && x.CenterCode.ToLower().Contains(val)),
                                    "nodalcode" => query.Where(x => x.NodalCode != null && x.NodalCode.ToLower().Contains(val)),
                                    "collegecode" => query.Where(x => x.CollegeCode.ToString().Contains(val)),
                                    "collegename" => query.Where(x => x.CollegeName != null && x.CollegeName.ToLower().Contains(val)),
                                    "coursename" => query.Where(x => x.CourseName != null && x.CourseName.ToLower().Contains(val)),
                                    "subjectname" => query.Where(x => x.SubjectName != null && x.SubjectName.ToLower().Contains(val)),
                                    "examdate" => query.Where(x => x.ExamDate != null && x.ExamDate.ToLower().Contains(val)),
                                    "examtime" => query.Where(x => x.ExamTime != null && x.ExamTime.ToLower().Contains(val)),
                                    _ => query
                                };
                            }
                        }
                    }
                    catch { }
                }

                if (!string.IsNullOrEmpty(sortField))
                {
                    bool isAsc = string.IsNullOrEmpty(sortOrder) || sortOrder.ToLower() != "descend";
                    query = sortField.ToLower() switch
                    {
                        "catchno" => isAsc ? query.OrderBy(x => x.CatchNo) : query.OrderByDescending(x => x.CatchNo),
                        "centercode" => isAsc ? query.OrderBy(x => x.CenterCode) : query.OrderByDescending(x => x.CenterCode),
                        "nodalcode" => isAsc ? query.OrderBy(x => x.NodalCode) : query.OrderByDescending(x => x.NodalCode),
                        "collegecode" => isAsc ? query.OrderBy(x => x.CollegeCode) : query.OrderByDescending(x => x.CollegeCode),
                        "collegename" => isAsc ? query.OrderBy(x => x.CollegeName) : query.OrderByDescending(x => x.CollegeName),
                        "nrquantity" => isAsc ? query.OrderBy(x => x.NRQuantity) : query.OrderByDescending(x => x.NRQuantity),
                        _ => query.OrderBy(x => x.Id)
                    };
                }
                else
                {
                    query = query.OrderBy(x => x.Id);
                }

                var totalCount = await query.CountAsync();
                var items = await query.Skip((pageNo - 1) * pageSize).Take(pageSize).ToListAsync();

                return Ok(new { items, totalCount });
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
        }

        [HttpPut("{id}")]
        public async Task<IActionResult> UpdateTemporaryData(int id, [FromBody] TemporaryNrDatas updatedRecord)
        {
            if (id != updatedRecord.Id) return BadRequest("ID mismatch");

            var existingRecord = await _context.TemporaryNrDatas.FindAsync(id);
            if (existingRecord == null) return NotFound("Record not found");

            // Update allowed fields
            existingRecord.CenterCode = updatedRecord.CenterCode;
            existingRecord.NodalCode = updatedRecord.NodalCode;
            existingRecord.CollegeCode = updatedRecord.CollegeCode;
            existingRecord.CollegeName = updatedRecord.CollegeName;
            existingRecord.CourseName = updatedRecord.CourseName;
            existingRecord.SubjectName = updatedRecord.SubjectName;
            existingRecord.CatchNo = updatedRecord.CatchNo;
            existingRecord.NRQuantity = updatedRecord.NRQuantity;
            existingRecord.ExamDate = updatedRecord.ExamDate;
            existingRecord.ExamTime = updatedRecord.ExamTime;
            existingRecord.Day = updatedRecord.Day;

            try
            {
                await _context.SaveChangesAsync();
                return Ok(existingRecord);
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
        }

        [HttpDelete("{id}")]
        public async Task<IActionResult> DeleteTemporaryData(int id)
        {
            try
            {
                var record = await _context.TemporaryNrDatas.FindAsync(id);
                if (record == null)
                {
                    return NotFound(new { message = "Record not found" });
                }

                _context.TemporaryNrDatas.Remove(record);
                await _context.SaveChangesAsync();

                return Ok(new { message = "Record deleted successfully", id });
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
        }

        [HttpDelete("deleteAll/{projectId}")]
        public async Task<IActionResult> DeleteAllTemporaryData(int projectId)
        {
            try
            {
                var affected = await _context.TemporaryNrDatas
                    .Where(x => x.ProjectId == projectId)
                    .ExecuteDeleteAsync();

                return Ok(new { message = "All temporary preview records deleted successfully", count = affected });
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
        }

        [HttpPost("batch-delete")]
        public async Task<IActionResult> BatchDeleteTemporaryData([FromBody] List<int> ids)
        {
            if (ids == null || !ids.Any()) return BadRequest("No records selected for deletion.");
            try
            {
                var affected = await _context.TemporaryNrDatas
                    .Where(x => ids.Contains(x.Id))
                    .ExecuteDeleteAsync();

                return Ok(new { message = $"{affected} record(s) deleted successfully", count = affected });
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
        }
    }
}
