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

        public class UpdateTemporaryDataDto
        {
            public int Id { get; set; }
            public int ProjectId { get; set; }
            public string? CourseName { get; set; }
            public string? SubjectName { get; set; }
            public System.Text.Json.JsonElement? CenterCode { get; set; }
            public int CollegeCode { get; set; }
            public string? CollegeName { get; set; }
            public int NRQuantity { get; set; }
            public System.Text.Json.JsonElement? CatchNo { get; set; }
            public string? ExamDate { get; set; }
            public string? ExamTime { get; set; }
            public string? Day { get; set; }
            public System.Text.Json.JsonElement? NodalCode { get; set; }
        }

        private static string? JsonElementToString(System.Text.Json.JsonElement? element)
        {
            if (!element.HasValue) return null;
            var e = element.Value;
            if (e.ValueKind == System.Text.Json.JsonValueKind.Null || e.ValueKind == System.Text.Json.JsonValueKind.Undefined) return null;
            if (e.ValueKind == System.Text.Json.JsonValueKind.Number) return e.GetRawText();
            if (e.ValueKind == System.Text.Json.JsonValueKind.String) return e.GetString();
            return e.ToString();
        }

        [HttpPut("{id}")]
        public async Task<IActionResult> UpdateTemporaryData(int id, [FromBody] UpdateTemporaryDataDto updatedRecord)
        {
            if (id != updatedRecord.Id && updatedRecord.Id != 0)
            {
                return BadRequest("ID mismatch");
            }

            var existingRecord = await _context.TemporaryNrDatas.FindAsync(id);
            if (existingRecord == null) return NotFound("Record not found");

            var cCode = JsonElementToString(updatedRecord.CenterCode);
            var nCode = JsonElementToString(updatedRecord.NodalCode);
            var catchNo = JsonElementToString(updatedRecord.CatchNo);

            if (cCode != null) existingRecord.CenterCode = cCode;
            if (nCode != null) existingRecord.NodalCode = nCode;
            if (catchNo != null) existingRecord.CatchNo = catchNo;

            if (updatedRecord.CollegeCode != 0) existingRecord.CollegeCode = updatedRecord.CollegeCode;
            if (updatedRecord.CollegeName != null) existingRecord.CollegeName = updatedRecord.CollegeName;
            if (updatedRecord.CourseName != null) existingRecord.CourseName = updatedRecord.CourseName;
            if (updatedRecord.SubjectName != null) existingRecord.SubjectName = updatedRecord.SubjectName;
            if (updatedRecord.NRQuantity != 0) existingRecord.NRQuantity = updatedRecord.NRQuantity;
            if (updatedRecord.ExamDate != null) existingRecord.ExamDate = updatedRecord.ExamDate;
            if (updatedRecord.ExamTime != null) existingRecord.ExamTime = updatedRecord.ExamTime;
            if (updatedRecord.Day != null) existingRecord.Day = updatedRecord.Day;

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
