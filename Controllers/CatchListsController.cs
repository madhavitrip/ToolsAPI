using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Text.Json;
using System.Reflection;
using Tools.Models;
using ERPToolsAPI.Data;
using ERPToolsAPI.Data;

namespace Tools.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class CatchListsController : ControllerBase
    {
        private readonly ERPToolsDbContext _context;

        public CatchListsController(ERPToolsDbContext context)
        {
            _context = context;
        }

        [HttpGet("{projectId}")]
        public async Task<IActionResult> GetCatchLists(int projectId, [FromQuery] int pageNo = 1, [FromQuery] int pageSize = 10, [FromQuery] string? search = null, [FromQuery] string? sortField = null, [FromQuery] string? sortOrder = null, [FromQuery] string? columnFilters = null)
        {
            try
            {
                var query = _context.CatchList.Where(x => x.ProjectId == projectId).AsQueryable();

                if (!string.IsNullOrEmpty(search))
                {
                    search = search.ToLower();
                    query = query.Where(x => 
                        (x.CatchNo != null && x.CatchNo.ToLower().Contains(search)) ||
                        (x.CollegeName != null && x.CollegeName.ToLower().Contains(search)) ||
                        (x.CourseName != null && x.CourseName.ToLower().Contains(search)) ||
                        (x.SubjectName != null && x.SubjectName.ToLower().Contains(search)) ||
                        (x.PaperCode != null && x.PaperCode.ToLower().Contains(search)) ||
                        x.CollegeCode.ToString().Contains(search) ||
                        (x.NRDatas != null && x.NRDatas.ToLower().Contains(search))
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
                                    "collegecode" => query.Where(x => x.CollegeCode.ToString().Contains(val)),
                                    "collegename" => query.Where(x => x.CollegeName != null && x.CollegeName.ToLower().Contains(val)),
                                    "papercode" => query.Where(x => x.PaperCode != null && x.PaperCode.ToLower().Contains(val)),
                                    "coursename" => query.Where(x => x.CourseName != null && x.CourseName.ToLower().Contains(val)),
                                    "subjectname" => query.Where(x => x.SubjectName != null && x.SubjectName.ToLower().Contains(val)),
                                    "nrquantity" => query.Where(x => x.NRQuantity.ToString().Contains(val)),
                                    "examdate" => query.Where(x => x.ExamDate != null && x.ExamDate.ToLower().Contains(val)),
                                    "examtime" => query.Where(x => x.ExamTime != null && x.ExamTime.ToLower().Contains(val)),
                                    "transgender" => query.Where(x => x.Transgender.ToString().Contains(val)),
                                    "male" => query.Where(x => x.Male.ToString().Contains(val)),
                                    "female" => query.Where(x => x.Female.ToString().Contains(val)),
                                    "semester" => query.Where(x => x.Semester != null && x.Semester.ToLower().Contains(val)),
                                    _ => query.Where(x => x.NRDatas != null && x.NRDatas.ToLower().Contains(val))
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
                        "collegecode" => isAsc ? query.OrderBy(x => x.CollegeCode) : query.OrderByDescending(x => x.CollegeCode),
                        "collegename" => isAsc ? query.OrderBy(x => x.CollegeName) : query.OrderByDescending(x => x.CollegeName),
                        "coursename" => isAsc ? query.OrderBy(x => x.CourseName) : query.OrderByDescending(x => x.CourseName),
                        "subjectname" => isAsc ? query.OrderBy(x => x.SubjectName) : query.OrderByDescending(x => x.SubjectName),
                        "papercode" => isAsc ? query.OrderBy(x => x.PaperCode) : query.OrderByDescending(x => x.PaperCode),
                        "nrquantity" => isAsc ? query.OrderBy(x => x.NRQuantity) : query.OrderByDescending(x => x.NRQuantity),
                        "examdate" => isAsc ? query.OrderBy(x => x.ExamDate) : query.OrderByDescending(x => x.ExamDate),
                        "examtime" => isAsc ? query.OrderBy(x => x.ExamTime) : query.OrderByDescending(x => x.ExamTime),
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
        public async Task<IActionResult> UpdateCatchList(int id, [FromBody] CatchList updatedRecord)
        {
            if (id != updatedRecord.Id) return BadRequest("ID mismatch");

            var existingRecord = await _context.CatchList.FindAsync(id);
            if (existingRecord == null) return NotFound("Record not found");

            existingRecord.CatchNo = updatedRecord.CatchNo;
            existingRecord.CollegeCode = updatedRecord.CollegeCode;
            existingRecord.CollegeName = updatedRecord.CollegeName;
            existingRecord.PaperCode = updatedRecord.PaperCode;
            existingRecord.CourseName = updatedRecord.CourseName;
            existingRecord.SubjectName = updatedRecord.SubjectName;
            existingRecord.NRQuantity = updatedRecord.NRQuantity;
            existingRecord.ExamDate = updatedRecord.ExamDate;
            existingRecord.ExamTime = updatedRecord.ExamTime;
            existingRecord.Transgender = updatedRecord.Transgender;
            existingRecord.Male = updatedRecord.Male;
            existingRecord.Female = updatedRecord.Female;
            existingRecord.Semester = updatedRecord.Semester;

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

        [HttpPost("Upload")]
        public async Task<IActionResult> UploadCatchList([FromBody] JsonElement inputData)
        {
            if (!ModelState.IsValid)
                return BadRequest(ModelState);

            try
            {
                var projProp = inputData.GetProperty("projectId");
                int projectId = projProp.ValueKind == JsonValueKind.String ? int.Parse(projProp.GetString()!) : projProp.GetInt32();
                var incomingData = inputData.GetProperty("data");

                var properties = typeof(CatchList)
                    .GetProperties()
                    .ToDictionary(p => p.Name.ToLower(), p => p);

                var catchListsToAdd = new List<CatchList>();

                for (int i = 0; i < incomingData.GetArrayLength(); i++)
                {
                    var item = incomingData[i];
                    var catchList = new CatchList { ProjectId = projectId };
                    var extraData = new Dictionary<string, string>();

                    foreach (var prop in item.EnumerateObject())
                    {
                        string key = prop.Name.Replace(" ", "").ToLower();
                        string value = prop.Value.ToString().Trim();

                        if (properties.TryGetValue(key, out var propInfo))
                        {
                            if (propInfo.Name.Equals("Id", StringComparison.OrdinalIgnoreCase) || 
                                propInfo.Name.Equals("ProjectId", StringComparison.OrdinalIgnoreCase)) 
                                continue;

                            try
                            {
                                var targetType = Nullable.GetUnderlyingType(propInfo.PropertyType) ?? propInfo.PropertyType;
                                object? convertedValue = null;

                                if (!string.IsNullOrWhiteSpace(value))
                                {
                                    convertedValue = Convert.ChangeType(value, targetType);
                                }

                                if (convertedValue != null)
                                {
                                    propInfo.SetValue(catchList, convertedValue);
                                }
                                else if (!propInfo.PropertyType.IsValueType)
                                {
                                    propInfo.SetValue(catchList, null);
                                }
                            }
                            catch { }
                        }
                        else if (prop.Name != "projectId")
                        {
                            extraData[prop.Name] = value;
                        }
                    }

                    if (extraData.Any()) catchList.NRDatas = JsonSerializer.Serialize(extraData);
                    catchListsToAdd.Add(catchList);
                }

                if (catchListsToAdd.Any())
                {
                    await _context.CatchList.AddRangeAsync(catchListsToAdd);
                    await _context.SaveChangesAsync();
                }

                return Ok(new
                {
                    message = "Catch List data uploaded successfully",
                    NewRecords = catchListsToAdd.Count
                });
            }
            catch (Exception ex)
            {
                var innerMessage = ex.InnerException != null ? ex.InnerException.Message : ""; if (ex.InnerException?.InnerException != null) innerMessage += " | " + ex.InnerException.InnerException.Message;
                return StatusCode(500, $"Internal server error: {ex.Message}. Inner: {innerMessage}");
            }
        }
    }
}
