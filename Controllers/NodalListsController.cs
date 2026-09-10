using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Text.Json;
using System.Reflection;
using Tools.Models;
using ERPToolsAPI.Data;

namespace Tools.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class NodalListsController : ControllerBase
    {
        private readonly ERPToolsDbContext _context;

        public NodalListsController(ERPToolsDbContext context)
        {
            _context = context;
        }

        [HttpGet("{projectId}")]
        public async Task<IActionResult> GetNodalLists(int projectId, [FromQuery] int pageNo = 1, [FromQuery] int pageSize = 10, [FromQuery] string? search = null, [FromQuery] string? sortField = null, [FromQuery] string? sortOrder = null, [FromQuery] string? columnFilters = null)
        {
            try
            {
                var query = _context.NodalList.Where(x => x.ProjectId == projectId).AsQueryable();

                if (!string.IsNullOrEmpty(search))
                {
                    search = search.ToLower();
                    query = query.Where(x => 
                        x.CollegeCode.ToString().Contains(search) ||
                        (x.CollegeName != null && x.CollegeName.ToLower().Contains(search)) ||
                        x.ExamCenterCode.ToString().Contains(search) ||
                        (x.ExamCenterName != null && x.ExamCenterName.ToLower().Contains(search)) ||
                        (x.Gender != null && x.Gender.ToLower().Contains(search)) ||
                        x.NodalCode.ToString().Contains(search) ||
                        (x.NodalName != null && x.NodalName.ToLower().Contains(search)) ||
                        (x.OtherFields != null && x.OtherFields.ToLower().Contains(search))
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
                                    "collegecode" => query.Where(x => x.CollegeCode.ToString().Contains(val)),
                                    "collegename" => query.Where(x => x.CollegeName != null && x.CollegeName.ToLower().Contains(val)),
                                    "examcentercode" => query.Where(x => x.ExamCenterCode.ToString().Contains(val)),
                                    "examcentername" => query.Where(x => x.ExamCenterName != null && x.ExamCenterName.ToLower().Contains(val)),
                                    "gender" => query.Where(x => x.Gender != null && x.Gender.ToLower().Contains(val)),
                                    "nodalcode" => query.Where(x => x.NodalCode.ToString().Contains(val)),
                                    "nodalname" => query.Where(x => x.NodalName != null && x.NodalName.ToLower().Contains(val)),
                                    _ => query.Where(x => x.OtherFields != null && x.OtherFields.ToLower().Contains(val))
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
                        "collegecode" => isAsc ? query.OrderBy(x => x.CollegeCode) : query.OrderByDescending(x => x.CollegeCode),
                        "collegename" => isAsc ? query.OrderBy(x => x.CollegeName) : query.OrderByDescending(x => x.CollegeName),
                        "examcentercode" => isAsc ? query.OrderBy(x => x.ExamCenterCode) : query.OrderByDescending(x => x.ExamCenterCode),
                        "examcentername" => isAsc ? query.OrderBy(x => x.ExamCenterName) : query.OrderByDescending(x => x.ExamCenterName),
                        "nodalcode" => isAsc ? query.OrderBy(x => x.NodalCode) : query.OrderByDescending(x => x.NodalCode),
                        "nodalname" => isAsc ? query.OrderBy(x => x.NodalName) : query.OrderByDescending(x => x.NodalName),
                        "gender" => isAsc ? query.OrderBy(x => x.Gender) : query.OrderByDescending(x => x.Gender),
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
                var innerMessage = ex.InnerException != null ? ex.InnerException.Message : ""; if (ex.InnerException?.InnerException != null) innerMessage += " | " + ex.InnerException.InnerException.Message; return StatusCode(500, $"Internal server error: {ex.Message}. Inner: {innerMessage}");
            }
        }

        [HttpPut("{id}")]
        public async Task<IActionResult> UpdateNodalList(int id, [FromBody] NodalList updatedRecord)
        {
            if (id != updatedRecord.Id) return BadRequest("ID mismatch");

            var existingRecord = await _context.NodalList.FindAsync(id);
            if (existingRecord == null) return NotFound("Record not found");

            existingRecord.CollegeCode = updatedRecord.CollegeCode;
            existingRecord.CollegeName = updatedRecord.CollegeName;
            existingRecord.ExamCenterCode = updatedRecord.ExamCenterCode;
            existingRecord.ExamCenterName = updatedRecord.ExamCenterName;
            existingRecord.Gender = updatedRecord.Gender;
            existingRecord.NodalCode = updatedRecord.NodalCode;
            existingRecord.NodalName = updatedRecord.NodalName;

            try
            {
                await _context.SaveChangesAsync();
                return Ok(existingRecord);
            }
            catch (Exception ex)
            {
                var innerMessage = ex.InnerException != null ? ex.InnerException.Message : ""; if (ex.InnerException?.InnerException != null) innerMessage += " | " + ex.InnerException.InnerException.Message; return StatusCode(500, $"Internal server error: {ex.Message}. Inner: {innerMessage}");
            }
        }

        [HttpPost("Upload")]
        public async Task<IActionResult> UploadNodalList([FromBody] JsonElement inputData)
        {
            if (!ModelState.IsValid)
                return BadRequest(ModelState);

            try
            {
                var projProp = inputData.GetProperty("projectId");
                int projectId = projProp.ValueKind == JsonValueKind.String ? int.Parse(projProp.GetString()!) : projProp.GetInt32();
                var incomingData = inputData.GetProperty("data");

                var properties = typeof(NodalList)
                    .GetProperties()
                    .ToDictionary(p => p.Name.ToLower(), p => p);

                var nodalListsToAdd = new List<NodalList>();

                for (int i = 0; i < incomingData.GetArrayLength(); i++)
                {
                    var item = incomingData[i];
                    var nodalList = new NodalList { ProjectId = projectId };
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
                                    propInfo.SetValue(nodalList, convertedValue);
                                }
                                else if (!propInfo.PropertyType.IsValueType)
                                {
                                    propInfo.SetValue(nodalList, null);
                                }
                            }
                            catch { }
                        }
                        else if (prop.Name != "projectId")
                        {
                            extraData[prop.Name] = value;
                        }
                    }

                    if (extraData.Any()) nodalList.OtherFields = JsonSerializer.Serialize(extraData);
                    nodalListsToAdd.Add(nodalList);
                }

                if (nodalListsToAdd.Any())
                {
                    await _context.NodalList.AddRangeAsync(nodalListsToAdd);
                    await _context.SaveChangesAsync();
                }

                return Ok(new
                {
                    message = "Nodal List data uploaded successfully",
                    NewRecords = nodalListsToAdd.Count
                });
            }
            catch (Exception ex)
            {
                var innerMessage = ex.InnerException != null ? ex.InnerException.Message : ""; if (ex.InnerException?.InnerException != null) innerMessage += " | " + ex.InnerException.InnerException.Message; return StatusCode(500, $"Internal server error: {ex.Message}. Inner: {innerMessage}");
            }
        }
    }
}
