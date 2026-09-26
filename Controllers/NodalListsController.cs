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
        private static bool _schemaEnsured = false;
        private static readonly object _schemaLock = new object();

        public NodalListsController(ERPToolsDbContext context)
        {
            _context = context;
        }

        private async Task EnsureSchemaAsync()
        {
            if (_schemaEnsured) return;
            lock (_schemaLock)
            {
                if (_schemaEnsured) return;
            }

            try
            {
                await _context.Database.ExecuteSqlRawAsync("ALTER TABLE NodalList ADD COLUMN Status TINYINT(1) NOT NULL DEFAULT 1;");
            }
            catch
            {
                // Column may already exist
            }

            lock (_schemaLock)
            {
                _schemaEnsured = true;
            }
        }

        [HttpGet("{projectId}")]
        public async Task<IActionResult> GetNodalLists(int projectId, [FromQuery] int pageNo = 1, [FromQuery] int pageSize = 10, [FromQuery] string? search = null, [FromQuery] string? sortField = null, [FromQuery] string? sortOrder = null, [FromQuery] string? columnFilters = null)
        {
            try
            {
                await EnsureSchemaAsync();
                var query = _context.NodalList.Where(x => x.ProjectId == projectId && x.Status).AsQueryable();

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

            var existingRecord = await _context.NodalList.FirstOrDefaultAsync(x => x.Id == id && x.Status);
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

        [HttpDelete("{id}")]
        public async Task<IActionResult> DeleteNodalList(int id)
        {
            try
            {
                await EnsureSchemaAsync();
                var record = await _context.NodalList.FindAsync(id);
                if (record == null || !record.Status)
                {
                    return NotFound(new { message = "Record not found" });
                }

                record.Status = false;
                await _context.SaveChangesAsync();

                return Ok(new { message = "Record deleted successfully", id });
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
        }

        [HttpDelete("project/{projectId}")]
        [HttpDelete("deleteAll/{projectId}")]
        [HttpPost("delete-all/{projectId}")]
        public async Task<IActionResult> DeleteNodalListByProject(int projectId)
        {
            try
            {
                await EnsureSchemaAsync();
                var affected = await _context.Database.ExecuteSqlRawAsync(
                    "UPDATE NodalList SET Status = 0 WHERE ProjectId = {0} AND Status = 1;", projectId);

                return Ok(new { message = "All records soft-deleted successfully", count = affected });
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
        }

        [HttpPost("batch-delete")]
        public async Task<IActionResult> BatchDeleteNodalList([FromBody] List<int> ids)
        {
            if (ids == null || !ids.Any()) return BadRequest("No IDs provided");
            try
            {
                await EnsureSchemaAsync();
                var records = await _context.NodalList.Where(x => ids.Contains(x.Id) && x.Status).ToListAsync();
                foreach (var record in records)
                {
                    record.Status = false;
                }
                await _context.SaveChangesAsync();
                return Ok(new { message = $"{records.Count} records deleted successfully", count = records.Count });
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
        }

        public class ResolveCollegeCenterDto
        {
            public int ProjectId { get; set; }
            public int CollegeCode { get; set; }
            public string? CollegeName { get; set; }
            public string? Gender { get; set; }
            public int CorrectCenterCode { get; set; }
            public string? CorrectCenterName { get; set; }
            public int NodalCode { get; set; }
            public string? NodalName { get; set; }
        }

        [HttpPost("resolve-college-center")]
        public async Task<IActionResult> ResolveCollegeCenter([FromBody] ResolveCollegeCenterDto req)
        {
            try
            {
                await EnsureSchemaAsync();
                var allNodals = await _context.NodalList
                    .Where(x => x.ProjectId == req.ProjectId && x.Status)
                    .ToListAsync();

                static string GetEffectiveCollegeCode(NodalList n)
                {
                    if (n.CollegeCode != 0) return n.CollegeCode.ToString();
                    if (!string.IsNullOrEmpty(n.CollegeName))
                    {
                        var m = System.Text.RegularExpressions.Regex.Match(n.CollegeName, @"^(\d+)");
                        if (m.Success) return m.Groups[1].Value;
                        return n.CollegeName.Trim().ToLowerInvariant();
                    }
                    return $"row_{n.Id}";
                }

                var reqCodeStr = req.CollegeCode.ToString();
                var records = allNodals.Where(n => {
                    var cCode = GetEffectiveCollegeCode(n);
                    if (req.CollegeCode != 0 && cCode == reqCodeStr) return true;
                    if (req.CollegeCode != 0 && n.CollegeCode == req.CollegeCode) return true;
                    if (!string.IsNullOrEmpty(req.CollegeName) && !string.IsNullOrEmpty(n.CollegeName) && (n.CollegeName.Trim().Equals(req.CollegeName.Trim(), StringComparison.OrdinalIgnoreCase) || (req.CollegeCode != 0 && n.CollegeName.StartsWith(reqCodeStr)))) return true;
                    return false;
                }).ToList();

                var centerNameStr = !string.IsNullOrWhiteSpace(req.CorrectCenterName) ? req.CorrectCenterName : $"Center {req.CorrectCenterCode}";
                var collegeNameStr = !string.IsNullOrWhiteSpace(req.CollegeName) ? req.CollegeName : (req.CollegeCode != 0 ? $"College {req.CollegeCode}" : "Unknown College");
                var nodalCodeInt = req.NodalCode != 0 ? req.NodalCode : (req.CorrectCenterCode != 0 ? req.CorrectCenterCode : 1);
                var nodalNameStr = !string.IsNullOrWhiteSpace(req.NodalName) ? req.NodalName : $"Nodal {nodalCodeInt}";

                if (!records.Any())
                {
                    var newRecord = new NodalList
                    {
                        ProjectId = req.ProjectId,
                        CollegeCode = req.CollegeCode,
                        CollegeName = collegeNameStr,
                        ExamCenterCode = req.CorrectCenterCode,
                        ExamCenterName = centerNameStr,
                        NodalCode = nodalCodeInt,
                        NodalName = nodalNameStr,
                        Gender = !string.IsNullOrWhiteSpace(req.Gender) ? req.Gender : "ALL",
                        Status = true
                    };
                    _context.NodalList.Add(newRecord);
                }
                else
                {
                    foreach (var r in records)
                    {
                        if (r.CollegeCode == 0 && req.CollegeCode != 0) r.CollegeCode = req.CollegeCode;
                        r.ExamCenterCode = req.CorrectCenterCode;
                        r.ExamCenterName = centerNameStr;
                        if (string.IsNullOrWhiteSpace(r.CollegeName)) r.CollegeName = collegeNameStr;
                        if (r.NodalCode == 0) r.NodalCode = nodalCodeInt;
                        if (string.IsNullOrWhiteSpace(r.NodalName)) r.NodalName = nodalNameStr;
                    }
                }

                // Update TemporaryNrDatas staging table for matching records
                var tempDatasToUpdate = await _context.TemporaryNrDatas
                    .Where(x => x.ProjectId == req.ProjectId)
                    .ToListAsync();

                foreach (var t in tempDatasToUpdate)
                {
                    bool match = (req.CollegeCode != 0 && t.CollegeCode == req.CollegeCode) ||
                                 (!string.IsNullOrWhiteSpace(req.CollegeName) && !string.IsNullOrWhiteSpace(t.CollegeName) && t.CollegeName.Trim().Equals(req.CollegeName.Trim(), StringComparison.OrdinalIgnoreCase)) ||
                                 (!string.IsNullOrWhiteSpace(t.CenterCode) && t.CenterCode == reqCodeStr) ||
                                 (!string.IsNullOrWhiteSpace(t.CollegeName) && t.CollegeName.Contains(reqCodeStr));

                    if (match)
                    {
                        t.CenterCode = req.CorrectCenterCode.ToString();
                        t.NodalCode = nodalCodeInt.ToString();
                        if (t.CollegeCode == 0 && req.CollegeCode != 0) t.CollegeCode = req.CollegeCode;
                        if (string.IsNullOrWhiteSpace(t.CollegeName) && !string.IsNullOrWhiteSpace(collegeNameStr)) t.CollegeName = collegeNameStr;
                    }
                }

                var conflictsToResolve = await _context.ConflictingFields
                    .Where(c => c.ProjectId == req.ProjectId && c.Status == 1 &&
                               (c.UniqueField.Contains($"Rule1_MultiCenter_{reqCodeStr}") || c.UniqueField.Contains($"Rule2_Unassigned_{reqCodeStr}") || c.UniqueField.Contains(reqCodeStr)))
                    .ToListAsync();
                foreach (var conf in conflictsToResolve)
                {
                    conf.Status = 0;
                }

                await _context.SaveChangesAsync();
                return Ok(new { message = $"Updated records to Center {req.CorrectCenterCode} & Nodal {nodalCodeInt} in Nodal List and Temporary Staging", count = records.Count });
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
        }

        public class ResolveCenterNodalDto
        {
            public int ProjectId { get; set; }
            public int CenterCode { get; set; }
            public int CorrectNodalCode { get; set; }
            public string? CorrectNodalName { get; set; }
        }

        [HttpPost("resolve-center-nodal")]
        public async Task<IActionResult> ResolveCenterNodal([FromBody] ResolveCenterNodalDto req)
        {
            try
            {
                await EnsureSchemaAsync();
                var records = await _context.NodalList
                    .Where(x => x.ProjectId == req.ProjectId && x.ExamCenterCode == req.CenterCode && x.Status)
                    .ToListAsync();

                if (!records.Any()) return NotFound(new { message = "No matching records found to update" });

                foreach (var r in records)
                {
                    r.NodalCode = req.CorrectNodalCode;
                    if (!string.IsNullOrEmpty(req.CorrectNodalName))
                    {
                        r.NodalName = req.CorrectNodalName;
                    }
                }

                var centerCodeStr = req.CenterCode.ToString();

                // Update TemporaryNrDatas staging table for matching records
                var tempDatasToUpdate = await _context.TemporaryNrDatas
                    .Where(x => x.ProjectId == req.ProjectId)
                    .ToListAsync();

                foreach (var t in tempDatasToUpdate)
                {
                    if (t.CenterCode == centerCodeStr)
                    {
                        t.NodalCode = req.CorrectNodalCode.ToString();
                    }
                }

                var nodalsToResolve = await _context.ConflictingFields
                    .Where(c => c.ProjectId == req.ProjectId && c.Status == 1 &&
                               (c.UniqueField.Contains($"Rule1_MultiNodal_{centerCodeStr}") || c.UniqueField.Contains(centerCodeStr)))
                    .ToListAsync();
                foreach (var conf in nodalsToResolve)
                {
                    conf.Status = 0;
                }

                await _context.SaveChangesAsync();
                return Ok(new { message = $"Updated records to Nodal {req.CorrectNodalCode} in Nodal List and Temporary Staging", count = records.Count });
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
        }



        [HttpPost]
        public async Task<IActionResult> CreateNodalList([FromBody] NodalList newRecord)
        {
            if (newRecord == null) return BadRequest("Invalid data");
            try
            {
                await EnsureSchemaAsync();
                newRecord.Status = true;
                await _context.NodalList.AddAsync(newRecord);
                await _context.SaveChangesAsync();
                return Ok(newRecord);
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
        }

        [HttpPost("Upload")]
        public async Task<IActionResult> UploadNodalList([FromBody] JsonElement inputData)
        {
            if (!ModelState.IsValid)
                return BadRequest(ModelState);

            try
            {
                await EnsureSchemaAsync();
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
