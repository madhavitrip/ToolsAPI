using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using ERPToolsAPI.Data;
using Tools.Models;
using System.Text.Json;
using System.Reflection;
using Tools.Services;
using Microsoft.CodeAnalysis;

namespace Tools.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class NRDatasController : ControllerBase
    {
        private readonly ERPToolsDbContext _context;
        private readonly ILoggerService _loggerService;

        public NRDatasController(ERPToolsDbContext context, ILoggerService loggerService)
        {
            _context = context;
            _loggerService = loggerService;
        }

        // GET: api/NRDatas
        [HttpGet]
        public async Task<ActionResult<IEnumerable<NRData>>> GetNRDatas()
        {
            return await _context.NRDatas.ToListAsync();
        }

        // GET: api/NRDatas/5
        [HttpGet("{id}")]
        public async Task<ActionResult<NRData>> GetNRData(int id)
        {
            var nRData = await _context.NRDatas.FindAsync(id);

            if (nRData == null)
            {
                return NotFound();
            }

            return nRData;
        }

        // GET: api/NRDatas/GetByProjectId/5
        [HttpGet("GetByProjectId/{projectId}")]
        public async Task<ActionResult> GetByProjectId(int projectId, int pageSize, int pageNo, string?search=null, string? key = null)
        {
            IQueryable<NRData> query = _context.NRDatas
             .Where(d => d.ProjectId == projectId);

            // ⭐ APPLY SEARCH IF KEY + SEARCH PROVIDED
            if (!string.IsNullOrWhiteSpace(search) && !string.IsNullOrWhiteSpace(key))
            {
                search = search.ToLower();

                switch (key)
                {
                    case "CatchNo":
                        query = query.Where(d => d.CatchNo.ToLower().Contains(search));
                        break;
                    case "CenterCode":
                        query = query.Where(d => d.CenterCode.ToLower().Contains(search));
                        break;
                    case "SubjectName":
                        query = query.Where(d => d.SubjectName.ToLower().Contains(search));
                        break;
                    case "CourseName":
                        query = query.Where(d => d.CourseName.ToLower().Contains(search));
                        break;
                    case "NodalCode":
                        query = query.Where(d => d.NodalCode.ToLower().Contains(search));
                        break;
                    case "Route":
                        query = query.Where(d => d.Route.ToLower().Contains(search));
                        break;
                    case "ExamDate":
                        query = query.Where(d => d.ExamDate.ToString().Contains(search));
                        break;
                    case "ExamTime":
                        query = query.Where(d => d.ExamTime.ToString().Contains(search));
                        break;
                    case "NRQuantity":
                        query = query.Where(d => d.NRQuantity.ToString().Contains(search));
                        break;
                    case "Quantity":
                        query = query.Where(d => d.Quantity.ToString().Contains(search));
                        break;
                  
                    default:
                        return BadRequest($"Key '{key}' is not searchable.");
                }
            }
            int totalCount = await query.CountAsync();
            int totalPages = (int)Math.Ceiling(totalCount / (double)pageSize);

            var nrDataList = await query
                .OrderBy(d=>d.Id)
                .Skip((pageNo - 1) * pageSize)
                .Take(pageSize)
                .Select(d => new
                {
                    d.CatchNo,
                    d.CenterCode,
                    d.ExamDate,
                    d.ExamTime,
                    d.SubjectName,
                    d.CourseName,
                    d.NodalCode,
                    d.NRQuantity,
                    d.Quantity,
                    d.Route,
                    d.Pages,
                    d.CenterSort,
                    d.NodalSort,
                    d.RouteSort,
                    d.Symbol,

                })
                .ToListAsync();

            if (!nrDataList.Any())
            {
                var allColumns = typeof(NRData).GetProperties()
              .Select(p => p.Name)
               .Where(name => name != "Id" && name != "ProjectId" && name!= "NRDatas")
             .ToList();

                return Ok(new
                {
                    data = new List<object>(),
                    columns = allColumns,
                    totalRecords = 0,
                    totalPages = 0
                });
            }

            // Find columns where *all* values are null or empty
            var properties = nrDataList.First().GetType().GetProperties();
            // Build dynamic objects that only include non-empty columns
            var result = nrDataList.Select(d =>
            {
                var dict = new Dictionary<string, object>();
                foreach (var prop in properties)
                {
                    dict[prop.Name] = prop.GetValue(d);
                }
                return dict;
            });

            return Ok(new
            {
                items = result,
                columns = properties.Select(p => p.Name).ToList(),
                totalCount,
                totalPages
            });
        }



        // PUT: api/NRDatas/5
        // To protect from overposting attacks, see https://go.microsoft.com/fwlink/?linkid=2123754
        [HttpPut("{id}")]
        public async Task<IActionResult> PutNRData(int id, NRData nRData)
        {
            if (id != nRData.Id)
            {
                return BadRequest();
            }

            _context.Entry(nRData).State = EntityState.Modified;

            try
            {
                await _context.SaveChangesAsync();
                _loggerService.LogEvent($"Updated NrData for {id}", "NRData", User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0,nRData.ProjectId);
            }
            catch (Exception ex)
            {
                if (!NRDataExists(id))
                {
                    _loggerService.LogEvent($"Nrdata with ID {id} not found", "NRData", User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0,nRData.ProjectId);
                    return NotFound();
                }
                else
                {
                    _loggerService.LogError("Error updating NRData", ex.Message, nameof(NRDatasController));
                    return StatusCode(500, "Internal server error");
                }
            }

            return NoContent();
        }

        // POST: api/NRDatas
        // To protect from overposting attacks, see https://go.microsoft.com/fwlink/?linkid=2123754
        [HttpPost]
        public async Task<IActionResult> PostNRData([FromBody] JsonElement inputData)
        {
            if (!ModelState.IsValid)
                return BadRequest(ModelState);

            try
            {
                int projectId = inputData.GetProperty("projectId").GetInt32();
                var dataArray = inputData.GetProperty("data").EnumerateArray();

                var extraConfigs = await _context.ExtraConfigurations
                    .Where(x => x.ProjectId == projectId)
                    .ToListAsync();

                var nrDatasToAdd = new List<NRData>();
                var extraEnvelopesToAdd = new List<ExtraEnvelopes>();

                var nRDataType = typeof(NRData);
                var properties = nRDataType
                    .GetProperties()
                    .ToDictionary(p => p.Name.ToLower(), p => p);

                foreach (var item in dataArray)
                {
                    var nRData = new NRData
                    {
                        ProjectId = projectId
                    };

                    var extraData = new Dictionary<string, string>();

                    foreach (var prop in item.EnumerateObject())
                    {
                        string key = prop.Name.Replace(" ", "").ToLower();
                        string value = prop.Value.ToString();

                        if (properties.TryGetValue(key, out var propInfo))
                        {
                            try
                            {
                                var targetType = Nullable.GetUnderlyingType(propInfo.PropertyType) ?? propInfo.PropertyType;

                                object convertedValue = string.IsNullOrWhiteSpace(value)
                                    ? null
                                    : Convert.ChangeType(value, targetType);

                                propInfo.SetValue(nRData, convertedValue);
                            }
                            catch (Exception e)
                            {
                                throw new Exception($"Error converting '{prop.Name}' value '{value}' to {propInfo.PropertyType.Name}", e);
                            }
                        }
                        else
                        {
                            extraData[prop.Name] = value;
                        }
                    }

                    if (extraData.Any())
                        nRData.NRDatas = JsonSerializer.Serialize(extraData);

                    // =============================
                    // EXTRA CENTER LOGIC
                    // =============================

                    int? extraTypeId = nRData.CenterCode switch
                    {
                        "Nodal Extra" => 1,
                        "University Extra" => 2,
                        "Office Extra" => 3,
                        _ => null
                    };

                    if (extraTypeId.HasValue)
                    {
                        var config = extraConfigs.FirstOrDefault(x => x.ExtraType == extraTypeId);

                        if (config != null)
                        {
                            EnvelopeType envelopeType = null;

                            if (!string.IsNullOrWhiteSpace(config.EnvelopeType))
                            {
                                try
                                {
                                    envelopeType = JsonSerializer.Deserialize<EnvelopeType>(config.EnvelopeType);
                                }
                                catch { }
                            }

                            int? innerCapacity = envelopeType != null ? GetEnvelopeCapacity(envelopeType.Inner) : null;
                            int? outerCapacity = envelopeType != null ? GetEnvelopeCapacity(envelopeType.Outer) : null;

                            int roundedQty = nRData.Quantity;

                            if (innerCapacity > 0)
                                roundedQty = (int)Math.Ceiling((double)nRData.Quantity / innerCapacity.Value) * innerCapacity.Value;
                            else if (outerCapacity > 0)
                                roundedQty = (int)Math.Ceiling((double)nRData.Quantity / outerCapacity.Value) * outerCapacity.Value;

                            string innerEnvelope = innerCapacity > 0
                                ? Math.Ceiling((double)roundedQty / innerCapacity.Value).ToString()
                                : null;

                            string outerEnvelope = outerCapacity > 0
                                ? Math.Ceiling((double)roundedQty / outerCapacity.Value).ToString()
                                : null;

                            extraEnvelopesToAdd.Add(new ExtraEnvelopes
                            {
                                ProjectId = projectId,
                                CatchNo = nRData.CatchNo,
                                ExtraId = extraTypeId.Value,
                                Quantity = roundedQty,
                                InnerEnvelope = innerEnvelope,
                                OuterEnvelope = outerEnvelope
                            });
                        }

                        continue;
                    }

                    nrDatasToAdd.Add(nRData);
                }

                if (nrDatasToAdd.Any())
                    await _context.NRDatas.AddRangeAsync(nrDatasToAdd);

                if (extraEnvelopesToAdd.Any())
                    await _context.ExtrasEnvelope.AddRangeAsync(extraEnvelopesToAdd);

                await _context.SaveChangesAsync();

                _loggerService.LogEvent(
                    $"Created NRData/Extras for Project {projectId}",
                    "NRData",
                    User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0,
                    projectId);

                return Ok(new
                {
                    message = "Data inserted successfully",
                    NRDataCount = nrDatasToAdd.Count,
                    ExtraEnvelopeCount = extraEnvelopesToAdd.Count
                });
            }
            catch (DbUpdateException dbEx)
            {
                var error = dbEx.InnerException?.Message ?? dbEx.Message;

                _loggerService.LogError("Database update error", error, nameof(NRDatasController));

                return StatusCode(500, error);
            }
            catch (Exception ex)
            {
                _loggerService.LogError("NRData processing error", ex.ToString(), nameof(NRDatasController));

                return StatusCode(500, ex.Message);
            }
        }
        public class EnvelopeType
        {
            public string Inner { get; set; }
            public string Outer { get; set; }
        }

        private int GetEnvelopeCapacity(string envelopeCode)
        {
            if (string.IsNullOrWhiteSpace(envelopeCode))
                return 1; // default to 1 if null or invalid

            // Expecting format like "E10", "E25", etc.
            var numberPart = new string(envelopeCode.Where(char.IsDigit).ToArray());

            return int.TryParse(numberPart, out var capacity) ? capacity : 1;
        }


        [HttpGet("ErrorReport")]
        public async Task<ActionResult> GetDuplicateswrtCatch(int ProjectId)
        {
            // 1️⃣ Fetch NR Data
            var nrData = await _context.NRDatas
                .Where(p => p.ProjectId == ProjectId)
                .ToListAsync();

            if (!nrData.Any())
                return NotFound("No NRData found for this project.");

            // 2️⃣ Fetch Project Configuration
            var projectConfig = await _context.ProjectConfigs
                .FirstOrDefaultAsync(p => p.ProjectId == ProjectId);

            if (projectConfig == null)
                return BadRequest("Project configuration not found.");

            // 3️⃣ Extract FieldIds from ProjectConfig JSON
            List<int> fieldIds = new List<int>();

            List<int> ParseIds(string json)
            {
                if (string.IsNullOrWhiteSpace(json))
                    return new List<int>();

                try
                {
                    return JsonSerializer.Deserialize<List<int>>(json);
                }
                catch
                {
                    return new List<int>();
                }
            }

            fieldIds.AddRange(projectConfig.DuplicateCriteria ?? new List<int>());
            fieldIds.AddRange(projectConfig.EnvelopeMakingCriteria ?? new List<int>());
            fieldIds.AddRange(projectConfig.BoxBreakingCriteria ?? new List<int>());
            fieldIds.AddRange(projectConfig.InnerBundlingCriteria ?? new List<int>());

            fieldIds = fieldIds.Distinct().ToList();

            // 4️⃣ Get Field Names from Fields table
            var requiredFields = await _context.Fields
                .Where(f => fieldIds.Contains(f.FieldId))
                .Select(f => f.Name.Trim())
                .ToListAsync();

            // 5️⃣ Unique fields (already configured)
            var uniqueFields = await _context.Fields
                .Where(f => f.IsUnique == true)
                .Select(f => f.Name.Trim())
                .ToListAsync();

            var props = typeof(NRData).GetProperties(BindingFlags.Public | BindingFlags.Instance);

            var catchProp = props.FirstOrDefault(p =>
                string.Equals(p.Name, "CatchNo", StringComparison.OrdinalIgnoreCase));

            var centreProp = props.FirstOrDefault(p =>
                string.Equals(p.Name, "CenterCode", StringComparison.OrdinalIgnoreCase));

            var nodalProp = props.FirstOrDefault(p =>
                string.Equals(p.Name, "NodalCode", StringComparison.OrdinalIgnoreCase));

            if (catchProp == null)
                return BadRequest("CatchNo field not found in NRData.");

            var errorReport = new List<object>();

            // ---------------------------------------------------
            // 6️⃣ CatchNo vs Unique Field Conflict Check
            // ---------------------------------------------------

            foreach (var uniqueField in uniqueFields)
            {
                var uniqueProp = props.FirstOrDefault(p =>
                    string.Equals(p.Name, uniqueField, StringComparison.OrdinalIgnoreCase));

                if (uniqueProp == null)
                    continue;

                var catchGroups = nrData
                    .GroupBy(d => catchProp.GetValue(d)?.ToString()?.Trim() ?? "")
                    .Where(g => !string.IsNullOrEmpty(g.Key));

                foreach (var group in catchGroups)
                {
                    var catchNo = group.Key;

                    var uniqueValues = group
                        .Select(r => uniqueProp.GetValue(r)?.ToString()?.Trim() ?? "")
                        .Distinct()
                        .Where(v => !string.IsNullOrEmpty(v))
                        .ToList();

                    if (uniqueValues.Count > 1)
                    {
                        errorReport.Add(new
                        {
                            CatchNo = catchNo,
                            UniqueField = uniqueField,
                            ConflictingValues = uniqueValues
                        });

                        var conflictRecord = new ConflictingFields
                        {
                            CatchNo = catchNo,
                            ProjectId = ProjectId,
                            UniqueField = uniqueField,
                            ConflictingField = JsonSerializer.Serialize(uniqueValues)
                        };

                        _context.ConflictingFields.Add(conflictRecord);
                    }
                }
            }

            // ---------------------------------------------------
            // 7️⃣ CentreCode mapped to multiple NodalCodes
            // ---------------------------------------------------

            if (centreProp != null && nodalProp != null)
            {
                var centreGroups = nrData
                    .GroupBy(d => centreProp.GetValue(d)?.ToString()?.Trim() ?? "")
                    .Where(g => !string.IsNullOrEmpty(g.Key));

                foreach (var group in centreGroups)
                {
                    var nodalValues = group
                        .Select(r => nodalProp.GetValue(r)?.ToString()?.Trim() ?? "")
                        .Distinct()
                        .Where(v => !string.IsNullOrEmpty(v))
                        .ToList();

                    if (nodalValues.Count > 1)
                    {
                        errorReport.Add(new
                        {
                            CentreCode = group.Key,
                            Error = "CentreCode mapped to multiple NodalCodes",
                            NodalCodes = nodalValues
                        });
                    }
                }
            }

            // ---------------------------------------------------
            // 8️⃣ Required Fields Empty Check
            // ---------------------------------------------------

            foreach (var field in requiredFields)
            {
                var prop = props.FirstOrDefault(p =>
                    string.Equals(p.Name, field, StringComparison.OrdinalIgnoreCase));

                if (prop == null)
                    continue;

                var emptyRows = nrData
                    .Where(r => string.IsNullOrWhiteSpace(prop.GetValue(r)?.ToString()))
                    .Select(r => catchProp.GetValue(r)?.ToString())
                    .ToList();

                if (emptyRows.Any())
                {
                    errorReport.Add(new
                    {
                        Field = field,
                        Error = "Required field is empty",
                        CatchNos = emptyRows
                    });
                }
            }

            await _context.SaveChangesAsync();

            // ---------------------------------------------------
            // 9️⃣ Final Response
            // ---------------------------------------------------

            if (!errorReport.Any())
                return Ok("All validations passed. No errors found.");

            return Ok(new
            {
                ErrorsFound = true,
                Errors = errorReport
            });
        }
        [HttpGet("Counts")]
        public async Task<ActionResult> GetCount(int ProjectId)
        {
            int Conflict = await _context.ConflictingFields.Where(p=>p.ProjectId == ProjectId).CountAsync();
            int NrData = await _context.NRDatas.Where(p => p.ProjectId == ProjectId).CountAsync();
            return Ok(new { Conflict, NrData });
        }

        [HttpPut]
        public async Task<ActionResult> ResolveConflicts(int ProjectId, [FromBody] ConflictResolutionDto payload)
        {
            try
            {
                var NRdata = await _context.NRDatas.Where(a => a.CatchNo == payload.CatchNo && a.ProjectId == ProjectId).ToListAsync();
                if (!NRdata.Any())
                    return NotFound("No matching records found");

                // Use reflection to update the dynamic conflicting field
                foreach (var item in NRdata)
                {
                    var property = item.GetType().GetProperty(payload.UniqueField);
                    if (property != null && property.CanWrite)
                    {
                        property.SetValue(item, payload.SelectedValue);
                    }
                }
                await _context.SaveChangesAsync();
                _loggerService.LogEvent($"Updated NRdata for CatchNo {payload.CatchNo} and ProjectId {ProjectId}", "NRData", User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0,ProjectId);
                return Ok("Conflict resolved successfully");
            }
            catch (Exception ex)
            {
                _loggerService.LogError("Error saving Nrdata", ex.Message, nameof(NRDatasController));
                return StatusCode(500, "Internal Server Error");
            }
        }

        public class ConflictResolutionDto
        {
            public string CatchNo { get; set; }
            public string UniqueField { get; set; }
            public string SelectedValue { get; set; }
        }

        // DELETE: api/NRDatas/5
        [HttpDelete("{id}")]
        public async Task<IActionResult> DeleteNRData(int id)
        {
            try
            {
                var nRData = await _context.NRDatas.FindAsync(id);
                if (nRData == null)
                {
                    return NotFound();
                }

                _context.NRDatas.Remove(nRData);
                await _context.SaveChangesAsync();
                _loggerService.LogEvent($"Deleted NRdata of {id}", "NRData", User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0,nRData.ProjectId);
                return NoContent();
            }
            catch (Exception ex)
            {
                _loggerService.LogError("Error deleting Nrdata", ex.Message, nameof(NRDatasController));
                return StatusCode(500, "Internal Server Error");
            }
        }

        [HttpDelete("DeleteByProject/{ProjectId}")]
        public async Task<IActionResult> DeleteNR(int ProjectId)
        {
            try
            {
                // Fetch all NRData records for the given project
                var nrDataList = await _context.NRDatas
                    .Where(d => d.ProjectId == ProjectId)
                    .ToListAsync();

                if (!nrDataList.Any())
                {
                    return NotFound($"No NRData found for ProjectId {ProjectId}");
                }
                var reportPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", ProjectId.ToString());
                if (Directory.Exists(reportPath))
                {
                    Directory.Delete(reportPath, true); // 'true' allows recursive deletion of files and subdirectories
                }
                // Remove all records
                _context.NRDatas.RemoveRange(nrDataList);
                await _context.SaveChangesAsync();

                _loggerService.LogEvent(
                    $"Deleted all NRData for ProjectId {ProjectId}",
                    "NRData",
                    User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0,ProjectId
                );

                return NoContent();
            }
            catch (Exception ex)
            {
                _loggerService.LogError("Error deleting NRData", ex.Message, nameof(NRDatasController));
                return StatusCode(500, "Internal Server Error");
            }
        }

        private bool NRDataExists(int id)
        {
            return _context.NRDatas.Any(e => e.Id == id);
        }
    }
}
