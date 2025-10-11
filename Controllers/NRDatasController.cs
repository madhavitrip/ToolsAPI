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
        public async Task<ActionResult<IEnumerable<NRData>>> GetByProjectId(int projectId)
        {
            var nrData = await _context.NRDatas
                .Where(d => d.ProjectId == projectId)
                .Select(d=>new
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
                    d.Pages,
                })
                .ToListAsync();

            if (!nrData.Any())
            {
                return NotFound($"No NRData found for ProjectId {projectId}");
            }



            return Ok(nrData);
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
                _loggerService.LogEvent($"Updated NrData for {id}", "NRData", User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0);
            }
            catch (Exception ex)
            {
                if (!NRDataExists(id))
                {
                    _loggerService.LogEvent($"Nrdata with ID {id} not found", "NRData", User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0);
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
        public async Task<ActionResult> PostNRData([FromBody] JsonElement inputData)
        {
            try
            {
                int projectId = inputData.GetProperty("projectId").GetInt32();
                var dataArray = inputData.GetProperty("data").EnumerateArray();

                foreach (var item in dataArray)
                {
                    var nRData = new NRData();
                    nRData.ProjectId = projectId;

                    var extraData = new Dictionary<string, string>();
                    var nRDataType = typeof(NRData);

                    foreach (var prop in item.EnumerateObject())
                    {
                        string key = prop.Name;
                        string value = prop.Value.ToString();

                        var propInfo = nRDataType.GetProperty(key.Replace(" ", ""),
                                            System.Reflection.BindingFlags.IgnoreCase |
                                            System.Reflection.BindingFlags.Public |
                                            System.Reflection.BindingFlags.Instance);

                        if (propInfo != null)
                        {
                            var targetType = Nullable.GetUnderlyingType(propInfo.PropertyType) ?? propInfo.PropertyType;
                            object convertedValue = string.IsNullOrEmpty(value) ? null : Convert.ChangeType(value, targetType);
                            propInfo.SetValue(nRData, convertedValue);
                        }
                        else
                        {
                            extraData[key] = value;
                        }
                    }

                    if (extraData.Count > 0)
                        nRData.NRDatas = System.Text.Json.JsonSerializer.Serialize(extraData);

                    _context.NRDatas.Add(nRData);
                }

                await _context.SaveChangesAsync();
                _loggerService.LogEvent($"Created new NRadat with ID {projectId}", "NRData", User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0);
                return Ok("Data inserted successfully");
            }
            catch (Exception ex)
            {
                _loggerService.LogError("Error saving Nrdata", ex.Message, nameof(NRDatasController));
                return StatusCode(500, "Internal Server Error");
            }
        }


        [HttpGet("ErrorReport")]
        public async Task<ActionResult> GetDuplicateswrtCatch(int ProjectId)
        {
            var nrData = await _context.NRDatas
                .Where(p => p.ProjectId == ProjectId)
                .ToListAsync();

            if (!nrData.Any())
                return NotFound("No NRData found for this project.");

            var uniqueFields = await _context.Fields
                .Where(f => f.IsUnique == true)
                .Select(f => f.Name.Trim())
                .ToListAsync();

            if (!uniqueFields.Any())
                return BadRequest("No unique fields configured.");

            var props = typeof(NRData).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            var catchProp = props.FirstOrDefault(p => string.Equals(p.Name, "CatchNo", StringComparison.OrdinalIgnoreCase));

            if (catchProp == null)
                return BadRequest("CatchNo field not found in NRData.");

            var errorReport = new List<object>();

            foreach (var uniqueField in uniqueFields)
            {
                var uniqueProp = props.FirstOrDefault(p => string.Equals(p.Name, uniqueField, StringComparison.OrdinalIgnoreCase));
                if (uniqueProp == null)
                    continue;

                // Step 1: Group by CatchNo
                var catchGroups = nrData
                    .GroupBy(d => catchProp.GetValue(d)?.ToString()?.Trim() ?? "")
                    .Where(g => !string.IsNullOrEmpty(g.Key));

                foreach (var group in catchGroups)
                {
                    var catchNo = group.Key;

                    // Get all distinct values for the current unique field within this CatchNo group
                    var uniqueValues = group
                        .Select(r => uniqueProp.GetValue(r)?.ToString()?.Trim() ?? "")
                        .Distinct()
                        .Where(val => !string.IsNullOrEmpty(val))
                        .ToList();
                    var json = JsonSerializer.Serialize(uniqueValues);
                  

                    if (uniqueValues.Count > 1)
                    {
                        // Error: CatchNo maps to multiple values for the unique field
                        errorReport.Add(new
                        {
                            CatchNo = catchNo,
                            UniqueField = uniqueField,
                            ConflictingValues = uniqueValues
                        });
                    }
                    var conflictRecord = new ConflictingFields
                    {
                        CatchNo = catchNo,
                        ProjectId = ProjectId, // set this from your current context
                        UniqueField = uniqueField,
                        ConflictingField = json
                    };

                    _context.ConflictingFields.Add(conflictRecord);
                }
            }

            if (!errorReport.Any())
                return Ok("All CatchNo mappings to unique fields are valid.");

            return Ok(new
            {
                DuplicatesFound = true,
                Errors = errorReport
            });
        }

        [HttpGet]
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
                _loggerService.LogEvent($"Updated NRdata for CatchNo {payload.CatchNo} and ProjectId {ProjectId}", "NRData", User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0);
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
                _loggerService.LogEvent($"Deleted NRdata of {id}", "NRData", User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0);
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

                // Remove all records
                _context.NRDatas.RemoveRange(nrDataList);
                await _context.SaveChangesAsync();

                _loggerService.LogEvent(
                    $"Deleted all NRData for ProjectId {ProjectId}",
                    "NRData",
                    User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0
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
