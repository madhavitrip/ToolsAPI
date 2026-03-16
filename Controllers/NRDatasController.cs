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
        private const string CatchUniqueFieldConflict = "catch_unique_field";
        private const string CenterMultipleNodalsConflict = "center_multiple_nodals";
        private const string CollegeMultipleNodalsConflict = "college_multiple_nodals";
        private const string CollegeMultipleCentersConflict = "college_multiple_centers";
        private const string RequiredFieldEmptyConflict = "required_field_empty";
        private const string ZeroNrQuantityConflict = "zero_nr_quantity";
        private const string NodalCodeDigitMismatchConflict = "nodal_code_digit_mismatch";
        private const int ConflictStatusPending = 0;
        private const int ConflictStatusResolved = 1;
        private const int ConflictStatusIgnored = 2;

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
        public async Task<ActionResult> GetByProjectId(
            int projectId,
            int pageSize,
            int pageNo,
            string? search = null,
            string? key = null,
            string? sortField = null,
            string? sortOrder = null)
        {
            IQueryable<NRData> query = _context.NRDatas
             .Where(d => d.ProjectId == projectId && d.Status == true);

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

            // Server-side sorting so it works across all pages.
            var normalizedSortField = NormalizeText(sortField);
            var normalizedSortOrder = NormalizeText(sortOrder).ToLowerInvariant();
            var isAscending = normalizedSortOrder != "descend";

            if (!string.IsNullOrWhiteSpace(normalizedSortField))
            {
                query = normalizedSortField switch
                {
                    "CatchNo" => isAscending ? query.OrderBy(d => d.CatchNo) : query.OrderByDescending(d => d.CatchNo),
                    "CenterCode" => isAscending ? query.OrderBy(d => d.CenterCode) : query.OrderByDescending(d => d.CenterCode),
                    "ExamDate" => isAscending ? query.OrderBy(d => d.ExamDate) : query.OrderByDescending(d => d.ExamDate),
                    "ExamTime" => isAscending ? query.OrderBy(d => d.ExamTime) : query.OrderByDescending(d => d.ExamTime),
                    "NRQuantity" => isAscending ? query.OrderBy(d => d.NRQuantity) : query.OrderByDescending(d => d.NRQuantity),
                    "Quantity" => isAscending ? query.OrderBy(d => d.Quantity) : query.OrderByDescending(d => d.Quantity),
                    "NodalCode" => isAscending ? query.OrderBy(d => d.NodalCode) : query.OrderByDescending(d => d.NodalCode),
                    "CourseName" => isAscending ? query.OrderBy(d => d.CourseName) : query.OrderByDescending(d => d.CourseName),
                    "SubjectName" => isAscending ? query.OrderBy(d => d.SubjectName) : query.OrderByDescending(d => d.SubjectName),
                    _ => query.OrderBy(d => d.Id),
                };
            }
            else
            {
                query = query.OrderBy(d => d.Id);
            }

            int totalCount = await query.CountAsync();
            int totalPages = (int)Math.Ceiling(totalCount / (double)pageSize);

            var nrDataList = await query
                .Skip((pageNo - 1) * pageSize)
                .Take(pageSize)
                .Select(d => new
                {
                    d.Id,
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

        [HttpPost("merge-catchnos")]
        public async Task<ActionResult> MergeCatchNos(int ProjectId, [FromBody] MergeCatchNoRequest request)
        {
            if (request == null)
            {
                return BadRequest("Please select rows from 2 catch numbers to merge.");
            }

            var separator = NormalizeText(request.Separator);
            if (separator != "/" && separator != "-")
            {
                return BadRequest("Separator must be '/' or '-'.");
            }

            var requestedCatchNos = (request.CatchNos ?? new List<string>())
                .Select(NormalizeText)
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (requestedCatchNos.Count != 2)
            {
                return BadRequest("Please select exactly 2 catch numbers to merge.");
            }

            // Expand across the whole dataset (not just the current page selection).
            var normalizedRequestedCatchNos = requestedCatchNos
                .Select(value => value.Trim().ToLowerInvariant())
                .Distinct()
                .ToList();

            var selectedRows = await _context.NRDatas
                .Where(item =>
                    item.ProjectId == ProjectId &&
                    item.Status == true &&
                    item.CatchNo != null &&
                    normalizedRequestedCatchNos.Contains(item.CatchNo.Trim().ToLower()))
                .ToListAsync();

            if (selectedRows.Count == 0)
            {
                return NotFound("Selected catch numbers were not found for this project.");
            }

            var selectedCatchNos = selectedRows
                .Select(item => NormalizeText(item.CatchNo))
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (selectedCatchNos.Count != 2)
            {
                return BadRequest("Please select rows from exactly 2 different catch numbers.");
            }

            var firstCatchNo = requestedCatchNos[0];
            var secondCatchNo = requestedCatchNos[1];

            var normalizedSchedules = selectedRows
                .Select(item => new
                {
                    ExamDate = NormalizeText(item.ExamDate),
                    ExamTime = NormalizeText(item.ExamTime)
                })
                .Distinct()
                .ToList();

            if (normalizedSchedules.Count != 1)
            {
                return BadRequest("Catch numbers can only be merged when ExamDate and ExamTime are the same.");
            }

            var mergedCatchNo = $"{firstCatchNo}{separator}{secondCatchNo}";
            // Merge per-center:
            // - If both catch numbers exist for the same center, add quantities and collapse into one row for that center.
            // - If a center appears for only one catch number, keep rows (no add), but rewrite CatchNo to the merged form.
            var selectedRowsByCenter = selectedRows
                .GroupBy(item => NormalizeText(item.CenterCode), StringComparer.OrdinalIgnoreCase)
                .ToList();

            int centersCollapsed = 0;
            int rowsDeleted = 0;

            foreach (var centerGroup in selectedRowsByCenter)
            {
                var groupRows = centerGroup.ToList();
                if (groupRows.Count == 0)
                {
                    continue;
                }

                var groupCatchNos = groupRows
                    .Select(item => NormalizeText(item.CatchNo))
                    .Where(value => !string.IsNullOrWhiteSpace(value))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();

                var hasFirst = groupCatchNos.Any(value => string.Equals(value, firstCatchNo, StringComparison.OrdinalIgnoreCase));
                var hasSecond = groupCatchNos.Any(value => string.Equals(value, secondCatchNo, StringComparison.OrdinalIgnoreCase));

                if (hasFirst && hasSecond)
                {
                    centersCollapsed++;

                    var primaryRow = groupRows
                        .OrderBy(item => item.Id)
                        .First();

                    primaryRow.CatchNo = mergedCatchNo;
                    primaryRow.Quantity = groupRows.Sum(item => item.Quantity);
                    primaryRow.NRQuantity = groupRows.Sum(item => item.NRQuantity);

                    foreach (var row in groupRows.Where(item => item.Id != primaryRow.Id))
                    {
                        row.Status = false;              // soft delete
                        row.NRDataId = primaryRow.Id; // reference merged row
                        rowsDeleted++;
                    }

                    continue;
                }

                foreach (var row in groupRows)
                {
                    row.CatchNo = mergedCatchNo;
                }
            }

            await _context.SaveChangesAsync();

            var actionLabel = centersCollapsed > 0
                ? $"Catch numbers merged. Collapsed {centersCollapsed} center(s) (deleted {rowsDeleted} row(s)) and added quantities where centers matched."
                : "Catch numbers merged. No centers had both catch numbers, so quantities were kept unchanged.";

            _loggerService.LogEvent(
                $"Merged catch numbers for ProjectId {ProjectId}: {firstCatchNo} and {secondCatchNo}",
                "NRData",
                User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0,
                ProjectId);

            return Ok(new
            {
                message = actionLabel,
                catchNo = mergedCatchNo,
                centersCollapsed,
                rowsDeleted
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
            var nrData = await _context.NRDatas
                .Where(item => item.ProjectId == ProjectId)
                .ToListAsync();
            if (nrData.Count == 0)
            {
                return Ok(new
                {
                    ErrorsFound = false,
                    Errors = new List<ConflictReportItem>()
                });
            }

            var rowsWithMeta = BuildConflictRows(nrData);
            var projectConfig = await _context.ProjectConfigs
                .FirstOrDefaultAsync(item => item.ProjectId == ProjectId);

            var fieldIds = new List<int>();
            if (projectConfig != null)
            {
                fieldIds.AddRange(projectConfig.DuplicateCriteria ?? new List<int>());
                fieldIds.AddRange(projectConfig.EnvelopeMakingCriteria ?? new List<int>());
                fieldIds.AddRange(projectConfig.BoxBreakingCriteria ?? new List<int>());
                fieldIds.AddRange(projectConfig.InnerBundlingCriteria ?? new List<int>());
            }

            var requiredFields = await _context.Fields
                .Where(field => fieldIds.Distinct().Contains(field.FieldId))
                .Select(field => field.Name.Trim())
                .ToListAsync();

            if (!requiredFields.Contains("NRQuantity", StringComparer.OrdinalIgnoreCase))
            {
                requiredFields.Add("NRQuantity");
            }

            var uniqueFields = await _context.Fields
                .Where(field => field.IsUnique == true)
                .Select(field => field.Name.Trim())
                .ToListAsync();

            var errorReport = BuildConflictReport(rowsWithMeta, requiredFields, uniqueFields);

            await SyncConflictStatuses(ProjectId, errorReport);

            return Ok(new
            {
                ErrorsFound = errorReport.Count > 0,
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
                if (payload == null || string.IsNullOrWhiteSpace(payload.ConflictType) || string.IsNullOrWhiteSpace(payload.SelectedValue))
                {
                    return BadRequest("Conflict type and selected value are required.");
                }

                var rowsToUpdate = await GetRowsForConflict(ProjectId, payload);
                if (!rowsToUpdate.Any())
                {
                    return NotFound("No matching records found");
                }

                foreach (var item in rowsToUpdate)
                {
                    ApplyConflictResolution(item, payload);
                }

                await _context.SaveChangesAsync();
                await UpsertConflictStatus(ProjectId, payload, ConflictStatusResolved);

                _loggerService.LogEvent($"Updated NRdata for CatchNo {payload.CatchNo} and ProjectId {ProjectId}", "NRData", User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0,ProjectId);
                return Ok("Conflict resolved successfully");
            }
            catch (Exception ex)
            {
                _loggerService.LogError("Error saving Nrdata", ex.Message, nameof(NRDatasController));
                return StatusCode(500, "Internal Server Error");
            }
        }

        [HttpPut("conflicts/status")]
        public async Task<ActionResult> UpdateConflictStatus(int ProjectId, [FromBody] ConflictStatusUpdateDto payload)
        {
            if (payload == null || string.IsNullOrWhiteSpace(payload.ConflictType))
            {
                return BadRequest("Conflict details are required.");
            }

            var normalizedStatus = NormalizeStatusCode(payload.Status);
            if (normalizedStatus == ConflictStatusIgnored && !CanIgnoreConflict(payload.ConflictType))
            {
                return BadRequest("Ignore is not allowed for this conflict type.");
            }

            await UpsertConflictStatus(ProjectId, payload, normalizedStatus);

            return Ok(new
            {
                message = $"Conflict marked as {NormalizeStatusLabel(normalizedStatus)}.",
                status = NormalizeStatusLabel(normalizedStatus)
            });
        }

        private static List<NRDataConflictProjection> BuildConflictRows(IEnumerable<NRData> rows)
        {
            return rows.Select(row =>
            {
                var jsonValues = ParseNrDatas(row.NRDatas);
                var collegeData = GetCollegeData(row);
                return new NRDataConflictProjection
                {
                    Row = row,
                    CatchNo = NormalizeText(row.CatchNo),
                    ExamDate = NormalizeText(row.ExamDate),
                    ExamTime = NormalizeText(row.ExamTime),
                    CenterCode = NormalizeText(row.CenterCode),
                    NodalCode = NormalizeText(row.NodalCode),
                    NrQuantity = row.NRQuantity,
                    ImportRowNo = ParseImportRowNo(jsonValues),
                    CollegeName = collegeData.CollegeName,
                    CollegeCode = collegeData.CollegeCode,
                };
            }).ToList();
        }

        private List<ConflictReportItem> BuildConflictReport(
            List<NRDataConflictProjection> rowsWithMeta,
            List<string> requiredFields,
            List<string> uniqueFields)
        {
            var conflicts = new List<ConflictReportItem>();

            AddCatchUniqueFieldConflicts(conflicts, rowsWithMeta, uniqueFields);
            AddCenterMultipleNodalConflicts(conflicts, rowsWithMeta);
            AddCollegeMappingConflicts(conflicts, rowsWithMeta);
            AddRequiredFieldConflicts(conflicts, rowsWithMeta, requiredFields);
            AddZeroQuantityConflicts(conflicts, rowsWithMeta);
            AddNodalCodeDigitMismatchConflicts(conflicts, rowsWithMeta);

            return conflicts;
        }

        private void AddCatchUniqueFieldConflicts(
            List<ConflictReportItem> conflicts,
            List<NRDataConflictProjection> rowsWithMeta,
            IEnumerable<string> uniqueFields)
        {
            var catchGroups = rowsWithMeta
                .Where(item => !string.IsNullOrWhiteSpace(item.CatchNo))
                .GroupBy(item => item.CatchNo)
                .ToList();

            foreach (var uniqueField in uniqueFields.Distinct(StringComparer.OrdinalIgnoreCase))
            {
                var uniqueProperty = GetNRDataProperty(uniqueField);
                if (uniqueProperty == null)
                {
                    continue;
                }

                foreach (var group in catchGroups)
                {
                    var uniqueValues = DistinctNonEmpty(group.Select(item => uniqueProperty.GetValue(item.Row)?.ToString()));
                    if (uniqueValues.Count <= 1)
                    {
                        continue;
                    }

                    conflicts.Add(new ConflictReportItem
                    {
                        ConflictType = CatchUniqueFieldConflict,
                        CatchNo = group.Key,
                        CatchNos = new List<string> { group.Key },
                        RowIds = group.Select(item => item.Row.Id).Distinct().OrderBy(id => id).ToList(),
                        ImportRowNos = group.Where(item => item.ImportRowNo.HasValue).Select(item => item.ImportRowNo!.Value).Distinct().OrderBy(id => id).ToList(),
                        UniqueField = uniqueField,
                        ConflictingValues = uniqueValues,
                        CanIgnore = true,
                        CanResolve = true,
                        Summary = $"Catch No {group.Key} has multiple {uniqueField} values."
                    });
                }
            }
        }

        private static void AddCenterMultipleNodalConflicts(
            List<ConflictReportItem> conflicts,
            List<NRDataConflictProjection> rowsWithMeta)
        {
            var centerGroups = rowsWithMeta
                .Where(item => !string.IsNullOrWhiteSpace(item.CenterCode))
                .GroupBy(item => item.CenterCode);

            foreach (var group in centerGroups)
            {
                var nodalValues = DistinctNonEmpty(group.Select(item => item.NodalCode));
                if (nodalValues.Count <= 1)
                {
                    continue;
                }

                conflicts.Add(new ConflictReportItem
                {
                    ConflictType = CenterMultipleNodalsConflict,
                    CentreCode = group.Key,
                    CatchNos = DistinctNonEmpty(group.Select(item => item.CatchNo)),
                    RowIds = group.Select(item => item.Row.Id).Distinct().OrderBy(id => id).ToList(),
                    ImportRowNos = group.Where(item => item.ImportRowNo.HasValue).Select(item => item.ImportRowNo!.Value).Distinct().OrderBy(id => id).ToList(),
                    UniqueField = "NodalCode",
                    NodalCodes = nodalValues,
                    ConflictingValues = nodalValues,
                    CanIgnore = false,
                    CanResolve = true,
                    Summary = $"Centre {group.Key} is linked with multiple nodal codes."
                });
            }
        }

        private static void AddCollegeMappingConflicts(
            List<ConflictReportItem> conflicts,
            List<NRDataConflictProjection> rowsWithMeta)
        {
            BuildCollegeConflicts(conflicts, rowsWithMeta, "CollegeName", CollegeMultipleNodalsConflict, "NodalCode", useNodalValues: true);
            BuildCollegeConflicts(conflicts, rowsWithMeta, "CollegeCode", CollegeMultipleNodalsConflict, "NodalCode", useNodalValues: true);
            BuildCollegeConflicts(conflicts, rowsWithMeta, "CollegeName", CollegeMultipleCentersConflict, "CenterCode", useNodalValues: false);
            BuildCollegeConflicts(conflicts, rowsWithMeta, "CollegeCode", CollegeMultipleCentersConflict, "CenterCode", useNodalValues: false);
        }

        private static void AddRequiredFieldConflicts(
            List<ConflictReportItem> conflicts,
            List<NRDataConflictProjection> rowsWithMeta,
            IEnumerable<string> requiredFields)
        {
            foreach (var field in requiredFields.Distinct(StringComparer.OrdinalIgnoreCase))
            {
                var property = GetNRDataProperty(field);
                if (property == null)
                {
                    continue;
                }

                var emptyRows = rowsWithMeta
                    .Where(item => string.IsNullOrWhiteSpace(property.GetValue(item.Row)?.ToString()));

                foreach (var emptyRow in emptyRows)
                {
                    conflicts.Add(new ConflictReportItem
                    {
                        ConflictType = RequiredFieldEmptyConflict,
                        Field = field,
                        CatchNo = emptyRow.CatchNo,
                        CatchNos = string.IsNullOrWhiteSpace(emptyRow.CatchNo) ? new List<string>() : new List<string> { emptyRow.CatchNo },
                        RowIds = new List<int> { emptyRow.Row.Id },
                        ImportRowNos = emptyRow.ImportRowNo.HasValue ? new List<int> { emptyRow.ImportRowNo.Value } : new List<int>(),
                        CanIgnore = false,
                        CanResolve = true,
                        Summary = BuildRequiredFieldSummary(field, emptyRow),
                    });
                }
            }
        }

        private static void AddZeroQuantityConflicts(
            List<ConflictReportItem> conflicts,
            List<NRDataConflictProjection> rowsWithMeta)
        {
            var positiveQuantities = rowsWithMeta
                .Where(item => item.NrQuantity > 0)
                .Select(item => item.NrQuantity)
                .ToList();

            var minNrQuantity = positiveQuantities.Count > 0 ? positiveQuantities.Min() : (int?)null;
            var maxNrQuantity = positiveQuantities.Count > 0 ? positiveQuantities.Max() : (int?)null;

            var zeroQuantityCatchNos = DistinctNonEmpty(
                rowsWithMeta
                    .Where(item => item.NrQuantity == 0)
                    .Select(item => item.CatchNo));

            if (zeroQuantityCatchNos.Count == 0)
            {
                return;
            }

            conflicts.Add(new ConflictReportItem
            {
                ConflictType = ZeroNrQuantityConflict,
                Field = "NRQuantity",
                CatchNos = zeroQuantityCatchNos,
                RowIds = rowsWithMeta
                    .Where(item => item.NrQuantity == 0)
                    .Select(item => item.Row.Id)
                    .Distinct()
                    .OrderBy(id => id)
                    .ToList(),
                ImportRowNos = rowsWithMeta
                    .Where(item => item.NrQuantity == 0 && item.ImportRowNo.HasValue)
                    .Select(item => item.ImportRowNo!.Value)
                    .Distinct()
                    .OrderBy(id => id)
                    .ToList(),
                MinNrQuantity = minNrQuantity,
                MaxNrQuantity = maxNrQuantity,
                CanIgnore = true,
                CanResolve = true,
                Summary = $"NRQuantity is 0 for {zeroQuantityCatchNos.Count} catch number(s)."
            });
        }

        private static void AddNodalCodeDigitMismatchConflicts(
            List<ConflictReportItem> conflicts,
            List<NRDataConflictProjection> rowsWithMeta)
        {
            var nodalRows = rowsWithMeta
                .Where(item => !string.IsNullOrWhiteSpace(item.NodalCode))
                .ToList();
            var rowsWithDigitCount = nodalRows
                .Select(item => new
                {
                    Row = item,
                    DigitCount = GetDigitCount(item.NodalCode),
                })
                .Where(item => item.DigitCount > 0)
                .ToList();

            var distinctLengths = rowsWithDigitCount
                .Select(item => item.DigitCount)
                .Distinct()
                .OrderBy(length => length)
                .ToList();

            if (distinctLengths.Count <= 1)
            {
                return;
            }

            var expectedDigitCount = rowsWithDigitCount
                .GroupBy(item => item.DigitCount)
                .OrderByDescending(group => group.Count())
                .ThenBy(group => group.Key)
                .Select(group => group.Key)
                .First();

            foreach (var item in rowsWithDigitCount.Where(item => item.DigitCount != expectedDigitCount))
            {
                conflicts.Add(new ConflictReportItem
                {
                    ConflictType = NodalCodeDigitMismatchConflict,
                    CatchNo = item.Row.CatchNo,
                    CatchNos = string.IsNullOrWhiteSpace(item.Row.CatchNo) ? new List<string>() : new List<string> { item.Row.CatchNo },
                    RowIds = new List<int> { item.Row.Row.Id },
                    ImportRowNos = item.Row.ImportRowNo.HasValue ? new List<int> { item.Row.ImportRowNo.Value } : new List<int>(),
                    NodalCode = item.Row.NodalCode,
                    NodalCodeGroup = expectedDigitCount.ToString(),
                    UniqueField = "NodalCode",
                    ConflictingValues = new List<string> { item.Row.NodalCode },
                    CanIgnore = false,
                    CanResolve = true,
                    Summary = $"Nodal code {item.Row.NodalCode} has {item.DigitCount} digits; expected {expectedDigitCount} digits."
                });
            }
        }

        private async Task<List<NRData>> GetRowsForConflict(int projectId, ConflictActionDto payload)
        {
            var projectRows = await _context.NRDatas
                .Where(item => item.ProjectId == projectId)
                .ToListAsync();

            var catchNoSet = new HashSet<string>(
                (payload.CatchNos ?? new List<string>())
                    .Where(value => !string.IsNullOrWhiteSpace(value))
                    .Select(value => value.Trim()),
                StringComparer.OrdinalIgnoreCase);

            if (!string.IsNullOrWhiteSpace(payload.CatchNo))
            {
                catchNoSet.Add(payload.CatchNo.Trim());
            }

            return (payload.ConflictType ?? string.Empty).Trim().ToLowerInvariant() switch
            {
                CatchUniqueFieldConflict => projectRows
                    .Where(item => string.Equals(item.CatchNo, payload.CatchNo, StringComparison.OrdinalIgnoreCase))
                    .ToList(),
                CenterMultipleNodalsConflict => projectRows
                    .Where(item => string.Equals(item.CenterCode, payload.CentreCode, StringComparison.OrdinalIgnoreCase))
                    .ToList(),
                CollegeMultipleNodalsConflict => projectRows
                    .Where(item => MatchesCollege(item, payload))
                    .ToList(),
                CollegeMultipleCentersConflict => projectRows
                    .Where(item => MatchesCollege(item, payload))
                    .ToList(),
                RequiredFieldEmptyConflict => projectRows
                    .Where(item => (payload.RowIds ?? new List<int>()).Contains(item.Id))
                    .ToList(),
                ZeroNrQuantityConflict => projectRows
                    .Where(item => item.CatchNo != null && catchNoSet.Contains(item.CatchNo) && item.NRQuantity == 0)
                    .ToList(),
                NodalCodeDigitMismatchConflict => projectRows
                    .Where(item => (payload.RowIds ?? new List<int>()).Contains(item.Id))
                    .ToList(),
                _ => new List<NRData>(),
            };
        }

        private void ApplyConflictResolution(NRData item, ConflictResolutionDto payload)
        {
            var conflictType = (payload.ConflictType ?? string.Empty).Trim().ToLowerInvariant();
            switch (conflictType)
            {
                case CatchUniqueFieldConflict:
                    SetPropertyValue(item, payload.UniqueField, payload.SelectedValue);
                    break;
                case CenterMultipleNodalsConflict:
                case CollegeMultipleNodalsConflict:
                case NodalCodeDigitMismatchConflict:
                    item.NodalCode = payload.SelectedValue.Trim();
                    break;
                case CollegeMultipleCentersConflict:
                    item.CenterCode = payload.SelectedValue.Trim();
                    break;
                case RequiredFieldEmptyConflict:
                    SetPropertyValue(item, payload.Field, payload.SelectedValue);
                    break;
                case ZeroNrQuantityConflict:
                    SetPropertyValue(item, payload.Field ?? "NRQuantity", payload.SelectedValue);
                    break;
                default:
                    throw new InvalidOperationException($"Unsupported conflict type '{payload.ConflictType}'.");
            }
        }

        private static void SetPropertyValue(NRData item, string? propertyName, string? selectedValue)
        {
            if (string.IsNullOrWhiteSpace(propertyName))
            {
                throw new InvalidOperationException("Target field is required.");
            }

            var property = item.GetType().GetProperty(propertyName, BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);
            if (property == null || !property.CanWrite)
            {
                throw new InvalidOperationException($"Field '{propertyName}' cannot be updated.");
            }

            var targetType = Nullable.GetUnderlyingType(property.PropertyType) ?? property.PropertyType;
            object? convertedValue = string.IsNullOrWhiteSpace(selectedValue)
                ? null
                : Convert.ChangeType(selectedValue.Trim(), targetType);

            property.SetValue(item, convertedValue);
        }

        private async Task SyncConflictStatuses(int projectId, List<ConflictReportItem> conflicts)
        {
            var existingRecords = await _context.ConflictingFields
                .Where(item => item.ProjectId == projectId)
                .ToListAsync();

            var existingLookup = existingRecords
                .GroupBy(GetConflictStorageKey)
                .ToDictionary(group => group.Key, group => group.OrderByDescending(item => item.Id).First());

            var newRecords = new List<ConflictingFields>();

            foreach (var conflict in conflicts)
            {
                var storageKey = BuildConflictStorageKey(projectId, conflict);
                if (existingLookup.TryGetValue(storageKey, out var existingRecord))
                {
                    conflict.Status = NormalizeStatusLabel(existingRecord.Status);
                    continue;
                }

                conflict.Status = "pending";
                newRecords.Add(BuildConflictEntity(projectId, conflict, ConflictStatusPending));
            }

            if (newRecords.Count > 0)
            {
                _context.ConflictingFields.AddRange(newRecords);
                await _context.SaveChangesAsync();
            }
        }

        private async Task UpsertConflictStatus(int projectId, ConflictActionDto payload, int status)
        {
            var entity = BuildConflictEntity(projectId, payload, status);
            var storageKey = GetConflictStorageKey(entity);

            var existingRecord = await _context.ConflictingFields
                .Where(item => item.ProjectId == projectId)
                .ToListAsync();

            var matchedRecord = existingRecord
                .OrderByDescending(item => item.Id)
                .FirstOrDefault(item => GetConflictStorageKey(item) == storageKey);

            if (matchedRecord == null)
            {
                _context.ConflictingFields.Add(entity);
            }
            else
            {
                matchedRecord.Status = status;
            }

            await _context.SaveChangesAsync();
        }

        private static ConflictingFields BuildConflictEntity(int projectId, ConflictActionDto payload, int status)
        {
            var storagePayload = BuildConflictStoragePayload(payload);

            return new ConflictingFields
            {
                NRDataId = ResolveStorageNrDataId(payload),
                ProjectId = projectId,
                UniqueField = ResolveStorageUniqueField(payload),
                ConflictingField = JsonSerializer.Serialize(storagePayload),
                Status = status,
            };
        }

        private static string BuildConflictStorageKey(int projectId, ConflictActionDto payload)
        {
            var entity = BuildConflictEntity(projectId, payload, ConflictStatusPending);
            return GetConflictStorageKey(entity);
        }

        private static string GetConflictStorageKey(ConflictingFields entity)
        {
            return $"{entity.ProjectId}|{entity.NRDataId}|{entity.UniqueField}|{entity.ConflictingField}";
        }

        private static int ResolveStorageNrDataId(ConflictActionDto payload)
        {
            return (payload.RowIds ?? new List<int>())
                .Where(id => id > 0)
                .Distinct()
                .OrderBy(id => id)
                .FirstOrDefault();
        }

        private static string ResolveStorageUniqueField(ConflictActionDto payload)
        {
            var uniqueField = NormalizeText(payload.UniqueField);
            if (!string.IsNullOrWhiteSpace(uniqueField))
            {
                return uniqueField;
            }

            var field = NormalizeText(payload.Field);
            if (!string.IsNullOrWhiteSpace(field))
            {
                return field;
            }

            return NormalizeText(payload.ConflictType);
        }

        private static bool CanIgnoreConflict(string? conflictType)
        {
            var normalized = NormalizeText(conflictType);
            return string.Equals(normalized, CatchUniqueFieldConflict, StringComparison.OrdinalIgnoreCase)
                || string.Equals(normalized, ZeroNrQuantityConflict, StringComparison.OrdinalIgnoreCase);
        }

        private static PropertyInfo? GetNRDataProperty(string? propertyName)
        {
            return typeof(NRData).GetProperty(
                propertyName ?? string.Empty,
                BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);
        }

        private static string BuildRequiredFieldSummary(string field, NRDataConflictProjection row)
        {
            var catchLabel = string.IsNullOrWhiteSpace(row.CatchNo)
                ? string.Empty
                : $" (Catch No {row.CatchNo})";

            var rowLabel = row.ImportRowNo?.ToString() ?? row.Row.Id.ToString();
            return $"{field} is missing for row {rowLabel}{catchLabel}.";
        }

        private static bool MatchesCollege(NRData item, ConflictActionDto payload)
        {
            var collegeData = GetCollegeData(item);

            if (string.Equals(payload.CollegeKeyType, "CollegeCode", StringComparison.OrdinalIgnoreCase))
            {
                return string.Equals(collegeData.CollegeCode, payload.CollegeCode?.Trim(), StringComparison.OrdinalIgnoreCase);
            }

            return string.Equals(collegeData.CollegeName, payload.CollegeName?.Trim(), StringComparison.OrdinalIgnoreCase);
        }

        private static (string CollegeName, string CollegeCode) GetCollegeData(NRData row)
        {
            var jsonValues = ParseNrDatas(row.NRDatas);
            return (
                ReadValue(jsonValues, row, "CollegeName", "College Name", "college name", "colleage name", "colleageName", "College"),
                ReadValue(jsonValues, row, "CollegeCode", "College Code", "college code", "colleage code", "colleageCode", "collegecode")
            );
        }

        private static Dictionary<string, string> ParseNrDatas(string? nrDatas)
        {
            if (string.IsNullOrWhiteSpace(nrDatas))
            {
                return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            }

            try
            {
                var parsed = JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(nrDatas);
                return parsed?.ToDictionary(
                    item => item.Key,
                    item => item.Value.ValueKind == JsonValueKind.String ? item.Value.GetString() ?? string.Empty : item.Value.ToString(),
                    StringComparer.OrdinalIgnoreCase)
                    ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            }
            catch
            {
                return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            }
        }

        private static string ReadValue(Dictionary<string, string> jsonValues, NRData row, params string[] keys)
        {
            foreach (var key in keys)
            {
                var property = row.GetType().GetProperty(key, BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);
                var propertyValue = NormalizeText(property?.GetValue(row)?.ToString());
                if (!string.IsNullOrWhiteSpace(propertyValue))
                {
                    return propertyValue;
                }

                if (jsonValues.TryGetValue(key, out var jsonValue))
                {
                    var normalizedJsonValue = NormalizeText(jsonValue);
                    if (!string.IsNullOrWhiteSpace(normalizedJsonValue))
                    {
                        return normalizedJsonValue;
                    }
                }
            }

            var normalizedKeys = keys
                .Select(key => key.Replace(" ", string.Empty))
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            var fallbackEntry = jsonValues
                .FirstOrDefault(item => normalizedKeys.Contains(item.Key.Replace(" ", string.Empty)) && !string.IsNullOrWhiteSpace(NormalizeText(item.Value)));

            return NormalizeText(fallbackEntry.Value);
        }

        private static int? ParseImportRowNo(Dictionary<string, string> jsonValues)
        {
            var rawValue = ReadDictionaryValue(
                jsonValues,
                "ImportRowNo",
                "Import Row No",
                "ImportRow",
                "Import Row",
                "ExcelRowNo",
                "Excel Row No",
                "RowNo",
                "Row No");

            return int.TryParse(rawValue, out var importRowNo) ? importRowNo : null;
        }

        private static string ReadDictionaryValue(Dictionary<string, string> jsonValues, params string[] keys)
        {
            foreach (var key in keys)
            {
                if (jsonValues.TryGetValue(key, out var directValue))
                {
                    var normalized = NormalizeText(directValue);
                    if (!string.IsNullOrWhiteSpace(normalized))
                    {
                        return normalized;
                    }
                }
            }

            var normalizedKeys = keys
                .Select(key => key.Replace(" ", string.Empty))
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            var fallbackEntry = jsonValues.FirstOrDefault(item =>
                normalizedKeys.Contains(item.Key.Replace(" ", string.Empty)) &&
                !string.IsNullOrWhiteSpace(NormalizeText(item.Value)));

            return NormalizeText(fallbackEntry.Value);
        }

        private static void BuildCollegeConflicts(
            List<ConflictReportItem> conflicts,
            List<NRDataConflictProjection> rows,
            string identifierType,
            string conflictType,
            string uniqueField,
            bool useNodalValues)
        {
            var groupedRows = rows
                .Where(item => !string.IsNullOrWhiteSpace(identifierType == "CollegeCode" ? item.CollegeCode : item.CollegeName))
                .GroupBy(item => identifierType == "CollegeCode" ? item.CollegeCode : item.CollegeName);

            foreach (var group in groupedRows)
            {
                var values = useNodalValues
                    ? DistinctNonEmpty(group.Select(item => item.NodalCode))
                    : DistinctNonEmpty(group.Select(item => item.CenterCode));

                if (values.Count <= 1)
                {
                    continue;
                }

                conflicts.Add(new ConflictReportItem
                {
                    ConflictType = conflictType,
                    CollegeKeyType = identifierType,
                    CollegeName = identifierType == "CollegeName" ? group.Key : null,
                    CollegeCode = identifierType == "CollegeCode" ? group.Key : null,
                    CatchNos = DistinctNonEmpty(group.Select(item => item.CatchNo)),
                    RowIds = group.Select(item => item.Row.Id).Distinct().OrderBy(id => id).ToList(),
                    UniqueField = uniqueField,
                    NodalCodes = useNodalValues ? values : new List<string>(),
                    CenterCodes = useNodalValues ? new List<string>() : values,
                    ConflictingValues = values,
                    CanIgnore = false,
                    CanResolve = true,
                    Summary = useNodalValues
                        ? $"College {group.Key} is linked with multiple nodal codes."
                        : $"College {group.Key} is linked with multiple exam centres."
                });
            }
        }

        private static List<string> DistinctNonEmpty(IEnumerable<string?> values)
        {
            return values
                .Select(NormalizeText)
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(value => value, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static string NormalizeText(string? value)
        {
            return string.IsNullOrWhiteSpace(value) ? string.Empty : value.Trim();
        }

        private static string NormalizeCode(string? value)
        {
            var normalized = NormalizeText(value);
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return string.Empty;
            }

            if (normalized.All(char.IsDigit))
            {
                var stripped = normalized.TrimStart('0');
                return string.IsNullOrWhiteSpace(stripped) ? "0" : stripped;
            }

            return normalized.ToUpperInvariant();
        }

        private static string GetNodalDigitSignature(string? value)
        {
            var normalized = NormalizeText(value);
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return string.Empty;
            }

            var digitsOnly = new string(normalized.Where(char.IsDigit).ToArray());
            if (!string.IsNullOrWhiteSpace(digitsOnly))
            {
                var stripped = digitsOnly.TrimStart('0');
                return string.IsNullOrWhiteSpace(stripped) ? "0" : stripped;
            }

            return NormalizeCode(normalized);
        }

        private static int GetDigitCount(string? value)
        {
            return NormalizeText(value).Count(char.IsDigit);
        }

        private static int NormalizeStatusCode(string? status)
        {
            var normalized = NormalizeText(status).ToLowerInvariant();
            return normalized switch
            {
                "resolved" => ConflictStatusResolved,
                "ignored" => ConflictStatusIgnored,
                _ => ConflictStatusPending,
            };
        }

        private static string NormalizeStatusLabel(int status)
        {
            return status switch
            {
                ConflictStatusResolved => "resolved",
                ConflictStatusIgnored => "ignored",
                _ => "pending",
            };
        }

        [HttpPost("missing-data")]
        public async Task<ActionResult> SaveMissingData([FromBody] MissingDataSaveRequest request)
        {
            if (request == null || request.ProjectId <= 0)
            {
                return BadRequest("Valid projectId is required.");
            }

            if (request.Data == null || !request.Data.Any())
            {
                return BadRequest("Missing data payload is required.");
            }

            var effectiveRows = request.Data
                .Where(x =>
                    HasMeaningfulText(x.CatchNo) &&
                    (HasMeaningfulNumber(x.Pages) ||
                     HasMeaningfulText(x.ExamDate) ||
                     HasMeaningfulText(x.ExamTime)))
                .ToList();

            if (effectiveRows.Count<=0)
            {
                return BadRequest("At least one value is required to save.");
            }

            try
            {
                var catchNumbers = effectiveRows
                    .Select(x => x.CatchNo!.Trim())
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();

                var projectRows = await _context.NRDatas
                    .Where(x => x.ProjectId == request.ProjectId && x.CatchNo != null && catchNumbers.Contains(x.CatchNo))
                    .ToListAsync();

                if (projectRows.Count<=0)
                {
                    return NotFound("No matching NRData rows found for the provided catch numbers.");
                }

                int updatedCatchCount = 0;
                int updatedRowCount = 0;

                foreach (var item in effectiveRows)
                {
                    var trimmedCatchNo = item.CatchNo!.Trim();
                    var matchingRows = projectRows
                        .Where(x => string.Equals(x.CatchNo, trimmedCatchNo, StringComparison.OrdinalIgnoreCase))
                        .ToList();

                    if (!matchingRows.Any())
                    {
                        continue;
                    }

                    updatedCatchCount++;

                    foreach (var row in matchingRows)
                    {
                        if (HasMeaningfulNumber(item.Pages))
                        {
                            row.Pages = item.Pages!.Value;
                        }

                        if (HasMeaningfulText(item.ExamDate))
                        {
                            row.ExamDate = item.ExamDate.Trim();
                        }

                        if (HasMeaningfulText(item.ExamTime))
                        {
                            row.ExamTime = item.ExamTime.Trim();
                        }

                        updatedRowCount++;
                    }
                }

                if (updatedRowCount == 0)
                {
                    return NotFound("No matching NRData rows found to update.");
                }

                await _context.SaveChangesAsync();

                _loggerService.LogEvent(
                    $"Saved missing data for ProjectId {request.ProjectId}. Catches updated: {updatedCatchCount}, rows updated: {updatedRowCount}",
                    "NRData",
                    User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0,
                    request.ProjectId);

                return Ok(new
                {
                    message = "Missing data saved successfully",
                    updatedCatchCount,
                    updatedRowCount
                });
            }
            catch (Exception ex)
            {
                _loggerService.LogError("Error saving missing data", ex.Message, nameof(NRDatasController));
                return StatusCode(500, "Internal Server Error");
            }
        }

        public class ConflictActionDto
        {
            public string? ConflictType { get; set; }
            public string? CatchNo { get; set; }
            public List<string> CatchNos { get; set; } = new();
            public List<int> RowIds { get; set; } = new();
            public List<int> ImportRowNos { get; set; } = new();
            public string? UniqueField { get; set; }
            public string? Field { get; set; }
            public string? CentreCode { get; set; }
            public string? NodalCode { get; set; }
            public string? NodalCodeGroup { get; set; }
            public string? CollegeName { get; set; }
            public string? CollegeCode { get; set; }
            public string? CollegeKeyType { get; set; }
            public List<string> ConflictingValues { get; set; } = new();
            public List<string> NodalCodes { get; set; } = new();
            public List<string> CenterCodes { get; set; } = new();
            public int? MinNrQuantity { get; set; }
            public int? MaxNrQuantity { get; set; }
        }

        public class ConflictResolutionDto : ConflictActionDto
        {
            public string? SelectedValue { get; set; }
        }

        public class ConflictStatusUpdateDto : ConflictActionDto
        {
            public string? Status { get; set; }
        }

        public class ConflictReportItem : ConflictActionDto
        {
            public bool CanIgnore { get; set; }
            public bool CanResolve { get; set; }
            public string Status { get; set; } = "pending";
            public string? Summary { get; set; }
        }

        private static Dictionary<string, object> BuildConflictStoragePayload(ConflictActionDto payload)
        {
            var conflictType = NormalizeText(payload.ConflictType);
            var storagePayload = new Dictionary<string, object>
            {
                ["ConflictType"] = conflictType,
            };

            AddIfHasText(storagePayload, "UniqueField", payload.UniqueField);
            AddIfHasText(storagePayload, "Field", payload.Field);

            switch (conflictType)
            {
                case CatchUniqueFieldConflict:
                    AddIfHasText(storagePayload, "CatchNo", payload.CatchNo);
                    AddIfHasItems(storagePayload, "ConflictingValues", DistinctNonEmpty(payload.ConflictingValues ?? new List<string>()));
                    break;
                case CenterMultipleNodalsConflict:
                    AddIfHasText(storagePayload, "CentreCode", payload.CentreCode);
                    AddIfHasItems(storagePayload, "ConflictingValues", DistinctNonEmpty(payload.ConflictingValues ?? new List<string>()));
                    break;
                case CollegeMultipleNodalsConflict:
                case CollegeMultipleCentersConflict:
                    AddIfHasText(storagePayload, "CollegeName", payload.CollegeName);
                    AddIfHasText(storagePayload, "CollegeCode", payload.CollegeCode);
                    AddIfHasText(storagePayload, "CollegeKeyType", payload.CollegeKeyType);
                    AddIfHasItems(storagePayload, "ConflictingValues", DistinctNonEmpty(payload.ConflictingValues ?? new List<string>()));
                    break;
                case RequiredFieldEmptyConflict:
                    AddIfHasItems(storagePayload, "ImportRowNos", (payload.ImportRowNos ?? new List<int>()).Distinct().OrderBy(id => id).ToList());
                    break;
                case ZeroNrQuantityConflict:
                    AddIfHasItems(storagePayload, "CatchNos", DistinctNonEmpty(payload.CatchNos ?? new List<string>()));
                    AddIfHasItems(storagePayload, "ImportRowNos", (payload.ImportRowNos ?? new List<int>()).Distinct().OrderBy(id => id).ToList());
                    break;
                case NodalCodeDigitMismatchConflict:
                    AddIfHasText(storagePayload, "NodalCodeGroup", payload.NodalCodeGroup);
                    AddIfHasItems(storagePayload, "ImportRowNos", (payload.ImportRowNos ?? new List<int>()).Distinct().OrderBy(id => id).ToList());
                    AddIfHasText(storagePayload, "NodalCode", payload.NodalCode);
                    AddIfHasItems(storagePayload, "ConflictingValues", DistinctNonEmpty(payload.ConflictingValues ?? new List<string>()));
                    break;
                default:
                    AddIfHasText(storagePayload, "CatchNo", payload.CatchNo);
                    AddIfHasItems(storagePayload, "CatchNos", DistinctNonEmpty(payload.CatchNos ?? new List<string>()));
                    AddIfHasItems(storagePayload, "RowIds", (payload.RowIds ?? new List<int>()).Distinct().OrderBy(id => id).ToList());
                    AddIfHasText(storagePayload, "CentreCode", payload.CentreCode);
                    AddIfHasText(storagePayload, "CollegeName", payload.CollegeName);
                    AddIfHasText(storagePayload, "CollegeCode", payload.CollegeCode);
                    AddIfHasText(storagePayload, "CollegeKeyType", payload.CollegeKeyType);
                    AddIfHasText(storagePayload, "NodalCodeGroup", payload.NodalCodeGroup);
                    AddIfHasItems(storagePayload, "ConflictingValues", DistinctNonEmpty(payload.ConflictingValues ?? new List<string>()));
                    break;
            }

            return storagePayload;
        }

        private static void AddIfHasText(Dictionary<string, object> payload, string key, string? value)
        {
            var normalized = NormalizeText(value);
            if (!string.IsNullOrWhiteSpace(normalized))
            {
                payload[key] = normalized;
            }
        }

        private static void AddIfHasItems<T>(Dictionary<string, object> payload, string key, List<T> values)
        {
            if (values != null && values.Count > 0)
            {
                payload[key] = values;
            }
        }

        private class NRDataConflictProjection
        {
            public required NRData Row { get; set; }
            public string CatchNo { get; set; } = string.Empty;
            public string ExamDate { get; set; } = string.Empty;
            public string ExamTime { get; set; } = string.Empty;
            public string CenterCode { get; set; } = string.Empty;
            public string NodalCode { get; set; } = string.Empty;
            public int NrQuantity { get; set; }
            public int? ImportRowNo { get; set; }
            public string CollegeName { get; set; } = string.Empty;
            public string CollegeCode { get; set; } = string.Empty;
        }

        public class MissingDataSaveRequest
        {
            public int ProjectId { get; set; }
            public List<MissingDataItem> Data { get; set; } = new();
        }

        public class MissingDataItem
        {
            public string? CatchNo { get; set; }
            public int? Pages { get; set; }
            public string? ExamDate { get; set; }
            public string? ExamTime { get; set; }
        }

        public class MergeCatchNoRequest
        {
            public List<int> RowIds { get; set; } = new();
            public List<string> CatchNos { get; set; } = new();
            public string? Separator { get; set; }
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
                var nrDataList = await _context.NRDatas
                    .Where(d => d.ProjectId == ProjectId)
                    .ToListAsync();

                var conflictList = await _context.ConflictingFields
                    .Where(c => c.ProjectId == ProjectId)
                    .ToListAsync();

                if (!nrDataList.Any())
                {
                    return NotFound($"No NRData found for ProjectId {ProjectId}");
                }

                var reportPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", ProjectId.ToString());

                if (Directory.Exists(reportPath))
                {
                    Directory.Delete(reportPath, true);
                }

                _context.NRDatas.RemoveRange(nrDataList);

                if (conflictList.Any())
                {
                    _context.ConflictingFields.RemoveRange(conflictList);
                }

                await _context.SaveChangesAsync();

                int userId = 0;
                int.TryParse(User.Identity?.Name, out userId);

                _loggerService.LogEvent(
                    $"Deleted all NRData and conflict records for ProjectId {ProjectId}",
                    "NRData",
                    userId,
                    ProjectId
                );

                return NoContent();
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.ToString());
            }
        }
        private bool NRDataExists(int id)
        {
            return _context.NRDatas.Any(e => e.Id == id);
        }

        private static bool HasMeaningfulText(string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            var normalized = value.Trim();
            return !string.Equals(normalized, "undefined", StringComparison.OrdinalIgnoreCase)
                && !string.Equals(normalized, "null", StringComparison.OrdinalIgnoreCase);
        }

        private static bool HasMeaningfulNumber(int? value)
        {
            return value.HasValue && value.Value > 0;
        }
    }
}
