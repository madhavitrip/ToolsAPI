using ERPToolsAPI.Data;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Options;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.Json;
using System.Threading.Tasks;
using Tools.Models;
using Tools.Services;
using ToolsAPI.Helpers;
using ToolsAPI.Models;

namespace Tools.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class NrData1Controller : ControllerBase
    {
        private readonly ERPToolsDbContext _context;
        private readonly ILoggerService _loggerService;
        private readonly IOptions<ApiSettings> _apiSettings;
        private readonly IDispatchService _dispatchService;

        public NrData1Controller(
            ERPToolsDbContext context,
            ILoggerService loggerService,
            IOptions<ApiSettings> apiSettings,
            IDispatchService dispatchService)
        {
            _context = context;
            _loggerService = loggerService;
            _apiSettings = apiSettings;
            _dispatchService = dispatchService;
        }

        // GET: api/NrData1
        [HttpGet]
        public async Task<ActionResult<IEnumerable<NrData1>>> GetNrData1()
        {
            return await _context.NrData1.ToListAsync();
        }

        // GET: api/NrData1/5
        [HttpGet("{id}")]
        public async Task<ActionResult> GetNrData1(int id)
        {
            var record = await _context.NrData1.FindAsync(id);
            if (record == null)
            {
                return NotFound();
            }

            var centers = await _context.CenterList
                .Where(c => c.NRDataId == id && c.Status)
                .ToListAsync();

            return Ok(new
            {
                NrData = record,
                Centers = centers
            });
        }

        // GET: api/NrData1/GetByProjectId/5
        [HttpGet("GetByProjectId/{projectId}")]
        public async Task<ActionResult> GetByProjectId(
            int projectId,
            [FromQuery] int? batchNo = null,
            [FromQuery] int pageNo = 1,
            [FromQuery] int pageSize = 50,
            [FromQuery] string? search = null)
        {
            var query = _context.NrData1
                .Where(d => d.ProjectId == projectId);

            if (batchNo.HasValue)
            {
                query = query.Where(d => d.Batch == batchNo.Value);
            }

            if (!string.IsNullOrWhiteSpace(search))
            {
                string s = search.Trim().ToLower();
                query = query.Where(d =>
                    (d.CatchNo != null && d.CatchNo.ToLower().Contains(s)) ||
                    (d.CourseName != null && d.CourseName.ToLower().Contains(s)) ||
                    (d.SubjectName != null && d.SubjectName.ToLower().Contains(s)));
            }

            int totalRecords = await query.CountAsync();
            var records = await query
                .OrderBy(d => d.Id)
                .Skip((pageNo - 1) * pageSize)
                .Take(pageSize)
                .ToListAsync();

            return Ok(new
            {
                Total = totalRecords,
                PageNo = pageNo,
                PageSize = pageSize,
                Data = records
            });
        }

        // GET: api/NrData1/centers/5
        [HttpGet("centers/{nrDataId}")]
        public async Task<ActionResult<IEnumerable<CenterList>>> GetCentersByNrDataId(int nrDataId)
        {
            var centers = await _context.CenterList
                .Where(c => c.NRDataId == nrDataId && c.Status)
                .ToListAsync();

            return Ok(centers);
        }

        // POST: api/NrData1
        [HttpPost]
        [HttpPost("PostNRData")]
        public async Task<IActionResult> PostNRData([FromBody] JsonElement inputData, [FromQuery] int? lotNo = null)
        {
            if (!ModelState.IsValid)
                return BadRequest(ModelState);

            try
            {
                int projectId = inputData.GetProperty("projectId").GetInt32();
                var incomingData = inputData.GetProperty("data");

                bool isChangedNR = false;
                if (inputData.TryGetProperty("isChangedNR", out JsonElement isChangedNRElem) && isChangedNRElem.ValueKind == JsonValueKind.True)
                {
                    isChangedNR = true;
                }

                // 1. Fetch project configs and fields
                var projectConfig = await _context.ProjectConfigs
                    .FirstOrDefaultAsync(x => x.ProjectId == projectId);

                var extraConfigs = await _context.ExtraConfigurations
                    .Where(x => x.ProjectId == projectId)
                    .ToListAsync();

                var allFields = await _context.Fields
                    .AsNoTracking()
                    .ToListAsync();

                // Unique fields from Fields table (where IsUnique is true)
                var uniqueFieldsFromDb = allFields
                    .Where(f => f.IsUnique)
                    .Select(f => f.Name.Trim())
                    .ToList();

                var uniqueFieldSet = new HashSet<string>(uniqueFieldsFromDb, StringComparer.OrdinalIgnoreCase)
                {
                    "CatchNo" // CatchNo is always part of unique criteria
                };

                // Dynamic matching fields if duplicate criteria is configured
                var duplicateFieldIds = projectConfig?.DuplicateCriteria ?? new List<int>();
                var duplicateCriteriaFields = allFields
                    .Where(f => duplicateFieldIds.Contains(f.FieldId))
                    .Select(f => f.Name.Trim())
                    .ToList();

                // 2. Reflection Property Maps
                var nrData1Properties = typeof(NrData1)
                    .GetProperties()
                    .ToDictionary(p => p.Name.ToLower(), p => p);

                var centerListProperties = typeof(CenterList)
                    .GetProperties()
                    .ToDictionary(p => p.Name.ToLower(), p => p);

                // Center-specific property names that should never be mapped to NrData1
                var centerSpecificNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                {
                    "centercode", "nodalcode", "route", "routesort", "centersort",
                    "nodalsort", "district", "districtsort", "centerdata", "status",
                    "centername", "collegename", "collegecode"
                };

                // 3. Batch determination
                var existingNrData1List = await _context.NrData1
                    .Where(x => x.ProjectId == projectId)
                    .ToListAsync();

                int batchId = 1;
                if (existingNrData1List.Any() && isChangedNR)
                {
                    batchId = existingNrData1List.Max(x => x.Batch) + 1;
                }

                // 4. Parse incoming rows and group by Catch / Unique fields
                // We keep a data structure per unique catch:
                // CatchGroup: NrData1 prototype, accumulated Quantity, accumulated NRQuantity, and list of CenterList records
                var catchGroups = new Dictionary<string, CatchGroupData>(StringComparer.OrdinalIgnoreCase);
                var extraEnvelopesToAdd = new List<ExtraEnvelopes>();

                // Fetch all ExtraTypes from database to avoid hardcoding names
                var allExtraTypes = await _context.ExtraType.ToListAsync();
                var extraTypesByName = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                foreach (var et in allExtraTypes)
                {
                    if (!string.IsNullOrWhiteSpace(et.Type))
                    {
                        extraTypesByName[et.Type.Trim()] = et.ExtraTypeId;
                        if (!et.Type.EndsWith("Extra", StringComparison.OrdinalIgnoreCase))
                        {
                            extraTypesByName[$"{et.Type.Trim()} Extra"] = et.ExtraTypeId;
                        }
                    }
                }

                int totalItems = incomingData.GetArrayLength();

                for (int i = 0; i < totalItems; i++)
                {
                    var item = incomingData[i];

                    // Extract catch number and center code
                    string itemCatchNo = "";
                    string itemCenterCode = "";
                    int itemQty = 0;
                    int itemNrQty = 0;

                    var uniqueExtraData = new Dictionary<string, string>();
                    var centerExtraData = new Dictionary<string, string>();

                    // Temporary dictionaries for mapping properties
                    var nrData1FieldValues = new Dictionary<PropertyInfo, object?>();
                    var centerListFieldValues = new Dictionary<PropertyInfo, object?>();

                    foreach (var prop in item.EnumerateObject())
                    {
                        string rawKey = prop.Name;
                        string key = rawKey.Replace(" ", "").ToLower();
                        string value = prop.Value.ToString().Trim();

                        if (key == "catchno") itemCatchNo = value;
                        if (key == "centercode") itemCenterCode = value;
                        if (key == "quantity" && int.TryParse(value, out int qVal)) itemQty = qVal;
                        if (key == "nrquantity" && int.TryParse(value, out int nrqVal)) itemNrQty = nrqVal;

                        // Check if this property is a UNIQUE field:
                        // 1. In uniqueFieldSet (IsUnique == true in Fields table)
                        // 2. OR is a typed property on NrData1 (and NOT center-specific)
                        bool isFieldUnique = uniqueFieldSet.Contains(rawKey) ||
                                             uniqueFieldSet.Contains(key) ||
                                             (nrData1Properties.ContainsKey(key) && !centerSpecificNames.Contains(key));

                        if (isFieldUnique)
                        {
                            // Unique field -> belongs to NrData1
                            if (nrData1Properties.TryGetValue(key, out var propInfo))
                            {
                                if (propInfo.Name.Equals("Id", StringComparison.OrdinalIgnoreCase) ||
                                    propInfo.Name.Equals("ProjectId", StringComparison.OrdinalIgnoreCase))
                                {
                                    continue;
                                }

                                var converted = ConvertPropertyValue(propInfo, value);
                                if (converted != null)
                                {
                                    nrData1FieldValues[propInfo] = converted;
                                }
                            }
                            else if (rawKey != "projectId" && rawKey != "isCorrectedNrdataReport")
                            {
                                uniqueExtraData[rawKey] = value;
                            }
                        }
                        else
                        {
                            // Non-unique field -> belongs to CenterList ("rest details")
                            if (centerListProperties.TryGetValue(key, out var propInfo))
                            {
                                if (propInfo.Name.Equals("Id", StringComparison.OrdinalIgnoreCase) ||
                                    propInfo.Name.Equals("ProjectId", StringComparison.OrdinalIgnoreCase) ||
                                    propInfo.Name.Equals("NRDataId", StringComparison.OrdinalIgnoreCase))
                                {
                                    continue;
                                }

                                var converted = ConvertPropertyValue(propInfo, value);
                                if (converted != null)
                                {
                                    centerListFieldValues[propInfo] = converted;
                                }
                            }
                            else if (rawKey != "projectId" && rawKey != "isCorrectedNrdataReport")
                            {
                                centerExtraData[rawKey] = value;
                            }
                        }
                    }

                    // Ensure Quantity and NRQuantity for this center are preserved in CenterData
                    if (!centerExtraData.ContainsKey("Quantity") && itemQty > 0)
                        centerExtraData["Quantity"] = itemQty.ToString();
                    if (!centerExtraData.ContainsKey("NRQuantity") && itemNrQty > 0)
                        centerExtraData["NRQuantity"] = itemNrQty.ToString();

                    // Determine unique grouping key (CatchNo, or composite matching fields)
                    string catchKey = !string.IsNullOrWhiteSpace(itemCatchNo) ? itemCatchNo.Trim() : $"ROW_{i}";

                    if (!catchGroups.TryGetValue(catchKey, out var catchGroup))
                    {
                        var nrData1 = new NrData1
                        {
                            ProjectId = projectId,
                            CatchNo = itemCatchNo,
                            Batch = batchId,
                            Steps = PipelineNavigator.STEP_UPLOADED,
                          
                        };

                        if (lotNo.HasValue)
                        {
                            nrData1.EnvLotNo = lotNo.Value;
                        }

                        // Apply typed unique properties
                        foreach (var kvp in nrData1FieldValues)
                        {
                            if (kvp.Key.CanWrite && kvp.Key.Name != "Quantity" && kvp.Key.Name != "NRQuantity")
                            {
                                kvp.Key.SetValue(nrData1, kvp.Value);
                            }
                        }

                        // Calculate Day from ExamDate
                        if (!string.IsNullOrWhiteSpace(nrData1.ExamDate))
                        {
                            if (DateTime.TryParseExact(
                                nrData1.ExamDate.Trim(),
                                "dd-MM-yyyy",
                                CultureInfo.InvariantCulture,
                                DateTimeStyles.None,
                                out DateTime examDate))
                            {
                                nrData1.Day = examDate.ToString("dddd", CultureInfo.InvariantCulture);
                            }
                            else if (DateTime.TryParse(nrData1.ExamDate, out DateTime fallbackDate))
                            {
                                nrData1.Day = fallbackDate.DayOfWeek.ToString();
                            }
                        }

                        catchGroup = new CatchGroupData
                        {
                            NrData1 = nrData1,
                            UniqueExtraData = new Dictionary<string, string>(uniqueExtraData, StringComparer.OrdinalIgnoreCase)
                        };
                        catchGroups[catchKey] = catchGroup;
                    }
                    else
                    {
                        // Merge any additional unique dynamic data
                        foreach (var kvp in uniqueExtraData)
                        {
                            if (!catchGroup.UniqueExtraData.ContainsKey(kvp.Key))
                            {
                                catchGroup.UniqueExtraData[kvp.Key] = kvp.Value;
                            }
                        }
                    }
               
                    // Build CenterList record
                    bool isExtraRow = !string.IsNullOrWhiteSpace(itemCenterCode) && extraTypesByName.ContainsKey(itemCenterCode.Trim());
                    int? extraTypeId = isExtraRow ? extraTypesByName[itemCenterCode.Trim()] : null;

                    var centerList = new CenterList
                    {
                        ProjectId = projectId,
                        CenterCode = itemCenterCode,
                        Status = true,
                        ExtraId = extraTypeId ?? 0,
                        CenterData = centerExtraData.Any() ? JsonSerializer.Serialize(centerExtraData) : "{}"
                    };

                    // Apply typed center properties
                    foreach (var kvp in centerListFieldValues)
                    {
                        if (kvp.Key.CanWrite)
                        {
                            kvp.Key.SetValue(centerList, kvp.Value);
                        }
                    }

                    if (!isExtraRow)
                    {
                        catchGroup.Centers.Add(centerList);
                    }

                    // Process Extra Envelopes
                    var extraEnvelopes = ProcessExtraEnvelopes(
                        itemCatchNo,
                        itemCenterCode,
                        itemQty,
                        itemNrQty,
                        projectId,
                        extraConfigs,
                        extraTypesByName
                    );
                    extraEnvelopesToAdd.AddRange(extraEnvelopes);
                }

                // 5. Finalize NrData1 dynamic JSON serialization
                var nrData1List = new List<NrData1>();
                var allCenterListToAdd = new List<CenterList>();

                foreach (var group in catchGroups.Values)
                {
                    if (group.UniqueExtraData.Any())
                    {
                        group.NrData1.NRDatas = JsonSerializer.Serialize(group.UniqueExtraData);
                    }
                    nrData1List.Add(group.NrData1);
                }

                // 6. If re-uploading (Batch > 1), deactivate existing active center records for catches being updated
                if (batchId > 1 && existingNrData1List.Any())
                {
                    var incomingCatchNos = nrData1List
                        .Select(x => x.CatchNo)
                        .Where(x => !string.IsNullOrWhiteSpace(x))
                        .ToHashSet(StringComparer.OrdinalIgnoreCase);

                    var previousNrDataIds = existingNrData1List
                        .Where(x => x.CatchNo != null && incomingCatchNos.Contains(x.CatchNo))
                        .Select(x => x.Id)
                        .ToList();

                    if (previousNrDataIds.Any())
                    {
                        var oldCenters = await _context.CenterList
                            .Where(c => c.ProjectId == projectId && previousNrDataIds.Contains(c.NRDataId) && c.Status)
                            .ToListAsync();

                        foreach (var oldCenter in oldCenters)
                        {
                            oldCenter.Status = false;
                        }
                    }
                }

                // 7. Save NrData1 records first to generate their Identity IDs
                if (nrData1List.Any())
                {
                    await _context.NrData1.AddRangeAsync(nrData1List);
                    await _context.SaveChangesAsync();
                }

                // 8. Associate NRDataId with CenterList records and save them
                foreach (var group in catchGroups.Values)
                {
                    int generatedNrDataId = group.NrData1.Id;
                    foreach (var center in group.Centers)
                    {
                        center.NRDataId = generatedNrDataId;
                        allCenterListToAdd.Add(center);
                    }
                }

                if (allCenterListToAdd.Any())
                {
                    await _context.CenterList.AddRangeAsync(allCenterListToAdd);
                }

                // 9. Process and save Extra Envelopes
                var groupedExtras = GroupAndProcessExtras(
                    extraEnvelopesToAdd,
                    projectId,
                    extraConfigs
                );

                if (groupedExtras.Any())
                {
                    // Deactivate overlapping existing extras
                    var existingExtras = await _context.ExtrasEnvelope
                        .Where(e => e.ProjectId == projectId && e.Status == 1)
                        .ToListAsync();

                    foreach (var ex in existingExtras)
                    {
                        if (groupedExtras.Any(ge => ge.ExtraId == ex.ExtraId && ge.CatchNo == ex.CatchNo))
                        {
                            ex.Status = 0;
                            _context.ExtrasEnvelope.Update(ex);
                        }
                    }

                    await _context.ExtrasEnvelope.AddRangeAsync(groupedExtras);
                }

                await _context.SaveChangesAsync();

                // 10. Logging
                await _loggerService.LogEventAsync(
                    batchId == 1
                        ? $"Initial Upload Processed (Batch 1) into nrdata1 and centerlist"
                        : $"Upload Processed (Batch {batchId}) into nrdata1 and centerlist",
                    "NrData1",
                    LogHelper.GetTriggeredBy(User),
                    projectId
                );

                return Ok(new
                {
                    message = batchId == 1 ? "Initial data uploaded successfully" : "Data processed successfully",
                    Batch = batchId,
                    NewCatches = nrData1List.Count,
                    NewCenters = allCenterListToAdd.Count
                });
            }
            catch (Exception ex)
            {
                await _loggerService.LogErrorAsync(
                    "NrData1 Upload Error",
                    ex.ToString(),
                    nameof(NrData1Controller)
                );

                return StatusCode(500, ex.Message);
            }
        }

        private static string NormalizeKey(string? s)
        {
            if (string.IsNullOrWhiteSpace(s)) return "";
            return s.Replace(" ", "").Replace("_", "").ToLowerInvariant().Trim();
        }

        private static string GetFieldValue(CenterList c, NrData1? p, string fieldName)
        {
            var target = NormalizeKey(fieldName);

            // 1. Check CenterList direct properties
            var cProp = typeof(CenterList).GetProperties(BindingFlags.Public | BindingFlags.Instance)
                .FirstOrDefault(pr => NormalizeKey(pr.Name) == target);
            if (cProp != null)
            {
                var val = cProp.GetValue(c)?.ToString()?.Trim();
                if (!string.IsNullOrEmpty(val)) return val;
            }

            // 2. Check CenterList.CenterData JSON
            if (!string.IsNullOrWhiteSpace(c.CenterData))
            {
                try
                {
                    var dict = JsonSerializer.Deserialize<Dictionary<string, string>>(c.CenterData);
                    if (dict != null)
                    {
                        var kvp = dict.FirstOrDefault(k => NormalizeKey(k.Key) == target);
                        if (!string.IsNullOrEmpty(kvp.Value)) return kvp.Value.Trim();
                    }
                }
                catch { }
            }

            // 3. Check NrData1 direct properties
            if (p != null)
            {
                var pProp = typeof(NrData1).GetProperties(BindingFlags.Public | BindingFlags.Instance)
                    .FirstOrDefault(pr => NormalizeKey(pr.Name) == target);
                if (pProp != null)
                {
                    var val = pProp.GetValue(p)?.ToString()?.Trim();
                    if (!string.IsNullOrEmpty(val)) return val;
                }

                // 4. Check NrData1.NRDatas JSON
                if (!string.IsNullOrWhiteSpace(p.NRDatas))
                {
                    try
                    {
                        var dict = JsonSerializer.Deserialize<Dictionary<string, string>>(p.NRDatas);
                        if (dict != null)
                        {
                            var kvp = dict.FirstOrDefault(k => NormalizeKey(k.Key) == target);
                            if (!string.IsNullOrEmpty(kvp.Value)) return kvp.Value.Trim();
                        }
                    }
                    catch { }
                }
            }

            return "";
        }

        [HttpPost("MergeFields")]
        public async Task<IActionResult> MergeFields(int ProjectId, int? batchId = null, int? lotNo = null)
        {
            try
            {
                Console.WriteLine($"[NrData1Controller] MergeFields received ProjectId: {ProjectId}, batchId: {batchId}, lotNo: {lotNo}");

                IQueryable<NrData1> query = _context.NrData1
                    .Where(p => p.ProjectId == ProjectId);

                // 1. Try with batchId and lotNo if provided, prioritizing STEP_UPLOADED (0)
                var attemptQuery = query;
                if (batchId.HasValue && batchId.Value > 0)
                {
                    Console.WriteLine($"[NrData1Controller] Filtering by batchId: {batchId.Value}");
                    attemptQuery = attemptQuery.Where(p => p.Batch == batchId.Value);
                }
                if (lotNo.HasValue && lotNo.Value > 0)
                {
                    Console.WriteLine($"[NrData1Controller] Filtering by lotNo: {lotNo.Value}");
                    attemptQuery = attemptQuery.Where(p => p.LotNo == lotNo.Value || p.EnvLotNo == lotNo.Value);
                }

                var data = await attemptQuery.Where(p => p.Steps == PipelineNavigator.STEP_UPLOADED).ToListAsync();

                // 2. If no records found with lot filter, try with batchId + STEP_UPLOADED (records might have LotNo = 0)
                if (!data.Any() && lotNo.HasValue && lotNo.Value > 0)
                {
                    var batchQuery = query;
                    if (batchId.HasValue && batchId.Value > 0)
                    {
                        batchQuery = batchQuery.Where(p => p.Batch == batchId.Value);
                    }
                    data = await batchQuery.Where(p => p.Steps == PipelineNavigator.STEP_UPLOADED).ToListAsync();
                }

                // 3. If still no records found, find ANY records with STEP_UPLOADED for this ProjectId
                if (!data.Any())
                {
                    data = await query.Where(p => p.Steps == PipelineNavigator.STEP_UPLOADED).ToListAsync();
                }

                // 4. Fallback for re-running on specific batch if no STEP_UPLOADED found
                if (!data.Any() && batchId.HasValue && batchId.Value > 0)
                {
                    data = await attemptQuery.ToListAsync();
                    if (!data.Any())
                    {
                        data = await query.Where(p => p.Batch == batchId.Value).ToListAsync();
                    }
                }

                // 5. Fallback for re-running across project
                if (!data.Any())
                {
                    data = await query.ToListAsync();
                }

                Console.WriteLine($"[NrData1Controller] Found {data.Count} records to process");

                if (!data.Any())
                    return BadRequest("No data found to process duplication. All data is already processed or no new data added.");

                var hasPendingDuplicateRows = data.Any(d => d.Steps == PipelineNavigator.STEP_UPLOADED);
                var projectconfig = await _context.ProjectConfigs
                    .FirstOrDefaultAsync(p => p.ProjectId == ProjectId);
                if (projectconfig == null)
                    return NotFound("Project config not exists for this project");

                var mergeFieldIds = projectconfig.DuplicateCriteria ?? new List<int>();
                if (!mergeFieldIds.Any())
                    return BadRequest("Duplicate criteria is not configured for this project.");

                var allFields = await _context.Fields.AsNoTracking().ToListAsync();
                var fieldNames = allFields
                    .Where(f => mergeFieldIds.Contains(f.FieldId))
                    .Select(f => f.Name.Trim())
                    .ToList();

                if (!fieldNames.Any())
                    return BadRequest("Duplicate criteria fields not found.");

                // Map of all catches for this project
                var nrDataIds = data.Select(d => d.Id).ToList();
                var nrDataMap = await _context.NrData1
                    .Where(d => d.ProjectId == ProjectId)
                    .ToDictionaryAsync(d => d.Id);

                // Fetch active regular CenterList records (ExtraId == 0) for these catches
                var activeCenters = await _context.CenterList
                    .Where(c => c.ProjectId == ProjectId && nrDataIds.Contains(c.NRDataId) && c.Status && c.ExtraId == 0)
                    .ToListAsync();

                if (!activeCenters.Any())
                {
                    return BadRequest("No active center list data found to process duplication.");
                }

                // Group active centers by duplicate criteria
                var centerGroups = activeCenters.GroupBy(c =>
                {
                    var p = nrDataMap.TryGetValue(c.NRDataId, out var nr) ? nr : null;
                    var keyParts = fieldNames.Select(fn => NormalizeKey(GetFieldValue(c, p, fn)));
                    return string.Join("|", keyParts);
                }).ToList();

                int mergedCount = 0;
                var deletedCenters = new List<CenterList>();
                var newCenters = new List<CenterList>();
                var affectedNrDataIds = new HashSet<int>();

                foreach (var cg in centerGroups)
                {
                    if (cg.Count() <= 1)
                        continue;

                    var firstCenter = cg.First();
                    int totalNrQty = cg.Sum(c => c.NRQuantity);
                    int finalQty = Math.Max(totalNrQty, cg.Max(c => c.Quantity));

                    // Merge dynamic CenterData JSON
                    var mergedCenterData = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    var remarksList = new List<string>();
                    var importRowNos = new List<string>();

                    foreach (var c in cg)
                    {
                        if (!string.IsNullOrWhiteSpace(c.CenterData))
                        {
                            try
                            {
                                var dict = JsonSerializer.Deserialize<Dictionary<string, string>>(c.CenterData);
                                if (dict != null)
                                {
                                    foreach (var kvp in dict)
                                    {
                                        if (string.Equals(kvp.Key, "Quantity", StringComparison.OrdinalIgnoreCase) ||
                                            string.Equals(kvp.Key, "NRQuantity", StringComparison.OrdinalIgnoreCase))
                                            continue;

                                        if (string.Equals(kvp.Key, "Remark", StringComparison.OrdinalIgnoreCase) ||
                                            string.Equals(kvp.Key, "Remarks", StringComparison.OrdinalIgnoreCase))
                                        {
                                            if (!string.IsNullOrWhiteSpace(kvp.Value))
                                                remarksList.Add(kvp.Value.Trim());
                                            continue;
                                        }

                                        if (string.Equals(kvp.Key, "ImportRowNo", StringComparison.OrdinalIgnoreCase))
                                        {
                                            if (!string.IsNullOrWhiteSpace(kvp.Value))
                                                importRowNos.Add(kvp.Value.Trim());
                                            continue;
                                        }

                                        if (!mergedCenterData.ContainsKey(kvp.Key) || string.IsNullOrWhiteSpace(mergedCenterData[kvp.Key]))
                                        {
                                            mergedCenterData[kvp.Key] = kvp.Value;
                                        }
                                    }
                                }
                            }
                            catch { }
                        }
                    }

                    if (remarksList.Any())
                        mergedCenterData["Remark"] = string.Join(" / ", remarksList.Distinct(StringComparer.OrdinalIgnoreCase));
                    if (importRowNos.Any())
                        mergedCenterData["ImportRowNo"] = string.Join(" / ", importRowNos.Distinct(StringComparer.OrdinalIgnoreCase));
                    mergedCenterData["Quantity"] = finalQty.ToString();
                    mergedCenterData["NRQuantity"] = totalNrQty.ToString();

                    int targetNrDataId = firstCenter.NRDataId;

                    // Create new merged CenterList record
                    var newMergedCenter = new CenterList
                    {
                        ProjectId = ProjectId,
                        NRDataId = targetNrDataId,
                        CenterCode = firstCenter.CenterCode,
                        NodalCode = firstCenter.NodalCode,
                        Route = firstCenter.Route,
                        RouteSort = firstCenter.RouteSort,
                        CenterSort = firstCenter.CenterSort,
                        NodalSort = firstCenter.NodalSort,
                        District = firstCenter.District,
                        DistrictSort = firstCenter.DistrictSort,
                        ExtraId = 0,
                        Quantity = finalQty,
                        NRQuantity = totalNrQty,
                        Status = true,
                        CenterData = JsonSerializer.Serialize(mergedCenterData)
                    };

                    await _context.CenterList.AddAsync(newMergedCenter);
                    newCenters.Add(newMergedCenter);

                    // Deactivate old duplicate centers
                    foreach (var c in cg)
                    {
                        c.Status = false;
                        _context.CenterList.Update(c);
                        deletedCenters.Add(c);
                        affectedNrDataIds.Add(c.NRDataId);
                    }

                    // Check if duplicate centers belonged to multiple distinct catches
                    var distinctCatchIds = cg.Select(c => c.NRDataId).Distinct().ToList();
                    if (distinctCatchIds.Count > 1)
                    {
                        var primaryCatchId = targetNrDataId;
                        var secondaryCatchIds = distinctCatchIds.Where(id => id != primaryCatchId).ToList();

                        if (nrDataMap.TryGetValue(primaryCatchId, out var primaryCatch))
                        {
                            var secondaryCatches = secondaryCatchIds
                                .Where(id => nrDataMap.ContainsKey(id))
                                .Select(id => nrDataMap[id])
                                .ToList();

                            var allSubjects = new[] { primaryCatch.SubjectName }
                                .Concat(secondaryCatches.Select(s => s.SubjectName))
                                .Where(s => !string.IsNullOrWhiteSpace(s))
                                .Distinct(StringComparer.OrdinalIgnoreCase)
                                .ToList();
                            if (allSubjects.Count > 1)
                                primaryCatch.SubjectName = string.Join(" / ", allSubjects);

                            var allCourses = new[] { primaryCatch.CourseName }
                                .Concat(secondaryCatches.Select(s => s.CourseName))
                                .Where(s => !string.IsNullOrWhiteSpace(s))
                                .Distinct(StringComparer.OrdinalIgnoreCase)
                                .ToList();
                            if (allCourses.Count > 1)
                                primaryCatch.CourseName = string.Join(" / ", allCourses);

                            _context.NrData1.Update(primaryCatch);

                            // Re-point remaining active centers to primary catch
                            var otherCenters = await _context.CenterList
                                .Where(c => c.ProjectId == ProjectId && secondaryCatchIds.Contains(c.NRDataId) && c.Status)
                                .ToListAsync();

                            foreach (var oc in otherCenters)
                            {
                                if (!deletedCenters.Any(d => d.Id == oc.Id))
                                {
                                    oc.NRDataId = primaryCatchId;
                                    _context.CenterList.Update(oc);
                                }
                            }

                            // Remove secondary catches from NrData1
                            foreach (var sc in secondaryCatches)
                            {
                                _context.NrData1.Remove(sc);
                            }
                        }
                    }

                    mergedCount += cg.Count() - 1;
                }

                // Advance steps to STEP_DUP_PARTIAL for data catches
                foreach (var row in data)
                {
                    if (_context.Entry(row).State == EntityState.Detached || _context.Entry(row).State == EntityState.Deleted)
                        continue;

                    row.Steps = PipelineNavigator.STEP_DUP_PARTIAL;
                    _context.NrData1.Update(row);
                }

                await _context.SaveChangesAsync();

                // Recalculate catch dynamic quantities for affected catches
                foreach (var catchId in affectedNrDataIds)
                {
                    if (nrDataMap.TryGetValue(catchId, out var catchRow) &&
                        _context.Entry(catchRow).State != EntityState.Deleted &&
                        _context.Entry(catchRow).State != EntityState.Detached)
                    {
                        var curCenters = await _context.CenterList
                            .Where(c => c.ProjectId == ProjectId && c.NRDataId == catchId && c.Status && c.ExtraId == 0)
                            .ToListAsync();

                        if (curCenters.Any())
                        {
                            var dyn = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                            if (!string.IsNullOrWhiteSpace(catchRow.NRDatas))
                            {
                                try
                                {
                                    dyn = JsonSerializer.Deserialize<Dictionary<string, string>>(catchRow.NRDatas)
                                        ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                                }
                                catch { }
                            }

                            dyn["Quantity"] = curCenters.Sum(c => c.Quantity).ToString();
                            dyn["NRQuantity"] = curCenters.Sum(c => c.NRQuantity).ToString();
                            catchRow.NRDatas = JsonSerializer.Serialize(dyn);
                            _context.NrData1.Update(catchRow);
                        }
                    }
                }

                await _context.SaveChangesAsync();

                var triggeredBy = LogHelper.GetTriggeredBy(User);

                await _loggerService.LogEventAsync(
                    "Duplicates merged and new rows created (CenterList / NrData1)",
                    "Duplicates",
                    triggeredBy,
                    ProjectId,
                    string.Empty,
                    LogHelper.ToJson(new { ProjectId, MergedCount = mergedCount })
                );

                // ================= REPORT =================
                var reportPath = FileStorageHelper.GetProjectFolder(ProjectId);

                var distinctLots = (lotNo.HasValue && lotNo.Value > 0)
                    ? new List<int> { lotNo.Value }
                    : data.Where(r => r.EnvLotNo > 0).Select(r => r.EnvLotNo).Distinct().OrderBy(l => l).ToList();
                var lotStr = distinctLots.Any() ? string.Join("_", distinctLots) : "All";

                var fileName = ReportVersionHelper.GetNextVersionFileName(reportPath, $"DuplicateTool_NrData1_{lotStr}.xlsx");
                var filePath = Path.Combine(reportPath, fileName);

                // Build report rows: include all centers involved in duplicate groups + new merged centers
                var reportCenters = new List<CenterList>();
                reportCenters.AddRange(deletedCenters);
                reportCenters.AddRange(newCenters);

                // If no duplicates found, show active centers
                if (!reportCenters.Any())
                {
                    reportCenters.AddRange(activeCenters);
                }

                // Collect extra dynamic headers
                var extraHeaders = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var c in reportCenters)
                {
                    if (!string.IsNullOrWhiteSpace(c.CenterData))
                    {
                        try
                        {
                            var dict = JsonSerializer.Deserialize<Dictionary<string, string>>(c.CenterData);
                            if (dict != null)
                            {
                                foreach (var key in dict.Keys)
                                {
                                    if (!key.Equals("Quantity", StringComparison.OrdinalIgnoreCase) &&
                                        !key.Equals("NRQuantity", StringComparison.OrdinalIgnoreCase))
                                    {
                                        extraHeaders.Add(key);
                                    }
                                }
                            }
                        }
                        catch { }
                    }
                }

                var extraHeaderList = extraHeaders.OrderBy(x => x).ToList();

                ExcelPackage.License.SetNonCommercialPersonal("Tools");
                using (var package = new ExcelPackage())
                {
                    // Sheet 1: Merge Report
                    var ws = package.Workbook.Worksheets.Add("Merge Report");

                    var baseHeaders = new List<string>
                    {
                        "CatchNo", "SubjectName", "CourseName", "NodalCode", "CenterCode",
                        "Route", "Quantity", "NRQuantity", "Status"
                    };

                    int col = 1;
                    foreach (var h in baseHeaders)
                    {
                        ws.Cells[1, col].Value = h;
                        ws.Cells[1, col].Style.Font.Bold = true;
                        col++;
                    }

                    foreach (var key in extraHeaderList)
                    {
                        ws.Cells[1, col].Value = key;
                        ws.Cells[1, col].Style.Font.Bold = true;
                        col++;
                    }

                    int rowIdx = 2;
                    foreach (var item in reportCenters)
                    {
                        nrDataMap.TryGetValue(item.NRDataId, out var catchInfo);

                        col = 1;
                        ws.Cells[rowIdx, col++].Value = catchInfo?.CatchNo ?? "";
                        ws.Cells[rowIdx, col++].Value = catchInfo?.SubjectName ?? "";
                        ws.Cells[rowIdx, col++].Value = catchInfo?.CourseName ?? "";
                        ws.Cells[rowIdx, col++].Value = item.NodalCode ?? "";
                        ws.Cells[rowIdx, col++].Value = item.CenterCode ?? "";
                        ws.Cells[rowIdx, col++].Value = item.Route ?? "";
                        ws.Cells[rowIdx, col++].Value = item.Quantity;
                        ws.Cells[rowIdx, col++].Value = item.NRQuantity;
                        ws.Cells[rowIdx, col++].Value = item.Status ? "Active" : "Inactive";

                        Dictionary<string, string>? cDict = null;
                        if (!string.IsNullOrWhiteSpace(item.CenterData))
                        {
                            try
                            {
                                cDict = JsonSerializer.Deserialize<Dictionary<string, string>>(item.CenterData);
                            }
                            catch { }
                        }

                        foreach (var key in extraHeaderList)
                        {
                            string val = "";
                            if (cDict != null && cDict.TryGetValue(key, out var dynVal))
                            {
                                val = dynVal ?? "";
                            }
                            ws.Cells[rowIdx, col++].Value = val;
                        }

                        using (var range = ws.Cells[rowIdx, 1, rowIdx, col - 1])
                        {
                            range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            if (deletedCenters.Any(x => x.Id == item.Id))
                                range.Style.Fill.BackgroundColor.SetColor(Color.LightCoral);
                            else if (newCenters.Any(x => x.Id == item.Id))
                                range.Style.Fill.BackgroundColor.SetColor(Color.LightGreen);
                        }

                        rowIdx++;
                    }

                    if (ws.Dimension != null)
                    {
                        ws.Cells[ws.Dimension.Address].AutoFitColumns();
                        ws.View.FreezePanes(2, 1);
                    }

                    // Sheet 2: Clean Data (Active Centers)
                    var wsClean = package.Workbook.Worksheets.Add("Clean Data");
                    col = 1;
                    foreach (var h in baseHeaders)
                    {
                        wsClean.Cells[1, col].Value = h;
                        wsClean.Cells[1, col].Style.Font.Bold = true;
                        col++;
                    }
                    foreach (var key in extraHeaderList)
                    {
                        wsClean.Cells[1, col].Value = key;
                        wsClean.Cells[1, col].Style.Font.Bold = true;
                        col++;
                    }

                    var cleanCenters = await _context.CenterList
                        .Where(x => x.ProjectId == ProjectId && x.Status && x.ExtraId == 0)
                        .ToListAsync();

                    int cleanRow = 2;
                    foreach (var item in cleanCenters)
                    {
                        nrDataMap.TryGetValue(item.NRDataId, out var catchInfo);

                        col = 1;
                        wsClean.Cells[cleanRow, col++].Value = catchInfo?.CatchNo ?? "";
                        wsClean.Cells[cleanRow, col++].Value = catchInfo?.SubjectName ?? "";
                        wsClean.Cells[cleanRow, col++].Value = catchInfo?.CourseName ?? "";
                        wsClean.Cells[cleanRow, col++].Value = item.NodalCode ?? "";
                        wsClean.Cells[cleanRow, col++].Value = item.CenterCode ?? "";
                        wsClean.Cells[cleanRow, col++].Value = item.Route ?? "";
                        wsClean.Cells[cleanRow, col++].Value = item.Quantity;
                        wsClean.Cells[cleanRow, col++].Value = item.NRQuantity;
                        wsClean.Cells[cleanRow, col++].Value = "Active";

                        Dictionary<string, string>? cDict = null;
                        if (!string.IsNullOrWhiteSpace(item.CenterData))
                        {
                            try
                            {
                                cDict = JsonSerializer.Deserialize<Dictionary<string, string>>(item.CenterData);
                            }
                            catch { }
                        }

                        foreach (var key in extraHeaderList)
                        {
                            string val = "";
                            if (cDict != null && cDict.TryGetValue(key, out var dynVal))
                            {
                                val = dynVal ?? "";
                            }
                            wsClean.Cells[cleanRow, col++].Value = val;
                        }

                        cleanRow++;
                    }

                    if (wsClean.Dimension != null)
                    {
                        wsClean.Cells[wsClean.Dimension.Address].AutoFitColumns();
                        wsClean.View.FreezePanes(2, 1);
                    }

                    package.SaveAs(new FileInfo(filePath));
                }

                // Record in ExcelReports table
                var dupVersionMatch = System.Text.RegularExpressions.Regex.Match(fileName, @"_v(\d+)\.xlsx$", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                int dupVersion = dupVersionMatch.Success ? int.Parse(dupVersionMatch.Groups[1].Value) : 1;
                await ExcelReportHelper.RecordExcelReportAsync(
                    _context,
                    ProjectId,
                    1, // Module 1 (Duplicate Tool)
                    dupVersion,
                    lotNo.HasValue && lotNo.Value > 0 ? lotNo : null,
                    filePath,
                    true,
                    LogHelper.GetTriggeredBy(User)
                );

                return Ok(new
                {
                    MergedRows = mergedCount,
                    message = mergedCount > 0
                        ? "Duplicates processed successfully for NrData1."
                        : "No pending duplicate rows. Report regenerated from active data.",
                    fileName,
                    filePath
                });
            }
            catch (Exception ex)
            {
                await _loggerService.LogErrorAsync("Error solving duplicates for NrData1", ex.ToString(), nameof(NrData1Controller));
                return StatusCode(500, "Internal server error: " + ex.Message);
            }
        }

        [HttpPost("Enhancement")]
        public async Task<IActionResult> ApplyEnhancement(int ProjectId, [FromQuery] int? batch = null, [FromQuery] int? lotNo = null)
        {
            try
            {
                Console.WriteLine($"[NrData1Controller] ApplyEnhancement API called for ProjectId: {ProjectId}, batch: {batch}, lotNo: {lotNo}");

                IQueryable<NrData1> query = _context.NrData1
                    .Where(p => p.ProjectId == ProjectId);

                if (batch.HasValue && batch.Value > 0)
                {
                    query = query.Where(p => p.Batch == batch.Value);
                }
                else
                {
                    query = query.Where(p => p.Batch == 1 || p.Steps == PipelineNavigator.STEP_DUP_PARTIAL);
                }

                if (lotNo.HasValue && lotNo.Value > 0)
                {
                    Console.WriteLine($"[NrData1Controller] Filtering by lotNo: {lotNo.Value}");
                    query = query.Where(p => p.LotNo == lotNo.Value);
                }

                var data = await query.ToListAsync();

                var projectconfig = await _context.ProjectConfigs
                    .Where(p => p.ProjectId == ProjectId).FirstOrDefaultAsync();

                if (projectconfig == null)
                {
                    return NotFound("Project config not exists for this project");
                }

                int smallestInner = 0;
                var innerEnv = projectconfig.Envelope;

                if (!string.IsNullOrEmpty(innerEnv))
                {
                    try
                    {
                        var envelopeDict = JsonSerializer.Deserialize<Dictionary<string, string>>(innerEnv);
                        if (envelopeDict != null && envelopeDict.TryGetValue("Inner", out var innerValue) &&
                            !string.IsNullOrWhiteSpace(innerValue))
                        {
                            var innerSizes = innerValue
                                .Split(',', StringSplitOptions.RemoveEmptyEntries)
                                .Select(e => e.Trim().ToUpper().Replace("E", ""))
                                .Where(x => int.TryParse(x, out _))
                                .Select(int.Parse)
                                .OrderBy(x => x)
                                .ToList();

                            if (innerSizes.Any())
                            {
                                smallestInner = innerSizes.First();
                            }
                        }
                        else if (envelopeDict != null && envelopeDict.TryGetValue("Outer", out var outerValue) &&
                                 !string.IsNullOrWhiteSpace(outerValue))
                        {
                            var outerSizes = outerValue
                                .Split(',', StringSplitOptions.RemoveEmptyEntries)
                                .Select(e => e.Trim().ToUpper().Replace("E", ""))
                                .Where(x => int.TryParse(x, out _))
                                .Select(int.Parse)
                                .OrderBy(x => x)
                                .ToList();

                            if (outerSizes.Any())
                            {
                                smallestInner = outerSizes.First();
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        await _loggerService.LogErrorAsync("Error in serializing Envelope", ex.Message, nameof(NrData1Controller));
                        smallestInner = 0;
                    }
                }

                // Consolidated calculation logic: Round Before (Optional) -> Enhance -> Round After
                if (data.Any())
                {
                    var nrDataIds = data.Select(x => x.Id).ToList();
                    var relatedCenters = await _context.CenterList
                        .Where(c => c.ProjectId == ProjectId && nrDataIds.Contains(c.NRDataId) && c.Status)
                        .ToListAsync();

                    var centersByNrDataId = relatedCenters
                        .GroupBy(c => c.NRDataId)
                        .ToDictionary(g => g.Key, g => g.ToList());

                    foreach (var d in data)
                    {
                        d.Steps = PipelineNavigator.STEP_ENHANCEMENT;

                        if (centersByNrDataId.TryGetValue(d.Id, out var centers))
                        {
                            foreach (var center in centers)
                            {
                                int cNrQty = center.NRQuantity;
                                if (cNrQty <= 0 && !string.IsNullOrWhiteSpace(center.CenterData))
                                {
                                    try
                                    {
                                        var cDict = JsonSerializer.Deserialize<Dictionary<string, string>>(center.CenterData);
                                        if (cDict != null && cDict.TryGetValue("NRQuantity", out var cNrQtyStr) &&
                                            int.TryParse(cNrQtyStr, out int parsedVal))
                                        {
                                            cNrQty = parsedVal;
                                        }
                                    }
                                    catch { }
                                }

                                if (cNrQty > 0)
                                {
                                    double cInitQty = cNrQty;
                                    if (projectconfig.RoundOffBeforeEnhancement && smallestInner > 0)
                                    {
                                        cInitQty = Math.Ceiling(cInitQty / (double)smallestInner) * smallestInner;
                                    }

                                    double cEnhVal = projectconfig.Enhancement > 0
                                        ? (projectconfig.Enhancement * cInitQty) / 100.0
                                        : 0;
                                    double cTotal = cInitQty + cEnhVal;
                                    int cEnhancedQty = smallestInner > 0
                                        ? (int)Math.Ceiling(cTotal / (double)smallestInner) * smallestInner
                                        : (int)Math.Round(cTotal);

                                    center.Quantity = cEnhancedQty;

                                    if (!string.IsNullOrWhiteSpace(center.CenterData))
                                    {
                                        try
                                        {
                                            var cDict = JsonSerializer.Deserialize<Dictionary<string, string>>(center.CenterData) ?? new Dictionary<string, string>();
                                            cDict["Quantity"] = cEnhancedQty.ToString();
                                            center.CenterData = JsonSerializer.Serialize(cDict);
                                        }
                                        catch { }
                                    }

                                    _context.CenterList.Update(center);
                                }
                            }
                        }
                    }

                    await _context.SaveChangesAsync();
                }

                var filePath = string.Empty;

                if (data.Any())
                {
                    var reportPath = FileStorageHelper.GetProjectFolder(ProjectId);

                    var distinctLots = (lotNo.HasValue && lotNo.Value > 0)
                        ? new List<int> { lotNo.Value }
                        : data.Where(r => r.EnvLotNo > 0).Select(r => r.EnvLotNo).Distinct().OrderBy(l => l).ToList();
                    var lotStr = distinctLots.Any() ? string.Join("_", distinctLots) : "All";

                    var fileName = ReportVersionHelper.GetNextVersionFileName(reportPath, $"EnhancementReport_NrData1_{lotStr}.xlsx");
                    filePath = Path.Combine(reportPath, fileName);

                    var baseProperties = typeof(NrData1).GetProperties(BindingFlags.Public | BindingFlags.Instance)
                        .Where(p => p.Name != "NRDatas" && p.Name != "ProjectId")
                        .ToList();

                    ExcelPackage.License.SetNonCommercialPersonal("Tools");
                    using (var package = new ExcelPackage())
                    {
                        var ws = package.Workbook.Worksheets.Add("Enhancement Report");

                        int col = 1;
                        foreach (var prop in baseProperties)
                        {
                            ws.Cells[1, col].Value = prop.Name;
                            ws.Cells[1, col].Style.Font.Bold = true;
                            col++;
                        }

                        int row = 2;
                        foreach (var item in data)
                        {
                            col = 1;
                            foreach (var prop in baseProperties)
                            {
                                ws.Cells[row, col++].Value = prop.GetValue(item)?.ToString();
                            }
                            row++;
                        }

                        ws.Cells[ws.Dimension.Address].AutoFitColumns();
                        ws.View.FreezePanes(2, 1);

                        package.SaveAs(new FileInfo(filePath));

                        // Record in ExcelReports table
                        var enhVersionMatch = System.Text.RegularExpressions.Regex.Match(fileName, @"_v(\d+)\.xlsx$", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                        int enhVersion = enhVersionMatch.Success ? int.Parse(enhVersionMatch.Groups[1].Value) : 1;
                        await ExcelReportHelper.RecordExcelReportAsync(
                            _context,
                            ProjectId,
                            2, // Module 2 (Envelope Setup and Enhancement)
                            enhVersion,
                            lotNo.HasValue && lotNo.Value > 0 ? lotNo : null,
                            filePath,
                            true,
                            LogHelper.GetTriggeredBy(User)
                        );
                    }

                    // Logging
                    var triggeredBy = LogHelper.GetTriggeredBy(User);

                    await _loggerService.LogEventAsync(
                        "Enhancement applied with report (NrData1)",
                        "Enhancement",
                        triggeredBy,
                        ProjectId,
                        string.Empty,
                        LogHelper.ToJson(new { ProjectId })
                    );
                }

                try
                {
                    Console.WriteLine("[NrData1Controller] Envelope breaking is called");
                    var envelopeController = new EnvelopeBreakagesController(_context, _loggerService, _apiSettings, _dispatchService);
                    await envelopeController.EnvelopeConfiguration(ProjectId, bypassDispatch: true);
                }
                catch (Exception ex)
                {
                    await _loggerService.LogErrorAsync("Error running combined EnvelopeConfiguration", ex.Message, nameof(NrData1Controller));
                }

                return Ok(new
                {
                    EnhancementApplied = data.Any() ? projectconfig.Enhancement : 0.0,
                    ReportPath = filePath,
                    fileName = !string.IsNullOrEmpty(filePath) ? Path.GetFileName(filePath) : null
                });
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[NrData1Controller ApplyEnhancement] Exception thrown: {ex}");
                await _loggerService.LogErrorAsync("Error applying enhancement for NrData1", ex.ToString(), nameof(NrData1Controller));
                return StatusCode(500, "Internal server error: " + ex.Message);
            }
        }

        [HttpPost("EnvelopeConfiguration")]
        public async Task<IActionResult> EnvelopeConfiguration(int ProjectId, [FromQuery] bool bypassDispatch = false)
        {
            try
            {
                var projectConfig = await _context.ProjectConfigs
                    .FirstOrDefaultAsync(s => s.ProjectId == ProjectId);

                if (projectConfig == null)
                    return NotFound("Project config not found for this project.");

                var envelopesJson = projectConfig.Envelope;
                if (string.IsNullOrWhiteSpace(envelopesJson))
                    return BadRequest("Envelope configuration not found in ProjectConfig.");

                var envelopeDict = JsonSerializer.Deserialize<Dictionary<string, string>>(envelopesJson);
                if (envelopeDict == null)
                    return BadRequest("Invalid envelope JSON format.");

                var innerSizes = envelopeDict["Inner"]
                    .Split(',', StringSplitOptions.RemoveEmptyEntries)
                    .Select(e => e.Trim().ToUpper().Replace("E", ""))
                    .Where(x => int.TryParse(x, out _))
                    .Select(int.Parse)
                    .OrderByDescending(x => x)
                    .ToList();

                var outerSizes = envelopeDict["Outer"]
                    .Split(',', StringSplitOptions.RemoveEmptyEntries)
                    .Select(e => e.Trim().ToUpper().Replace("E", ""))
                    .Where(x => int.TryParse(x, out _))
                    .Select(int.Parse)
                    .OrderByDescending(x => x)
                    .ToList();

                // 1. Fetch enhanced NrData1 records for Project
                var nrDataList = await _context.NrData1
                    .Where(s => s.ProjectId == ProjectId && s.Steps == PipelineNavigator.STEP_ENHANCEMENT)
                    .ToListAsync();

                if (!nrDataList.Any())
                {
                    return Ok("Envelope breakdown report has been successfully created (no pending enhanced records).");
                }

                var nrDataIds = nrDataList.Select(r => r.Id).ToList();

                // 2. Fetch active centers for these NrData1 records
                var centerList = await _context.CenterList
                    .Where(c => c.ProjectId == ProjectId && nrDataIds.Contains(c.NRDataId) && c.Status)
                    .ToListAsync();

                bool hasExtraConfig = await _context.ExtraConfigurations.AnyAsync(e => e.ProjectId == ProjectId);

                // 3. Deactivate existing active NewEnvelopeBreakage entries
                var existingBreakages = await _context.NewEnvelopeBreakages
                    .Where(p => p.ProjectId == ProjectId && nrDataIds.Contains(p.NrDataId) && p.Status == true)
                    .ToListAsync();

                if (existingBreakages.Any())
                {
                    foreach (var item in existingBreakages)
                    {
                        item.Status = false;
                    }
                    await _context.SaveChangesAsync();

                    await _loggerService.LogEventAsync(
                        $"Successfully deactivated {existingBreakages.Count} NewEnvelopeBreakage entries (Status=0) for ProjectID {ProjectId}",
                        "NewEnvelopeBreakages",
                        LogHelper.GetTriggeredBy(User),
                        ProjectId
                    );
                }

                var breakagesToAdd = new List<NewEnvelopeBreakage>();
                var legacyBreakagesToAdd = new List<EnvelopeBreakage>();

                // 4. Calculate envelope breaking for EACH center
                foreach (var center in centerList)
                {
                    int quantity = center.Quantity > 0 ? center.Quantity : center.NRQuantity;
                    if (quantity <= 0 && !string.IsNullOrWhiteSpace(center.CenterData))
                    {
                        try
                        {
                            var cDict = JsonSerializer.Deserialize<Dictionary<string, string>>(center.CenterData);
                            if (cDict != null && cDict.TryGetValue("Quantity", out var qStr) && int.TryParse(qStr, out int qVal))
                            {
                                quantity = qVal;
                            }
                            else if (cDict != null && cDict.TryGetValue("NRQuantity", out var nrqStr) && int.TryParse(nrqStr, out int nrqVal))
                            {
                                quantity = nrqVal;
                            }
                        }
                        catch { }
                    }

                    if (quantity <= 0) continue;

                    // Inner breakdown
                    int remaining = quantity;
                    var innerBreakdown = new Dictionary<string, string>();

                    if (innerSizes.Any())
                    {
                        foreach (var size in innerSizes)
                        {
                            int count = remaining / size;
                            if (count > 0)
                            {
                                innerBreakdown[$"E{size}"] = count.ToString();
                                remaining -= count * size;
                            }
                        }

                        if (remaining > 0)
                        {
                            int smallestInner = innerSizes.Last();
                            int count = (int)Math.Ceiling((double)remaining / smallestInner);
                            string key = $"E{smallestInner}";
                            innerBreakdown[key] = (innerBreakdown.ContainsKey(key)
                                ? int.Parse(innerBreakdown[key]) + count
                                : count).ToString();
                        }
                    }

                    // Outer breakdown
                    remaining = quantity;
                    var outerBreakdown = new Dictionary<string, string>();
                    int totalOuterCount = 0;

                    if (outerSizes.Any())
                    {
                        foreach (var size in outerSizes)
                        {
                            int count = remaining / size;
                            if (count > 0)
                            {
                                outerBreakdown[$"E{size}"] = count.ToString();
                                totalOuterCount += count;
                                remaining -= count * size;
                            }
                        }

                        if (remaining > 0)
                        {
                            int smallestSize = outerSizes.Last();
                            int count = (int)Math.Ceiling((double)remaining / smallestSize);

                            string key = $"E{smallestSize}";
                            if (outerBreakdown.ContainsKey(key))
                            {
                                outerBreakdown[key] = (int.Parse(outerBreakdown[key]) + count).ToString();
                            }
                            else
                            {
                                outerBreakdown[key] = count.ToString();
                            }

                            totalOuterCount += count;
                            remaining = 0;
                        }
                    }

                    if (!innerBreakdown.Any() && outerBreakdown.Any())
                    {
                        innerBreakdown = new Dictionary<string, string>(outerBreakdown);
                    }

                    var newBreakage = new NewEnvelopeBreakage
                    {
                        ProjectId = ProjectId,
                        NrDataId = center.NRDataId,
                        CenterListId = center.Id,
                        InnerEnvelope = JsonSerializer.Serialize(innerBreakdown),
                        OuterEnvelope = JsonSerializer.Serialize(outerBreakdown),
                        TotalEnvelope = totalOuterCount,
                        Status = true
                    };
                    breakagesToAdd.Add(newBreakage);

                    var legacyBreakage = new EnvelopeBreakage
                    {
                        ProjectId = ProjectId,
                        NrDataId = center.NRDataId,
                        InnerEnvelope = JsonSerializer.Serialize(innerBreakdown),
                        OuterEnvelope = JsonSerializer.Serialize(outerBreakdown),
                        TotalEnvelope = totalOuterCount,
                        Status = true
                    };
                    legacyBreakagesToAdd.Add(legacyBreakage);
                }

                if (breakagesToAdd.Any())
                {
                    await _context.NewEnvelopeBreakages.AddRangeAsync(breakagesToAdd);
                    await _context.EnvelopeBreakages.AddRangeAsync(legacyBreakagesToAdd);

                    // Update steps for NrData1
                    foreach (var nr in nrDataList)
                    {
                        if (hasExtraConfig)
                        {
                            nr.Steps = PipelineNavigator.STEP_ENV_BREAKING;
                        }
                        else
                        {
                            nr.Steps = PipelineNavigator.STEP_AWAITING_EXTRA;
                        }
                    }

                    await _context.SaveChangesAsync();
                    await _loggerService.LogEventAsync(
                        $"Created {breakagesToAdd.Count} NewEnvelopeBreakage entries for ProjectID {ProjectId} and updated steps",
                        "NewEnvelopeBreakages",
                        LogHelper.GetTriggeredBy(User),
                        ProjectId
                    );
                }

                return Ok("Envelope breakdown report has been successfully created.");
            }
            catch (Exception ex)
            {
                await _loggerService.LogErrorAsync("Error creating NewEnvelopeBreakage", ex.ToString(), nameof(NrData1Controller));
                return StatusCode(500, "Internal Server Error: " + ex.Message);
            }
        }

        [HttpPost("PostExtraEnvelopes")]
        public async Task<ActionResult> PostExtraEnvelopes(int ProjectId, int? uploadId = null, [FromQuery] int? batchNo = null, [FromQuery] int? lotNo = null)
        {
            try
            {
                var projectConfig = await _context.ProjectConfigs
                    .FirstOrDefaultAsync(p => p.ProjectId == ProjectId);

                var eligibleSteps = PipelineNavigator.GetEligiblePickupSteps(PipelineNavigator.STEP_AWAITING_EXTRA);

                var query = _context.NrData1
                    .Where(d => d.ProjectId == ProjectId && eligibleSteps.Contains(d.Steps) && d.Batch == (batchNo ?? 1));

                if (lotNo.HasValue && lotNo.Value > 0)
                {
                    query = query.Where(d => d.LotNo == lotNo.Value);
                }

                var nrDataList = await query.ToListAsync();

                if (!nrDataList.Any())
                {
                    // Fallback to all records in this batch
                    query = _context.NrData1.Where(d => d.ProjectId == ProjectId && d.Batch == (batchNo ?? 1));
                    if (lotNo.HasValue && lotNo.Value > 0)
                        query = query.Where(d => d.LotNo == lotNo.Value);
                    nrDataList = await query.ToListAsync();
                }

                if (!nrDataList.Any())
                    return BadRequest("No suitable data found for Extra Configuration. Either no data is at the required step or all data is already processed.");

                var extraConfig = await _context.ExtraConfigurations
                    .Where(c => c.ProjectId == ProjectId)
                    .ToListAsync();

                var allExtraTypes = await _context.ExtraType
                    .OrderBy(e => e.Order)
                    .ToListAsync();
                var extraTypesMap = allExtraTypes.ToDictionary(e => e.ExtraTypeId, e => e.Type);
                var extraTypeOrderMap = allExtraTypes.ToDictionary(e => e.ExtraTypeId, e => e.Order);

                // Sort extra configurations according to ExtraType table Order
                extraConfig = extraConfig
                    .OrderBy(c => extraTypeOrderMap.TryGetValue(c.ExtraType, out var ord) ? ord : c.ExtraType)
                    .ToList();

                var nrDataIds = nrDataList.Select(d => d.Id).ToList();
                var existingCenters = await _context.CenterList
                    .Where(c => c.ProjectId == ProjectId && nrDataIds.Contains(c.NRDataId) && c.ExtraId == 0 && c.Status)
                    .ToListAsync();

                bool isOnlyReport = extraConfig.Any(c => c.IsExtraProcessingAsPerNR);
                var envelopesToUse = new List<ExtraEnvelopeReportDto>();

                if (isOnlyReport)
                {
                    var extraCenters = await _context.CenterList
                        .Where(c => c.ProjectId == ProjectId && c.ExtraId > 0 && c.Status)
                        .ToListAsync();

                    if (!extraCenters.Any())
                        return BadRequest("No existing Extra Envelope data found in centerlist for report.");

                    var extraCenterIds = extraCenters.Select(c => c.Id).ToList();
                    var breakages = await _context.NewEnvelopeBreakages
                        .Where(b => b.ProjectId == ProjectId && extraCenterIds.Contains(b.CenterListId) && b.Status)
                        .ToListAsync();

                    var nrLookup = nrDataList.ToDictionary(n => n.Id, n => n.CatchNo);

                    foreach (var c in extraCenters)
                    {
                        var b = breakages.FirstOrDefault(x => x.CenterListId == c.Id);
                        nrLookup.TryGetValue(c.NRDataId, out var catchNo);

                        envelopesToUse.Add(new ExtraEnvelopeReportDto
                        {
                            CatchNo = catchNo,
                            NodalCode = c.NodalCode,
                            ExtraId = c.ExtraId,
                            Quantity = c.Quantity > 0 ? c.Quantity : c.NRQuantity,
                            InnerEnvelope = b?.InnerEnvelope,
                            OuterEnvelope = b?.OuterEnvelope
                        });
                    }
                }
                else
                {
                    if (!extraConfig.Any())
                        return BadRequest("No ExtraConfiguration found");

                    var envelopesToAdd = new List<ExtraEnvelopeReportDto>();

                    // Pre-fetch all active extra centers and their breakages once to avoid DB calls in loop
                    var existingExtraCenters = await _context.CenterList
                        .Where(c => c.ProjectId == ProjectId && c.ExtraId > 0 && c.Status)
                        .ToListAsync();

                    var existingExtraCenterIds = existingExtraCenters.Select(c => c.Id).ToList();
                    var existingNewBreakages = await _context.NewEnvelopeBreakages
                        .Where(b => b.ProjectId == ProjectId && existingExtraCenterIds.Contains(b.CenterListId) && b.Status)
                        .ToListAsync();

                    var allNodals = await _context.CenterList
                        .Where(n => n.ProjectId == ProjectId && n.Status && !string.IsNullOrEmpty(n.NodalCode))
                        .Select(n => n.NodalCode!.Trim())
                        .Distinct()
                        .ToListAsync();

                    var newCentersToInsert = new List<CenterList>();
                    var pendingBreakages = new List<(CenterList CenterRecord, int NrDataId, Dictionary<string, string> InnerBreakdown, Dictionary<string, string> OuterBreakdown, int TotalEnvelope)>();

                    foreach (var config in extraConfig)
                    {
                        var extraTypeEntity = allExtraTypes.FirstOrDefault(e => e.ExtraTypeId == config.ExtraType);
                        string extraTypeName = extraTypeEntity?.Type ?? (extraTypesMap.TryGetValue(config.ExtraType, out var t) ? t : "");
                        bool isNodalExtra = (!string.IsNullOrWhiteSpace(extraTypeName) && extraTypeName.Contains("Nodal", StringComparison.OrdinalIgnoreCase))
                            || (extraTypeEntity != null && extraTypeEntity.Order == 1);
                        bool useNodal = isNodalExtra || !string.IsNullOrWhiteSpace(config.nodalValue);

                        EnvelopeTypeConfig? envelopeType = null;
                        if (!string.IsNullOrWhiteSpace(config.EnvelopeType))
                        {
                            try { envelopeType = JsonSerializer.Deserialize<EnvelopeTypeConfig>(config.EnvelopeType); }
                            catch { }
                        }

                        int innerCapacity = envelopeType != null && !string.IsNullOrWhiteSpace(envelopeType.Inner) ? GetEnvelopeCapacity(envelopeType.Inner) : 1;
                        int outerCapacity = envelopeType != null && !string.IsNullOrWhiteSpace(envelopeType.Outer) ? GetEnvelopeCapacity(envelopeType.Outer) : 1;

                        // Group data
                        var groupedData = useNodal
                            ? nrDataList
                                .SelectMany(nr =>
                                {
                                    var centersForNr = existingCenters.Where(c => c.NRDataId == nr.Id && (!isNodalExtra || !string.IsNullOrWhiteSpace(c.NodalCode))).ToList();
                                    if (centersForNr.Any())
                                    {
                                        return centersForNr.Select(c => new
                                        {
                                            nr.CatchNo,
                                            NodalCode = (c.NodalCode ?? "").Trim(),
                                            Quantity = c.Quantity > 0 ? c.Quantity : c.NRQuantity,
                                            NrDataId = nr.Id
                                        });
                                    }
                                    return new[]
                                    {
                                        new
                                        {
                                            nr.CatchNo,
                                            NodalCode = "",
                                            Quantity = 0,
                                            NrDataId = nr.Id
                                        }
                                    };
                                })
                                .GroupBy(x => new { x.CatchNo, x.NodalCode })
                                .Select(g => new
                                {
                                    g.Key.CatchNo,
                                    g.Key.NodalCode,
                                    Quantity = g.Sum(x => x.Quantity),
                                    NrDataId = g.First().NrDataId
                                }).ToList()
                            : nrDataList
                                .GroupBy(x => x.CatchNo)
                                .Select(g =>
                                {
                                    var catchNrIds = g.Select(x => x.Id).ToList();
                                    var centersForCatch = existingCenters.Where(c => catchNrIds.Contains(c.NRDataId)).ToList();
                                    int totalCatchQty = centersForCatch.Sum(c => c.Quantity > 0 ? c.Quantity : c.NRQuantity);
                                    return new
                                    {
                                        CatchNo = g.Key,
                                        NodalCode = (string?)null,
                                        Quantity = totalCatchQty,
                                        NrDataId = g.First().Id
                                    };
                                }).ToList();

                        if (isNodalExtra && config.AttachExtraForEachCatchForAllNodal)
                        {
                            var distinctCatches = nrDataList.Select(x => x.CatchNo).Where(c => !string.IsNullOrEmpty(c)).Distinct().ToList();

                            var existingPairs = new HashSet<(string?, string)>(groupedData.Select(g => (g.CatchNo, g.NodalCode ?? "")));
                            foreach (var cNo in distinctCatches)
                            {
                                var matchingNr = nrDataList.FirstOrDefault(x => x.CatchNo == cNo);
                                foreach (var nCode in allNodals)
                                {
                                    if (!existingPairs.Contains((cNo, nCode)))
                                    {
                                        groupedData.Add(new
                                        {
                                            CatchNo = cNo,
                                            NodalCode = (string?)nCode,
                                            Quantity = 0,
                                            NrDataId = matchingNr?.Id ?? 0
                                        });
                                        existingPairs.Add((cNo, nCode));
                                    }
                                }
                            }
                            groupedData.RemoveAll(g => string.IsNullOrWhiteSpace(g.NodalCode));
                        }
                        else if (isNodalExtra)
                        {
                            groupedData.RemoveAll(g => string.IsNullOrWhiteSpace(g.NodalCode));
                        }

                        // Deduplicate within run so only 1 record per (NrDataId, NodalCode) is processed
                        groupedData = groupedData
                            .GroupBy(x => new { x.NrDataId, Nodal = (x.NodalCode ?? "").Trim() })
                            .Select(g => new
                            {
                                CatchNo = g.First().CatchNo,
                                NodalCode = g.First().NodalCode,
                                Quantity = g.Sum(x => x.Quantity),
                                NrDataId = g.Key.NrDataId
                            })
                            .ToList();

                        // Parse nodal configs if any
                        List<NodalValueConfig>? nodalConfigs = null;
                        if (useNodal && !string.IsNullOrWhiteSpace(config.nodalValue))
                        {
                            try
                            {
                                nodalConfigs = JsonSerializer.Deserialize<List<NodalValueConfig>>(
                                    config.nodalValue,
                                    new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
                            }
                            catch { }
                        }

                        foreach (var data in groupedData)
                        {
                            if (isNodalExtra && string.IsNullOrWhiteSpace(data.NodalCode))
                                continue;

                            int calculatedQuantity = 0;

                            if (useNodal && nodalConfigs != null)
                            {
                                var match = nodalConfigs.FirstOrDefault(nc =>
                                    !string.IsNullOrEmpty(nc.NodalCodes) &&
                                    nc.NodalCodes.Split(',')
                                        .Select(n => n.Trim())
                                        .Contains(data.NodalCode, StringComparer.OrdinalIgnoreCase));

                                if (match != null)
                                {
                                    switch (config.Mode)
                                    {
                                        case "Fixed":
                                            if (int.TryParse(match.Value, out var fv))
                                                calculatedQuantity = fv;
                                            break;

                                        case "Percentage":
                                            if (decimal.TryParse(match.Value, out var pv))
                                                calculatedQuantity = (int)Math.Round((double)(data.Quantity * pv) / 100);
                                            break;
                                    }
                                }
                                else
                                {
                                    switch (config.Mode)
                                    {
                                        case "Fixed":
                                            if (int.TryParse(config.Value, out var fqv))
                                                calculatedQuantity = fqv;
                                            break;

                                        case "Percentage":
                                            if (decimal.TryParse(config.Value, out var percent))
                                                calculatedQuantity = (int)Math.Round((double)(data.Quantity * percent) / 100);
                                            break;

                                        case "Range":
                                            if (!string.IsNullOrEmpty(config.RangeConfig))
                                            {
                                                var rangeConfig = JsonSerializer.Deserialize<RangeConfigModel>(config.RangeConfig);
                                                var range = rangeConfig?.ranges?
                                                    .FirstOrDefault(r => data.Quantity >= r.from && data.Quantity <= r.to);
                                                if (range != null)
                                                    calculatedQuantity = range.value;
                                            }
                                            break;
                                    }
                                }
                            }
                            else
                            {
                                switch (config.Mode)
                                {
                                    case "Fixed":
                                        if (int.TryParse(config.Value, out var fqv))
                                            calculatedQuantity = fqv;
                                        break;

                                    case "Percentage":
                                        if (decimal.TryParse(config.Value, out var percent))
                                            calculatedQuantity = (int)Math.Round((double)(data.Quantity * percent) / 100);
                                        break;

                                    case "Range":
                                        if (!string.IsNullOrEmpty(config.RangeConfig))
                                        {
                                            var rangeConfig = JsonSerializer.Deserialize<RangeConfigModel>(config.RangeConfig);
                                            var range = rangeConfig?.ranges?
                                                .FirstOrDefault(r => data.Quantity >= r.from && data.Quantity <= r.to);
                                            if (range != null)
                                                calculatedQuantity = range.value;
                                        }
                                        break;
                                }
                            }

                            int innerCount = innerCapacity > 0
                                ? (int)Math.Ceiling((double)calculatedQuantity / innerCapacity)
                                : 0;

                            int outerCount = outerCapacity > 0
                                ? (int)Math.Ceiling((double)calculatedQuantity / outerCapacity)
                                : 0;

                            var extraReportItem = new ExtraEnvelopeReportDto
                            {
                                CatchNo = data.CatchNo,
                                NodalCode = data.NodalCode,
                                ExtraId = config.ExtraType,
                                Quantity = calculatedQuantity,
                                InnerEnvelope = innerCount.ToString(),
                                OuterEnvelope = outerCount.ToString()
                            };
                            envelopesToAdd.Add(extraReportItem);

                            // 1. Post entry in CenterList (centerlist table) dynamically from ExtraType table
                            string extraCenterName = !string.IsNullOrWhiteSpace(extraTypeName)
                                ? extraTypeName
                                : $"Extra {config.ExtraType}";

                            var centerExtraDict = new Dictionary<string, string>
                            {
                                { "Quantity", calculatedQuantity.ToString() },
                                { "NRQuantity", calculatedQuantity.ToString() },
                                { "InnerEnvelope", innerCount.ToString() },
                                { "OuterEnvelope", outerCount.ToString() }
                            };

                            var matchingCenterForSave = existingCenters.FirstOrDefault(c => c.NRDataId == data.NrDataId && (!string.IsNullOrWhiteSpace(data.NodalCode) && c.NodalCode == data.NodalCode))
                                ?? existingCenters.FirstOrDefault(c => c.NRDataId == data.NrDataId);

                            double extraNodalSort = matchingCenterForSave?.NodalSort ?? 0;
                            double extraCenterSort = 0;
                            int extraRouteSort = matchingCenterForSave?.RouteSort ?? 0;

                            switch (config.ExtraType)
                            {
                                case 1:
                                    extraCenterSort = 10000;
                                    break;
                                case 2:
                                    extraNodalSort = 10000;
                                    extraCenterSort = 100000;
                                    extraRouteSort = 10000;
                                    break;
                                case 3:
                                    extraNodalSort = 100000;
                                    extraCenterSort = 1000000;
                                    extraRouteSort = 100000;
                                    break;
                            }

                            // Ensure only 1 active record can exist for combination (ExtraId + NodalCode + NRDataId)
                            string targetNodal = (data.NodalCode ?? "").Trim();
                            var oldMatchingCenters = existingExtraCenters
                                .Where(c => c.ExtraId == config.ExtraType &&
                                            c.NRDataId == data.NrDataId &&
                                            c.Status &&
                                            ((c.NodalCode == null && string.IsNullOrEmpty(targetNodal)) ||
                                             (c.NodalCode ?? "").Trim().Equals(targetNodal, StringComparison.OrdinalIgnoreCase)))
                                .ToList();

                            foreach (var oldCenter in oldMatchingCenters)
                            {
                                oldCenter.Status = false;
                                _context.CenterList.Update(oldCenter);

                                var oldNewBreakages = existingNewBreakages
                                    .Where(b => b.CenterListId == oldCenter.Id && b.Status)
                                    .ToList();
                                foreach (var ob in oldNewBreakages)
                                {
                                    ob.Status = false;
                                    _context.NewEnvelopeBreakages.Update(ob);
                                }
                            }

                            var extraCenterRecord = new CenterList
                            {
                                ProjectId = ProjectId,
                                NRDataId = data.NrDataId,
                                CenterCode = extraCenterName,
                                NodalCode = data.NodalCode,
                                ExtraId = config.ExtraType,
                                Quantity = calculatedQuantity,
                                NRQuantity = calculatedQuantity,
                                CenterSort = extraCenterSort,
                                NodalSort = extraNodalSort,
                                RouteSort = extraRouteSort,
                                Status = true,
                                CenterData = JsonSerializer.Serialize(centerExtraDict)
                            };
                            newCentersToInsert.Add(extraCenterRecord);

                            // 2. Prepare envelope breaking for batch insert
                            var innerBreakdown = new Dictionary<string, string>();
                            if (innerCapacity > 0 && innerCount > 0)
                            {
                                innerBreakdown[$"E{innerCapacity}"] = innerCount.ToString();
                            }

                            var outerBreakdown = new Dictionary<string, string>();
                            if (outerCapacity > 0 && outerCount > 0)
                            {
                                outerBreakdown[$"E{outerCapacity}"] = outerCount.ToString();
                            }

                            pendingBreakages.Add((extraCenterRecord, data.NrDataId, innerBreakdown, outerBreakdown, outerCount));
                        }
                    }

                    // 1. Batch insert new centers and commit deactivations in 1 roundtrip
                    if (newCentersToInsert.Any())
                    {
                        await _context.CenterList.AddRangeAsync(newCentersToInsert);
                    }
                    await _context.SaveChangesAsync(); // Generated IDs now populated on extraCenterRecord.Id

                    // 2. Batch insert NewEnvelopeBreakage & EnvelopeBreakage
                    var newEnvBreakages = new List<NewEnvelopeBreakage>();
                    var envBreakages = new List<EnvelopeBreakage>();

                    foreach (var pb in pendingBreakages)
                    {
                        newEnvBreakages.Add(new NewEnvelopeBreakage
                        {
                            ProjectId = ProjectId,
                            NrDataId = pb.NrDataId,
                            CenterListId = pb.CenterRecord.Id,
                            InnerEnvelope = JsonSerializer.Serialize(pb.InnerBreakdown),
                            OuterEnvelope = JsonSerializer.Serialize(pb.OuterBreakdown),
                            TotalEnvelope = pb.TotalEnvelope,
                            Status = true
                        });

                        envBreakages.Add(new EnvelopeBreakage
                        {
                            ProjectId = ProjectId,
                            NrDataId = pb.NrDataId,
                            InnerEnvelope = JsonSerializer.Serialize(pb.InnerBreakdown),
                            OuterEnvelope = JsonSerializer.Serialize(pb.OuterBreakdown),
                            TotalEnvelope = pb.TotalEnvelope,
                            Status = true
                        });
                    }

                    if (newEnvBreakages.Any())
                    {
                        await _context.NewEnvelopeBreakages.AddRangeAsync(newEnvBreakages);
                    }
                    if (envBreakages.Any())
                    {
                        await _context.EnvelopeBreakages.AddRangeAsync(envBreakages);
                    }

                    envelopesToUse = envelopesToAdd;
                }

                // Update steps for NrData1
                foreach (var nr in nrDataList)
                {
                    nr.Steps = PipelineNavigator.GetNextStep(PipelineNavigator.STEP_ENV_BREAKING, projectConfig?.Modules);
                    _context.NrData1.Update(nr);
                }

                await _context.SaveChangesAsync();

                // Generate Excel Report
                var extraConfigs = await _context.ExtraConfigurations
                    .Where(x => x.ProjectId == ProjectId)
                    .ToListAsync();

                var extraHeaders = new HashSet<string>();
                var innerKeys = new HashSet<string>();
                var outerKeys = new HashSet<string>();

                var allRows = new List<Dictionary<string, object>>();

                foreach (var extra in envelopesToUse)
                {
                    var baseRow = nrDataList.FirstOrDefault(x =>
                        x.CatchNo == extra.CatchNo);

                    var config = extraConfigs.FirstOrDefault(c => c.ExtraType == extra.ExtraId);
                    if (baseRow == null || config == null) continue;

                    var matchingCenter = existingCenters.FirstOrDefault(c => c.NRDataId == baseRow.Id && (!string.IsNullOrWhiteSpace(extra.NodalCode) && c.NodalCode == extra.NodalCode))
                        ?? existingCenters.FirstOrDefault(c => c.NRDataId == baseRow.Id);

                    double baseNodalSort = matchingCenter?.NodalSort ?? 0;
                    int baseRouteSort = matchingCenter?.RouteSort ?? 0;

                    if (baseNodalSort == 0 && !string.IsNullOrWhiteSpace(baseRow.NRDatas))
                    {
                        try
                        {
                            var dyn = JsonSerializer.Deserialize<Dictionary<string, string>>(baseRow.NRDatas);
                            if (dyn != null)
                            {
                                if (dyn.TryGetValue("NodalSort", out var nsStr) && double.TryParse(nsStr, out var nsVal))
                                    baseNodalSort = nsVal;
                                if (baseRouteSort == 0 && dyn.TryGetValue("RouteSort", out var rsStr) && int.TryParse(rsStr, out var rsVal))
                                    baseRouteSort = rsVal;
                            }
                        }
                        catch { }
                    }

                    var dict = new Dictionary<string, object>();
                    dict["CatchNo"] = baseRow.CatchNo ?? "";
                    dict["CourseName"] = baseRow.CourseName ?? "";
                    dict["SubjectName"] = baseRow.SubjectName ?? "";
                    dict["Quantity"] = extra.Quantity;
                    dict["NodalCode"] = extra.NodalCode ?? "";
                    dict["ExtraId"] = extra.ExtraId;
                    dict["InnerEnvelope"] = extra.InnerEnvelope ?? "";
                    dict["OuterEnvelope"] = extra.OuterEnvelope ?? "";

                    // Set CenterCode
                    switch (extra.ExtraId)
                    {
                        case 1:
                            dict["CenterCode"] = "Nodal Extra";
                            dict["NodalSort"] = baseNodalSort;
                            dict["CenterSort"] = 10000;
                            dict["RouteSort"] = baseRouteSort;
                            break;
                        case 2:
                            dict["CenterCode"] = "University Extra";
                            dict["NodalSort"] = 10000;
                            dict["CenterSort"] = 100000;
                            dict["RouteSort"] = 10000;
                            break;
                        case 3:
                            dict["CenterCode"] = "Office Extra";
                            dict["NodalSort"] = 100000;
                            dict["CenterSort"] = 1000000;
                            dict["RouteSort"] = 100000;
                            break;
                        default:
                            dict["CenterCode"] = "Extra";
                            break;
                    }

                    allRows.Add(dict);
                }

                var path = FileStorageHelper.GetProjectFolder(ProjectId);
                var distinctLots = (lotNo.HasValue && lotNo.Value > 0)
                    ? new List<int> { lotNo.Value }
                    : nrDataList.Where(r => r.EnvLotNo > 0).Select(r => r.EnvLotNo).Distinct().OrderBy(l => l).ToList();
                var lotStr = distinctLots.Any() ? string.Join("_", distinctLots) : "All";

                var fileName = uploadId.HasValue
                    ? $"ExtrasCalculation_NrData1_{lotStr}_v{uploadId}.xlsx"
                    : ReportVersionHelper.GetNextVersionFileName(path, $"ExtrasCalculation_NrData1_{lotStr}.xlsx");
                var filePath = Path.Combine(path, fileName);

                ExcelPackage.License.SetNonCommercialPersonal("Tools");
                using (var package = new ExcelPackage())
                {
                    var ws = package.Workbook.Worksheets.Add("Extra Envelope");
                    var headers = new List<string> { "CatchNo", "CourseName", "SubjectName", "CenterCode", "Quantity", "NodalCode", "ExtraId", "NodalSort", "CenterSort", "RouteSort", "InnerEnvelope", "OuterEnvelope" };
                    foreach (var rowDict in allRows)
                    {
                        foreach (var k in rowDict.Keys)
                        {
                            if (!headers.Contains(k))
                                headers.Add(k);
                        }
                    }

                    for (int i = 0; i < headers.Count; i++)
                        ws.Cells[1, i + 1].Value = headers[i];

                    int row = 2;
                    foreach (var rowDict in allRows)
                    {
                        for (int col = 0; col < headers.Count; col++)
                        {
                            rowDict.TryGetValue(headers[col], out object? val);
                            ws.Cells[row, col + 1].Value = val?.ToString();
                        }
                        row++;
                    }

                    if (ws.Dimension != null)
                        ws.Cells[ws.Dimension.Address].AutoFitColumns();
                    package.SaveAs(new FileInfo(filePath));
                }

                await ExcelReportHelper.RecordExcelReportAsync(
                    _context,
                    ProjectId,
                    3, // Module 3 (Extras Calculation)
                    uploadId ?? 1,
                    lotNo,
                    filePath,
                    true,
                    LogHelper.GetTriggeredBy(User)
                );

                return Ok(new
                {
                    message = isOnlyReport
                        ? "Report generated from existing data"
                        : "Calculated and report generated into centerlist and newenvelopebreakage",
                    data = envelopesToUse,
                    fileName = fileName
                });
            }
            catch (Exception ex)
            {
                await _loggerService.LogErrorAsync("Error creating ExtraEnvelope for NrData1", ex.ToString(), nameof(NrData1Controller));
                return StatusCode(500, "Internal server error: " + ex.Message);
            }
        }

        [HttpPost("ProcessEnvelopeBreaking")]
        public async Task<IActionResult> ProcessEnvelopeBreaking(int ProjectId, int triggeredBy = 0, [FromQuery] bool skipReset = false, [FromQuery] int? lotNo = null, [FromQuery] string? catchNo = null, [FromQuery] bool bypassDispatch = false, [FromQuery] int? batchNo = null)
        {
            var procController = new EnvelopeBreakageProcessingController(_context, _loggerService, _apiSettings, _dispatchService)
            {
                ControllerContext = this.ControllerContext
            };
            return await procController.ProcessEnvelopeBreaking(ProjectId, triggeredBy, skipReset, lotNo, catchNo, bypassDispatch, batchNo);
        }

        [HttpGet("GetEnvelopeBreakingReport")]
        public async Task<IActionResult> GetEnvelopeBreakingReport(int ProjectId, [FromQuery] int? lotNo = null)
        {
            var procController = new EnvelopeBreakageProcessingController(_context, _loggerService, _apiSettings, _dispatchService)
            {
                ControllerContext = this.ControllerContext
            };
            return await procController.GetEnvelopeBreakingReport(ProjectId, lotNo);
        }

        #region Helper Models
        public class NodalValueConfig
        {
            public string? NodalCodes { get; set; }
            public string? Value { get; set; }
        }

        public class RangeConfigModel
        {
            public List<RangeItem>? ranges { get; set; }
        }

        public class RangeItem
        {
            public int from { get; set; }
            public int to { get; set; }
            public int value { get; set; }
        }
        public class ExtraEnvelopeReportDto
        {
            public string? CatchNo { get; set; }
            public string? NodalCode { get; set; }
            public int ExtraId { get; set; }
            public int Quantity { get; set; }
            public string? InnerEnvelope { get; set; }
            public string? OuterEnvelope { get; set; }
        }
        #endregion

        private class CatchGroupData
        {
            public NrData1 NrData1 { get; set; } = null!;
            public Dictionary<string, string> UniqueExtraData { get; set; } = new(StringComparer.OrdinalIgnoreCase);
            public List<CenterList> Centers { get; set; } = new();
        }

        private static object? ConvertPropertyValue(PropertyInfo propInfo, string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return null;

            try
            {
                var targetType = Nullable.GetUnderlyingType(propInfo.PropertyType) ?? propInfo.PropertyType;

                if (targetType == typeof(bool))
                {
                    if (bool.TryParse(value, out bool bVal)) return bVal;
                    if (value.Equals("1") || value.Equals("true", StringComparison.OrdinalIgnoreCase)) return true;
                    if (value.Equals("0") || value.Equals("false", StringComparison.OrdinalIgnoreCase)) return false;
                    return null;
                }

                if (targetType == typeof(int))
                {
                    if (int.TryParse(value, out int iVal)) return iVal;
                    if (double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out double dVal)) return (int)dVal;
                    return null;
                }

                if (targetType == typeof(double))
                {
                    if (double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out double dVal)) return dVal;
                    return null;
                }

                if (targetType == typeof(DateTime))
                {
                    if (DateTime.TryParse(value, out DateTime dtVal)) return dtVal;
                    return null;
                }

                return Convert.ChangeType(value, targetType, CultureInfo.InvariantCulture);
            }
            catch
            {
                return null;
            }
        }

        private class EnvelopeTypeConfig
        {
            public string? Inner { get; set; }
            public string? Outer { get; set; }
        }

        private List<ExtraEnvelopes> ProcessExtraEnvelopes(
            string? catchNo,
            string? centerCode,
            int quantity,
            int nrQuantity,
            int projectId,
            List<ExtrasConfiguration> extraConfigs,
            Dictionary<string, int>? extraTypesByName = null)
        {
            var extraEnvelopesToAdd = new List<ExtraEnvelopes>();
            int? extraTypeId = null;
            if (!string.IsNullOrWhiteSpace(centerCode) && extraTypesByName != null && extraTypesByName.TryGetValue(centerCode.Trim(), out var foundId))
            {
                extraTypeId = foundId;
            }

            if (extraTypeId.HasValue)
            {
                var config = extraConfigs.FirstOrDefault(x => x.ExtraType == extraTypeId);
                if (config != null)
                {
                    EnvelopeTypeConfig? envelopeType = null;
                    if (!string.IsNullOrWhiteSpace(config.EnvelopeType))
                    {
                        try
                        {
                            envelopeType = JsonSerializer.Deserialize<EnvelopeTypeConfig>(config.EnvelopeType);
                        }
                        catch { }
                    }

                    int innerCapacity = envelopeType != null && !string.IsNullOrWhiteSpace(envelopeType.Inner) ? GetEnvelopeCapacity(envelopeType.Inner) : 1;
                    int outerCapacity = envelopeType != null && !string.IsNullOrWhiteSpace(envelopeType.Outer) ? GetEnvelopeCapacity(envelopeType.Outer) : 1;
                    int baseQty = nrQuantity > 0 ? nrQuantity : quantity;
                    int roundedQty = baseQty;

                    if (innerCapacity > 0)
                        roundedQty = (int)Math.Ceiling((double)baseQty / innerCapacity) * innerCapacity;
                    else if (outerCapacity > 0)
                        roundedQty = (int)Math.Ceiling((double)baseQty / outerCapacity) * outerCapacity;

                    extraEnvelopesToAdd.Add(new ExtraEnvelopes
                    {
                        ProjectId = projectId,
                        CatchNo = catchNo,
                        ExtraId = extraTypeId.Value,
                        Quantity = roundedQty,
                        InnerEnvelope = innerCapacity > 0 ? Math.Ceiling((double)roundedQty / innerCapacity).ToString() : "0",
                        OuterEnvelope = outerCapacity > 0 ? Math.Ceiling((double)roundedQty / outerCapacity).ToString() : "0",
                        Status = 1
                    });
                }
            }

            return extraEnvelopesToAdd;
        }

        private List<ExtraEnvelopes> GroupAndProcessExtras(
            List<ExtraEnvelopes> extras,
            int projectId,
            List<ExtrasConfiguration> extraConfigs)
        {
            var grouped = extras
                .GroupBy(e => new { e.ExtraId, e.CatchNo })
                .Select(g => new ExtraEnvelopes
                {
                    ProjectId = projectId,
                    ExtraId = g.Key.ExtraId,
                    CatchNo = g.Key.CatchNo,
                    Quantity = g.Sum(e => e.Quantity),
                    Status = 1
                })
                .ToList();

            foreach (var extra in grouped)
            {
                var config = extraConfigs.FirstOrDefault(x => x.ExtraType == extra.ExtraId);
                if (config != null)
                {
                    EnvelopeTypeConfig? envelopeType = null;
                    if (!string.IsNullOrWhiteSpace(config.EnvelopeType))
                    {
                        try
                        {
                            envelopeType = JsonSerializer.Deserialize<EnvelopeTypeConfig>(config.EnvelopeType);
                        }
                        catch { }
                    }

                    int innerCap = envelopeType != null && !string.IsNullOrWhiteSpace(envelopeType.Inner) ? GetEnvelopeCapacity(envelopeType.Inner) : 1;
                    int outerCap = envelopeType != null && !string.IsNullOrWhiteSpace(envelopeType.Outer) ? GetEnvelopeCapacity(envelopeType.Outer) : 1;

                    extra.InnerEnvelope = innerCap > 0 ? Math.Ceiling((double)extra.Quantity / innerCap).ToString() : "0";
                    extra.OuterEnvelope = outerCap > 0 ? Math.Ceiling((double)extra.Quantity / outerCap).ToString() : "0";
                }
            }

            return grouped;
        }

        private int GetEnvelopeCapacity(string envelopeCode)
        {
            if (string.IsNullOrWhiteSpace(envelopeCode))
                return 0;

            var numberPart = new string(envelopeCode.Where(char.IsDigit).ToArray());
            return int.TryParse(numberPart, out var capacity) ? capacity : 0;
        }
    }
}
