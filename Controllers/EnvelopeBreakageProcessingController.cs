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
using OfficeOpenXml;
using System.Reflection;
using Tools.Services;
using Microsoft.Extensions.Options;

namespace Tools.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class EnvelopeBreakageProcessingController : ControllerBase
    {
        private readonly ERPToolsDbContext _context;
        private readonly ILoggerService _loggerService;
        private readonly ApiSettings _apiSettings;

        public EnvelopeBreakageProcessingController(ERPToolsDbContext context, ILoggerService loggerService, IOptions<ApiSettings> apiSettings)
        {
            _context = context;
            _loggerService = loggerService;
            _apiSettings = apiSettings.Value;
        }

        [HttpPost("ProcessEnvelopeBreaking")]
        public async Task<IActionResult> ProcessEnvelopeBreaking(int ProjectId)
        {
            try
            {
                var envCaps = await _context.EnvelopesTypes
                    .Select(e => new { e.EnvelopeName, e.Capacity })
                    .ToListAsync();

                var nrData = await _context.NRDatas
                    .Where(p => p.ProjectId == ProjectId)
                    .OrderBy(p => p.CatchNo)
                    .ThenBy(p => p.RouteSort)
                    .ThenBy(p => p.NodalSort)
                    .ThenBy(p => p.CenterSort)
                    .ToListAsync();

                var envBreaking = await _context.EnvelopeBreakages
                    .Where(p => p.ProjectId == ProjectId)
                    .ToListAsync();

                var extras = await _context.ExtrasEnvelope
                    .Where(p => p.ProjectId == ProjectId)
                    .ToListAsync();

                var projectconfig = await _context.ProjectConfigs
                    .Where(p => p.ProjectId == ProjectId)
                    .FirstOrDefaultAsync();

                if (projectconfig == null)
                    return NotFound("Project config not found");

                var envelopeCapacities = envCaps.ToDictionary(x => x.EnvelopeName, x => x.Capacity);
                var envDict = envBreaking.ToDictionary(
                    e => e.NrDataId,
                    e => (e.TotalEnvelope, e.OuterEnvelope)
                );

                var resultList = new List<dynamic>();
                string prevNodalCode = null;
                string prevRoute = null;
                int prevRouteSort = 0;
                double prevNodalSort = 0;
                string prevCatchNo = null;
                string prevMergeField = null;
                string prevExtraMergeField = null;
                int centerEnvCounter = 0;
                int extraCenterEnvCounter = 0;
                var nodalExtrasAddedForNodalCatch = new HashSet<(string, string)>();
                var catchExtrasAdded = new HashSet<(int, string)>();
                var extrasconfig = await _context.ExtraConfigurations
                    .Where(p => p.ProjectId == ProjectId)
                    .ToListAsync();

                void AddExtraWithEnv(ExtraEnvelopes extra, string examDate, string examTime, string course, int NrQuantity,
                    string NodalCode, string CenterCode, int CenterSort, double NodalSort, int RouteSort, string Route, int nrDataId)
                {
                    var extraConfig = extrasconfig.FirstOrDefault(e => e.ExtraType == extra.ExtraId);
                    int envCapacity = 0;

                    if (extraConfig != null && !string.IsNullOrEmpty(extraConfig.EnvelopeType))
                    {
                        var envType = JsonSerializer.Deserialize<Dictionary<string, string>>(extraConfig.EnvelopeType);
                        if (envType != null && envType.TryGetValue("Outer", out string outerType)
                            && envelopeCapacities.TryGetValue(outerType, out int cap))
                        {
                            envCapacity = cap;
                        }
                    }

                    int totalEnv = (int)Math.Ceiling((double)extra.Quantity / envCapacity);
                    string currentMergeField = $"{extra.CatchNo}-{extra.ExtraId}";

                    if (currentMergeField != prevExtraMergeField)
                    {
                        extraCenterEnvCounter = 0;
                        prevExtraMergeField = currentMergeField;
                    }

                    for (int j = 1; j <= totalEnv; j++)
                    {
                        int envQuantity;
                        if (j == 1 && totalEnv > 1)
                        {
                            envQuantity = extra.Quantity - (envCapacity * (totalEnv - 1));
                        }
                        else
                        {
                            envQuantity = Math.Min(extra.Quantity - (envCapacity * (j - 1)), envCapacity);
                        }
                        extraCenterEnvCounter++;

                        int modifiedCenterSort = extra.ExtraId switch
                        {
                            1 => 10000,
                            2 => 100000,
                            3 => 1000000,
                            _ => CenterSort
                        };

                        double modifiedNodalSort = extra.ExtraId switch
                        {
                            1 => NodalSort + 0.1,
                            2 => 100000,
                            3 => 1000000,
                            _ => NodalSort
                        };

                        int modifiedRouteSort = extra.ExtraId switch
                        {
                            1 => RouteSort,
                            2 => 10000,
                            3 => 100000,
                            _ => RouteSort
                        };

                        var extraRow = new System.Dynamic.ExpandoObject();
                        var extraDict = (IDictionary<string, object>)extraRow;

                        extraDict["ExtraAttached"] = true;
                        extraDict["ExtraId"] = extra.ExtraId;
                        extraDict["NrDataId"] = nrDataId;
                        extraDict["CatchNo"] = extra.CatchNo;
                        extraDict["Quantity"] = extra.Quantity;
                        extraDict["EnvQuantity"] = envQuantity;
                        extraDict["InnerEnvelope"] = extra.InnerEnvelope;
                        extraDict["OuterEnvelope"] = extra.OuterEnvelope;
                        extraDict["CenterCode"] = extra.ExtraId switch
                        {
                            1 => "Nodal Extra",
                            2 => "University Extra",
                            3 => "Office Extra",
                            _ => "Extra"
                        };
                        extraDict["CenterEnv"] = extraCenterEnvCounter;
                        extraDict["ExamDate"] = examDate;
                        extraDict["CourseName"] = course;
                        extraDict["ExamTime"] = examTime;
                        extraDict["TotalEnv"] = totalEnv;
                        extraDict["Env"] = $"{j}/{totalEnv}";
                        extraDict["NRQuantity"] = NrQuantity;
                        extraDict["NodalCode"] = NodalCode;
                        extraDict["Route"] = Route;
                        extraDict["NodalSort"] = modifiedNodalSort;
                        extraDict["CenterSort"] = modifiedCenterSort;
                        extraDict["RouteSort"] = modifiedRouteSort;

                        resultList.Add(extraRow);
                    }
                }

                for (int i = 0; i < nrData.Count; i++)
                {
                    var current = nrData[i];
                    bool catchNoChanged = prevCatchNo != null && current.CatchNo != prevCatchNo;

                    if (catchNoChanged)
                    {
                        var prevNrData = nrData[i - 1];
                        if (!nodalExtrasAddedForNodalCatch.Contains((prevNrData.NodalCode, prevCatchNo)))
                        {
                            var extrasToAdd = extras.Where(e => e.ExtraId == 1 && e.CatchNo == prevCatchNo).ToList();
                            foreach (var extra in extrasToAdd)
                            {
                                AddExtraWithEnv(extra, prevNrData.ExamDate, prevNrData.ExamTime, prevNrData.CourseName,
                                    prevNrData.NRQuantity, prevNrData.NodalCode, prevNrData.CenterCode, prevNrData.CenterSort,
                                    prevNrData.NodalSort, prevNrData.RouteSort, prevNrData.Route, prevNrData.Id);
                            }
                            nodalExtrasAddedForNodalCatch.Add((prevNrData.NodalCode, prevCatchNo));
                        }

                        foreach (var extraId in new[] { 2, 3 })
                        {
                            if (!catchExtrasAdded.Contains((extraId, prevCatchNo)))
                            {
                                var extrasToAdd = extras.Where(e => e.ExtraId == extraId && e.CatchNo == prevCatchNo).ToList();
                                foreach (var extra in extrasToAdd)
                                {
                                    AddExtraWithEnv(extra, prevNrData.ExamDate, prevNrData.ExamTime, prevNrData.CourseName,
                                        prevNrData.NRQuantity, prevNrData.NodalCode, prevNrData.CenterCode, prevNrData.CenterSort,
                                        prevNrData.NodalSort, prevNrData.RouteSort, prevNrData.Route, prevNrData.Id);
                                }
                                catchExtrasAdded.Add((extraId, prevCatchNo));
                            }
                        }
                    }

                    if (!catchNoChanged && prevNodalCode != null && current.NodalCode != prevNodalCode)
                    {
                        if (!nodalExtrasAddedForNodalCatch.Contains((prevNodalCode, current.CatchNo)))
                        {
                            var extrasToAdd = extras.Where(e => e.ExtraId == 1 && e.CatchNo == current.CatchNo).ToList();
                            foreach (var extra in extrasToAdd)
                            {
                                AddExtraWithEnv(extra, current.ExamDate, current.ExamTime, current.CourseName,
                                    current.NRQuantity, prevNodalCode, current.CenterCode, current.CenterSort,
                                    prevNodalSort, prevRouteSort, prevRoute, current.Id);
                            }
                            nodalExtrasAddedForNodalCatch.Add((prevNodalCode, current.CatchNo));
                        }
                    }

                    int totalEnv = envDict.TryGetValue(current.Id, out var envData) && envData.TotalEnvelope > 0
                        ? envData.TotalEnvelope
                        : 1;

                    string currentMergeField = $"{current.CatchNo}-{current.CenterCode}";
                    if (currentMergeField != prevMergeField)
                    {
                        centerEnvCounter = 0;
                        prevMergeField = currentMergeField;
                    }

                    var envelopeBreakdown = new List<(string EnvType, int Count, int Capacity)>();

                    if (envDict.TryGetValue(current.Id, out var envInfo) && !string.IsNullOrEmpty(envInfo.OuterEnvelope))
                    {
                        try
                        {
                            var outerEnvDict = JsonSerializer.Deserialize<Dictionary<string, string>>(envInfo.OuterEnvelope);
                            if (outerEnvDict != null)
                            {
                                foreach (var kvp in outerEnvDict)
                                {
                                    if (int.TryParse(kvp.Value, out int count) && count > 0)
                                    {
                                        int capacity = 0;
                                        if (envelopeCapacities.TryGetValue(kvp.Key, out int cap))
                                        {
                                            capacity = cap;
                                        }
                                        else
                                        {
                                            var match = System.Text.RegularExpressions.Regex.Match(kvp.Key, @"\d+");
                                            if (match.Success)
                                            {
                                                capacity = int.Parse(match.Value);
                                            }
                                        }

                                        envelopeBreakdown.Add((kvp.Key, count, capacity));
                                    }
                                }
                            }
                        }
                        catch { }
                    }

                    envelopeBreakdown = envelopeBreakdown.OrderBy(x => x.Capacity).ToList();

                    if (envelopeBreakdown.Count > 0)
                    {
                        int envelopeIndex = 1;
                        int remainingQty = current.Quantity;
                        foreach (var (envType, count, capacity) in envelopeBreakdown)
                        {
                            for (int k = 0; k < count; k++)
                            {
                                centerEnvCounter++;
                                int envQty;
                                if (remainingQty > capacity)
                                {
                                    envQty = capacity;
                                }
                                else
                                {
                                    envQty = remainingQty;
                                }

                                var nrRow = new System.Dynamic.ExpandoObject();
                                var nrDict = (IDictionary<string, object>)nrRow;

                                nrDict["ExtraAttached"] = false;
                                nrDict["CatchNo"] = current.CatchNo;
                                nrDict["CenterCode"] = current.CenterCode;
                                nrDict["CourseName"] = current.CourseName;
                                nrDict["ExamTime"] = current.ExamTime;
                                nrDict["ExamDate"] = current.ExamDate;
                                nrDict["Quantity"] = current.Quantity;
                                nrDict["EnvQuantity"] = envQty;
                                nrDict["NodalCode"] = current.NodalCode;
                                nrDict["CenterEnv"] = centerEnvCounter;
                                nrDict["TotalEnv"] = totalEnv;
                                nrDict["Env"] = $"{envelopeIndex}/{totalEnv}";
                                nrDict["NRQuantity"] = current.NRQuantity;
                                nrDict["CenterSort"] = current.CenterSort;
                                nrDict["NodalSort"] = current.NodalSort;
                                nrDict["Route"] = current.Route;
                                nrDict["RouteSort"] = current.RouteSort;
                                nrDict["NrDataId"] = current.Id;
                                nrDict["ExtraId"] = (int?)null;

                                resultList.Add(nrRow);

                                remainingQty -= envQty;
                                envelopeIndex++;
                                if (remainingQty <= 0)
                                    break;
                            }
                            if (remainingQty <= 0)
                                break;
                        }
                    }

                    prevNodalCode = current.NodalCode;
                    prevNodalSort = current.NodalSort;
                    prevRouteSort = current.RouteSort;
                    prevRoute = current.Route;
                    prevCatchNo = current.CatchNo;
                }

                if (prevCatchNo != null)
                {
                    var lastNrData = nrData.LastOrDefault();
                    if (lastNrData != null)
                    {
                        if (!nodalExtrasAddedForNodalCatch.Contains((lastNrData.NodalCode, prevCatchNo)))
                        {
                            var extrasToAdd = extras.Where(e => e.ExtraId == 1 && e.CatchNo == prevCatchNo).ToList();
                            foreach (var extra in extrasToAdd)
                            {
                                AddExtraWithEnv(extra, lastNrData.ExamDate, lastNrData.ExamTime, lastNrData.CourseName,
                                    lastNrData.NRQuantity, lastNrData.NodalCode, lastNrData.CenterCode, lastNrData.CenterSort,
                                    lastNrData.NodalSort, lastNrData.RouteSort, lastNrData.Route, lastNrData.Id);
                            }
                        }

                        foreach (var extraId in new[] { 2, 3 })
                        {
                            if (!catchExtrasAdded.Contains((extraId, prevCatchNo)))
                            {
                                var extrasToAdd = extras.Where(e => e.ExtraId == extraId && e.CatchNo == prevCatchNo).ToList();
                                foreach (var extra in extrasToAdd)
                                {
                                    AddExtraWithEnv(extra, lastNrData.ExamDate, lastNrData.ExamTime, lastNrData.CourseName,
                                        lastNrData.NRQuantity, lastNrData.NodalCode, lastNrData.CenterCode, lastNrData.CenterSort,
                                        lastNrData.NodalSort, lastNrData.RouteSort, lastNrData.Route, lastNrData.Id);
                                }
                            }
                        }
                    }
                }

                var maxBatch = await _context.EnvelopeBreakingResults
                    .Where(r => r.ProjectId == ProjectId)
                    .MaxAsync(r => (int?)r.UploadBatch) ?? 0;
                int currentBatch = maxBatch + 1;

                int currentStartNumber = projectconfig.OmrSerialNumber;
                bool assignBookletSerial = currentStartNumber > 0;
                bool resetOmrSerialOnCatchChange = projectconfig.ResetOmrSerialOnCatchChange;
                string prevCatchForSerial = null;
                int serial = 1;

                var envelopeResults = new List<EnvelopeBreakingResult>();

                foreach (var item in resultList)
                {
                    var dict = (IDictionary<string, object>)item;
                    bool isExtra = (bool)dict["ExtraAttached"];
                    string catchNo = dict["CatchNo"]?.ToString();

                    // ✅ Reset serial number whenever CatchNo changes (not just for OmrSerial)
                    if (prevCatchForSerial != null && catchNo != prevCatchForSerial)
                    {
                        serial = 1;
                    }

                    // ✅ Reset OmrSerial only if the flag is set
                    if (resetOmrSerialOnCatchChange && prevCatchForSerial != null && catchNo != prevCatchForSerial)
                    {
                        currentStartNumber = projectconfig.OmrSerialNumber;
                    }

                    int? nrDataId = null;
                    int? extraId = null;

                    if (isExtra)
                    {
                        extraId = (int?)dict["ExtraId"];
                        nrDataId = (int?)dict["NrDataId"];
                    }
                    else
                    {
                        nrDataId = (int?)dict["NrDataId"];
                    }

                    string bookletSerial = "";
                    if (assignBookletSerial)
                    {
                        int envQuantity = (int)dict["EnvQuantity"];
                        bookletSerial = $"{currentStartNumber}-{currentStartNumber + envQuantity - 1}";
                        currentStartNumber += envQuantity;
                    }

                    double? modifiedNodalSort = null;
                    int? modifiedCenterSort = null;
                    int? modifiedRouteSort = null;
                    string nodalCodeRef = null;
                    string routeRef = null;

                    if (isExtra)
                    {
                        modifiedCenterSort = (int?)dict["CenterSort"];
                        modifiedNodalSort = Convert.ToDouble(dict["NodalSort"]);
                        modifiedRouteSort = (int?)dict["RouteSort"];
                        nodalCodeRef = dict["NodalCode"]?.ToString();
                        routeRef = dict["Route"]?.ToString();
                    }

                    envelopeResults.Add(new EnvelopeBreakingResult
                    {
                        ProjectId = ProjectId,
                        NrDataId = nrDataId ?? 0,
                        ExtraId = extraId,
                        CatchNo = catchNo,
                        EnvQuantity = (int)dict["EnvQuantity"],
                        CenterEnv = (int)dict["CenterEnv"],
                        TotalEnv = (int)dict["TotalEnv"],
                        Env = dict["Env"]?.ToString(),
                        SerialNumber = serial++,
                        BookletSerial = bookletSerial,
                        CenterCode = dict["CenterCode"]?.ToString(),
                        CenterSort = (int)dict["CenterSort"],
                        ExamTime = dict["ExamTime"]?.ToString(),
                        ExamDate = dict["ExamDate"]?.ToString(),
                        Quantity = (int)dict["Quantity"],
                        NodalCode = dict["NodalCode"]?.ToString(),
                        NodalSort = Convert.ToDouble(dict["NodalSort"]),
                        Route = dict["Route"]?.ToString(),
                        RouteSort = (int)dict["RouteSort"],
                        NRQuantity = (int)dict["NRQuantity"],
                        CourseName = dict["CourseName"]?.ToString(),
                        UploadBatch = currentBatch
                    });

                    prevCatchForSerial = catchNo;
                }

                _context.EnvelopeBreakingResults.AddRange(envelopeResults);
                await _context.SaveChangesAsync();

                _loggerService.LogEvent(
                    $"Saved {envelopeResults.Count} envelope breaking results for ProjectId {ProjectId}, Batch {currentBatch}",
                    "EnvelopeBreakageProcessing",
                    User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0,
                    ProjectId);

                return Ok(new
                {
                    message = "Envelope breaking data saved to database",
                    recordsCount = envelopeResults.Count,
                    uploadBatch = currentBatch
                });
            }
            catch (Exception ex)
            {
                _loggerService.LogError("Error processing envelope breaking", ex.Message, nameof(EnvelopeBreakageProcessingController));
                return StatusCode(500, new { error = ex.Message });
            }
        }
        [HttpGet("GetEnvelopeBreakingReport")]
        public async Task<IActionResult> GetEnvelopeBreakingReport(int ProjectId, int? uploadBatch = null)
        {
            try
            {
                var projectconfig = await _context.ProjectConfigs
                    .Where(p => p.ProjectId == ProjectId)
                    .FirstOrDefaultAsync();

                if (projectconfig == null)
                    return NotFound("Project config not found");

                var nrData = await _context.NRDatas
                    .Where(p => p.ProjectId == ProjectId)
                    .ToListAsync();

                var fields = await _context.Fields
                    .Where(f => projectconfig.EnvelopeMakingCriteria.Contains(f.FieldId))
                    .ToListAsync();

                var fieldNames = fields
                    .OrderBy(f => projectconfig.EnvelopeMakingCriteria.IndexOf(f.FieldId))
                    .Select(f => f.Name)
                    .ToList();

                // Get results from database
                IQueryable<EnvelopeBreakingResult> query = _context.EnvelopeBreakingResults
                    .Where(r => r.ProjectId == ProjectId);

                if (uploadBatch.HasValue)
                {
                    query = query.Where(r => r.UploadBatch == uploadBatch.Value);
                }
                else
                {
                    // Get latest batch
                    var maxBatch = await _context.EnvelopeBreakingResults
                        .Where(r => r.ProjectId == ProjectId)
                        .MaxAsync(r => (int?)r.UploadBatch);
                    if (maxBatch.HasValue)
                    {
                        query = query.Where(r => r.UploadBatch == maxBatch.Value);
                    }
                }

                var results = await query.ToListAsync();

                if (!results.Any())
                    return NotFound("No envelope breaking results found");

                // Build full data with all fields for sorting
                var fullData = new List<dynamic>();

                foreach (var result in results)
                {
                    var row = new System.Dynamic.ExpandoObject();
                    var rowDict = (IDictionary<string, object>)row;

                    // All fields already in DB - no need to join with NRData
                    rowDict["CatchNo"] = result.CatchNo ?? "";
                    rowDict["CenterCode"] = result.CenterCode ?? "";
                    // Use modified values if they exist (for extras), otherwise use original
                    rowDict["CenterSort"] = result.CenterSort;
                    rowDict["ExamTime"] = result.ExamTime ?? "";
                    rowDict["ExamDate"] = result.ExamDate ?? "";
                    rowDict["Quantity"] = result.Quantity;
                    rowDict["NodalCode"] = result.NodalCode ?? "";
                    rowDict["NodalSort"] = result.NodalSort;
                    rowDict["Route"] = result.Route ?? "";
                    rowDict["RouteSort"] =  result.RouteSort;
                    rowDict["NRQuantity"] = result.NRQuantity;
                    rowDict["CourseName"] = result.CourseName ?? "";

                    // Envelope breaking fields
                    rowDict["EnvQuantity"] = result.EnvQuantity;
                    rowDict["CenterEnv"] = result.CenterEnv;
                    rowDict["TotalEnv"] = result.TotalEnv;
                    rowDict["Env"] = result.Env ?? "";
                    rowDict["SerialNumber"] = result.SerialNumber;
                    rowDict["BookletSerial"] = result.BookletSerial ?? "";

                    fullData.Add(row);
                }

                // Apply sorting based on EnvelopeMakingCriteria
                IOrderedEnumerable<dynamic> ordered = null;

                foreach (var fieldName in fieldNames)
                {
                    Func<dynamic, object> keySelector = record =>
                    {
                        var dict = (IDictionary<string, object>)record;
                        if (!dict.ContainsKey(fieldName)) return null;

                        var val = dict[fieldName];
                        if (val == null) return null;

                        // Type-safe handling per field
                        switch (fieldName)
                        {
                            case "RouteSort":
                                if (int.TryParse(val.ToString(), out int routeSort))
                                    return routeSort;
                                return 0;
                            case "CenterSort":
                                if (int.TryParse(val.ToString(), out int centerSort))
                                    return centerSort;
                                return 0;
                            case "NodalSort":
                                if (double.TryParse(val.ToString(), out double nodalSort))
                                    return nodalSort;
                                return 0.0;
                            case "ExamDate":
                                if (DateTime.TryParseExact(val.ToString(), "dd-MM-yyyy",
                                    System.Globalization.CultureInfo.InvariantCulture,
                                    System.Globalization.DateTimeStyles.None, out DateTime examDate))
                                    return examDate;
                                return DateTime.MinValue;
                            default:
                                return val.ToString().Trim();
                        }
                    };

                    if (ordered == null)
                        ordered = fullData.OrderBy(keySelector);
                    else
                        ordered = ordered.ThenBy(keySelector);
                }

                var sortedList = ordered?.ToList() ?? fullData;

                // Calculate SerialNumber after sorting
                int serial = 1;
                string prevCatchNo = null;

                foreach (var item in sortedList)
                {
                    var dict = (IDictionary<string, object>)item;
                    string currentCatchNo = dict["CatchNo"]?.ToString();

                    if (prevCatchNo != null && currentCatchNo != prevCatchNo)
                    {
                        serial = 1;
                    }

                    dict["SerialNumber"] = serial++;
                    prevCatchNo = currentCatchNo;
                }

                // Generate Excel
                var reportPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", ProjectId.ToString());
                Directory.CreateDirectory(reportPath);

                var filePath = Path.Combine(reportPath, "EnvelopeBreakingFromDB.xlsx");
                if (System.IO.File.Exists(filePath))
                {
                    System.IO.File.Delete(filePath);
                }

                using (var package = new ExcelPackage())
                {
                    var ws = package.Workbook.Worksheets.Add("Envelope Report");

                    var headers = new[] { "Serial Number", "Catch No", "Center Code", "Center Sort", "Quantity", "EnvQuantity",
                        "Center Env", "Total Env", "Env", "NRQuantity", "Nodal Code", "Nodal Sort", "Route", "Route Sort",
                        "Exam Time", "Exam Date", "BookletSerial", "CourseName" };

                    for (int i = 0; i < headers.Length; i++)
                    {
                        ws.Cells[1, i + 1].Value = headers[i];
                        ws.Cells[1, i + 1].Style.Font.Bold = true;
                    }

                    int rowIdx = 2;
                    foreach (var item in sortedList)
                    {
                        var dict = (IDictionary<string, object>)item;

                        ws.Cells[rowIdx, 1].Value = dict.ContainsKey("SerialNumber") ? dict["SerialNumber"] : "";
                        ws.Cells[rowIdx, 2].Value = dict.ContainsKey("CatchNo") ? dict["CatchNo"] : "";
                        ws.Cells[rowIdx, 3].Value = dict.ContainsKey("CenterCode") ? dict["CenterCode"] : "";
                        ws.Cells[rowIdx, 4].Value = dict.ContainsKey("CenterSort") ? dict["CenterSort"] : 0;
                        ws.Cells[rowIdx, 5].Value = dict.ContainsKey("Quantity") ? dict["Quantity"] : 0;
                        ws.Cells[rowIdx, 6].Value = dict.ContainsKey("EnvQuantity") ? dict["EnvQuantity"] : 0;
                        ws.Cells[rowIdx, 7].Value = dict.ContainsKey("CenterEnv") ? dict["CenterEnv"] : 0;
                        ws.Cells[rowIdx, 8].Value = dict.ContainsKey("TotalEnv") ? dict["TotalEnv"] : 0;
                        ws.Cells[rowIdx, 9].Value = dict.ContainsKey("Env") ? dict["Env"] : "";
                        ws.Cells[rowIdx, 10].Value = dict.ContainsKey("NRQuantity") ? dict["NRQuantity"] : 0;
                        ws.Cells[rowIdx, 11].Value = dict.ContainsKey("NodalCode") ? dict["NodalCode"] : "";
                        ws.Cells[rowIdx, 12].Value = dict.ContainsKey("NodalSort") ? dict["NodalSort"] : 0;
                        ws.Cells[rowIdx, 13].Value = dict.ContainsKey("Route") ? dict["Route"] : "";
                        ws.Cells[rowIdx, 14].Value = dict.ContainsKey("RouteSort") ? dict["RouteSort"] : 0;
                        ws.Cells[rowIdx, 15].Value = dict.ContainsKey("ExamTime") ? dict["ExamTime"] : "";
                        ws.Cells[rowIdx, 16].Value = dict.ContainsKey("ExamDate") ? dict["ExamDate"] : "";
                        ws.Cells[rowIdx, 17].Value = dict.ContainsKey("BookletSerial") ? dict["BookletSerial"] : "";
                        ws.Cells[rowIdx, 18].Value = dict.ContainsKey("CourseName") ? dict["CourseName"] : "";
                        rowIdx++;
                    }

                    ws.Cells[ws.Dimension.Address].AutoFitColumns();
                    package.SaveAs(new FileInfo(filePath));
                }

                _loggerService.LogEvent(
                    $"Generated envelope breaking report from DB for ProjectId {ProjectId}",
                    "EnvelopeBreakageProcessing",
                    User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0,
                    ProjectId);

                return Ok(new
                {
                    message = "Report generated successfully",
                    filePath = filePath,
                    recordsCount = sortedList.Count
                });
            }
            catch (Exception ex)
            {
                _loggerService.LogError("Error generating envelope breaking report", ex.Message, nameof(EnvelopeBreakageProcessingController));
                return StatusCode(500, new { error = ex.Message });
            }
        }
    }
}
