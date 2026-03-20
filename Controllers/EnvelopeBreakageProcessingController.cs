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
using System.Globalization;

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
                    .Where(p => p.ProjectId == ProjectId && p.Status == true)
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
                    string NodalCode, string CenterCode, double CenterSort, double NodalSort, int RouteSort, string Route, int nrDataId)
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
                    //string currentMergeField = $"{extra.CatchNo}-{extra.ExtraId}";
                    extraCenterEnvCounter = 0;
                  /*  if (currentMergeField != prevExtraMergeField)
                    {
                        extraCenterEnvCounter = 0;
                        prevExtraMergeField = currentMergeField;
                    }*/

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

                        double modifiedCenterSort = extra.ExtraId switch
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

                        catch (Exception ex)
                        {
                            Console.WriteLine($"JSON Error: {envInfo.OuterEnvelope}");
                            Console.WriteLine(ex.Message);
                        }
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

                // Step 1: Sort resultList FIRST using EnvelopeMakingCriteria
                var sortFields = await _context.Fields
                    .Where(f => projectconfig.EnvelopeMakingCriteria.Contains(f.FieldId))
                    .ToListAsync();

                var sortingFieldNames = sortFields
                    .OrderBy(f => projectconfig.EnvelopeMakingCriteria.IndexOf(f.FieldId))
                    .Select(f => f.Name)
                    .ToList();

                IOrderedEnumerable<dynamic> sortedResultList = null;

                foreach (var fieldName in sortingFieldNames)
                {
                    Func<dynamic, object> keySelector = x =>
                    {
                        var dict = (IDictionary<string, object>)x;
                        if (!dict.ContainsKey(fieldName)) return null;
                        var val = dict[fieldName];
                        if (val == null) return null;

                        return fieldName switch
                        {
                            "RouteSort" => (object)(int.TryParse(val.ToString(), out int r) ? r : 0),
                            "CenterSort" => (object)(int.TryParse(val.ToString(), out int c) ? c : 0),
                            "NodalSort" => (object)(double.TryParse(val.ToString(), out double n) ? n : 0.0),
                            "ExamDate" => DateTime.TryParseExact(val.ToString(), "dd-MM-yyyy",
                                                CultureInfo.InvariantCulture, DateTimeStyles.None, out var d)
                                                ? (object)d : DateTime.MinValue,
                            _ => val.ToString().Trim()
                        };
                    };

                    sortedResultList = sortedResultList == null
                        ? resultList.OrderBy(keySelector)
                        : sortedResultList.ThenBy(keySelector);
                }

                var finalResultList = sortedResultList?.ToList() ?? resultList;

                // Step 2: Now assign serial, OmrSerial, BookletSerial on the SORTED list
                int bookletStart = projectconfig?.BookletSerialNumber ?? 0;
                int omrStart = projectconfig.OmrSerialNumber;
                bool assignBookletSerial = bookletStart > 0;
                bool assignOmrSerial = omrStart > 0;
                bool resetOmrSerialOnCatchChange = projectconfig.ResetOmrSerialOnCatchChange;
                bool resetBookletSerialOnCatchChange = projectconfig?.ResetBookletSerialOnCatchChange ?? false;
                string prevCatchForSerial = null;
                int serial = 1;

              


                var envelopeResults = new List<EnvelopeBreakingResult>();

                foreach (var item in finalResultList)
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
                        omrStart = projectconfig.OmrSerialNumber;
                    }

                    if (resetBookletSerialOnCatchChange && prevCatchForSerial != null && catchNo != prevCatchForSerial)
                    {
                        bookletStart = projectconfig?.BookletSerialNumber??0;
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
                    string omrSerial = "";
                    if (assignBookletSerial)
                    {
                        int envQuantity = (int)dict["EnvQuantity"];
                        bookletSerial = $"{bookletStart}-{bookletStart + envQuantity - 1}";
                        bookletStart += envQuantity;
                    }

                    if (assignOmrSerial)
                    {
                        int envQuantity = (int)dict["EnvQuantity"];
                        omrSerial = $"{omrStart}-{omrStart + envQuantity - 1}";
                        omrStart += envQuantity;
                    }

                    double? modifiedNodalSort = null;
                    int? modifiedCenterSort = null;
                    int? modifiedRouteSort = null;
                    string nodalCodeRef = null;
                    string routeRef = null;

                    if (isExtra)
                    {
                        // Safe casting for nullable int fields
                        if (dict["CenterSort"] is double dModCenterSort)
                            modifiedCenterSort = (int)dModCenterSort;
                        else if (dict["CenterSort"] is int iModCenterSort)
                            modifiedCenterSort = iModCenterSort;
                        else if (int.TryParse(dict["CenterSort"]?.ToString(), out int parsedModCenterSort))
                            modifiedCenterSort = parsedModCenterSort;

                        // Safe casting for double
                        if (dict["NodalSort"] is double dModNodalSort)
                            modifiedNodalSort = dModNodalSort;
                        else if (dict["NodalSort"] is int iModNodalSort)
                            modifiedNodalSort = (double)iModNodalSort;
                        else if (double.TryParse(dict["NodalSort"]?.ToString(), out double parsedModNodalSort))
                            modifiedNodalSort = parsedModNodalSort;

                        // Safe casting for nullable int fields
                        if (dict["RouteSort"] is double dModRouteSort)
                            modifiedRouteSort = (int)dModRouteSort;
                        else if (dict["RouteSort"] is int iModRouteSort)
                            modifiedRouteSort = iModRouteSort;
                        else if (int.TryParse(dict["RouteSort"]?.ToString(), out int parsedModRouteSort))
                            modifiedRouteSort = parsedModRouteSort;

                        nodalCodeRef = dict["NodalCode"]?.ToString();
                        routeRef = dict["Route"]?.ToString();
                    }

                    // Safe casting for numeric fields that might come as double
                    int centerSort = 0;
                    if (dict["CenterSort"] is double dCenterSort)
                        centerSort = (int)dCenterSort;
                    else if (dict["CenterSort"] is int iCenterSort)
                        centerSort = iCenterSort;
                    else if (int.TryParse(dict["CenterSort"]?.ToString(), out int parsedCenterSort))
                        centerSort = parsedCenterSort;

                    int routeSort = 0;
                    if (dict["RouteSort"] is double dRouteSort)
                        routeSort = (int)dRouteSort;
                    else if (dict["RouteSort"] is int iRouteSort)
                        routeSort = iRouteSort;
                    else if (int.TryParse(dict["RouteSort"]?.ToString(), out int parsedRouteSort))
                        routeSort = parsedRouteSort;

                    double nodalSort = 0.0;
                    if (dict["NodalSort"] is double dNodalSort)
                        nodalSort = dNodalSort;
                    else if (dict["NodalSort"] is int iNodalSort)
                        nodalSort = (double)iNodalSort;
                    else if (double.TryParse(dict["NodalSort"]?.ToString(), out double parsedNodalSort))
                        nodalSort = parsedNodalSort;

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
                        OmrSerial = omrSerial,
                        CenterCode = dict["CenterCode"]?.ToString(),
                        CenterSort = centerSort,
                        ExamTime = dict["ExamTime"]?.ToString(),
                        ExamDate = dict["ExamDate"]?.ToString(),
                        Quantity = (int)dict["Quantity"],
                        NodalCode = dict["NodalCode"]?.ToString(),
                        NodalSort = nodalSort,
                        Route = dict["Route"]?.ToString(),
                        RouteSort = routeSort,
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


        [HttpGet("CatchWithOmrSerialing")]
        public async Task<IActionResult> ProcessSerialingReport(int ProjectId)
        {
            try
            {
                var data = await _context.EnvelopeBreakingResults
                    .Where(x => x.ProjectId == ProjectId && (x.BookletSerial != null || x.OmrSerial != null))
                    .OrderBy(x => x.Id)
                    .ToListAsync();

                var grouped = data
                    .GroupBy(x => x.CatchNo)
                    .Select(g =>
                    {
                        var firstSerial = g.First().OmrSerial;
                        var lastSerial = g.Last().OmrSerial;
                        var firstBooklet = g.First().BookletSerial;
                        var lastBooklet = g.Last().BookletSerial;
                        var start = firstSerial.Split('-')[0];
                        var end = lastSerial.Split('-')[1];
                        var first = firstBooklet.Split('-')[0];
                        var last = lastBooklet.Split('-')[0];
                        return new
                        {
                            CatchNo = g.Key,
                            OmrSerialRange = $"{start}-{end}",
                            BookletSerialRange = $"{first}-{last}",
                            ExamDate = g.First().ExamDate,
                            ExamTime = g.First().ExamTime
                        };
                    })
                    .ToList();

                var reportPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", ProjectId.ToString());
                Directory.CreateDirectory(reportPath);

                var filePath = Path.Combine(reportPath, "CatchWiseBookletAndOmrSerialing.xlsx");
                if (System.IO.File.Exists(filePath))
                {
                    System.IO.File.Delete(filePath);
                }
                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("Serial Report");

                    worksheet.Cells[1, 1].Value = "Catch No";
                    worksheet.Cells[1, 2].Value = "Omr Serial Range";
                    worksheet.Cells[1, 3].Value = "Booklet Serial Range";
                    worksheet.Cells[1, 4].Value = "Exam Date";
                    worksheet.Cells[1, 5].Value = "Exam Time";

                    int row = 2;

                    foreach (var item in grouped)
                    {
                        worksheet.Cells[row, 1].Value = item.CatchNo;
                        worksheet.Cells[row, 2].Value = item.OmrSerialRange;
                        worksheet.Cells[row,3].Value = item.BookletSerialRange;
                        worksheet.Cells[row, 3].Value = item.ExamDate;
                        worksheet.Cells[row, 4].Value = item.ExamTime;
                        row++;
                    }

                    worksheet.Cells.AutoFitColumns();
                    worksheet.View.FreezePanes(2, 1);
                    package.SaveAs(new FileInfo(filePath));
                }
                return Ok(new { message = "Excel generated successfully", path = filePath });
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
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

                var nrDataDict = await _context.NRDatas
    .Where(p => p.ProjectId == ProjectId)
    .ToDictionaryAsync(p => p.Id, p => p.NRQuantity);

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
                    int nrQty = 0;

                    if (result.NrDataId != 0 && nrDataDict.TryGetValue(result.NrDataId, out var qty))
                    {
                        nrQty = qty;
                    }

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
                    rowDict["NRQuantity"] = nrQty;
                    rowDict["CourseName"] = result.CourseName ?? "";

                    // Envelope breaking fields
                    rowDict["EnvQuantity"] = result.EnvQuantity;
                    rowDict["CenterEnv"] = result.CenterEnv;
                    rowDict["TotalEnv"] = result.TotalEnv;
                    rowDict["Env"] = result.Env ?? "";
                    rowDict["SerialNumber"] = result.SerialNumber;
                    rowDict["BookletSerial"] = result.BookletSerial ?? "";
                    rowDict["OmrSerial"] = result.OmrSerial;
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
                        "Exam Time", "Exam Date", "BookletSerial","OmrSerial", "CourseName" };

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
                        ws.Cells[rowIdx, 18].Value = dict.ContainsKey("OmrSerial") ? dict["OmrSerial"] : "";
                        ws.Cells[rowIdx, 19].Value = dict.ContainsKey("CourseName") ? dict["CourseName"] : "";
                        rowIdx++;
                    }

                    ws.Cells[ws.Dimension.Address].AutoFitColumns();
                    ws.View.FreezePanes(2, 1);
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
