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
using System.Dynamic;

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
                    .Where(p => p.ProjectId == ProjectId )
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

                var mssTypes = projectconfig.MssTypes;
                string mssMode = projectconfig.MssAttached?.ToLower();
                var mssData = await _context.Mss
                    .Where(m => mssTypes.Contains(m.Id))
                    .ToListAsync();

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

                // Helper: Create MSS rows for a given catch
                List<dynamic> CreateMssRows(string catchNo, string examDate, string examTime, string courseName)
                {
                    var rows = new List<dynamic>();
                    foreach (var mss in mssData)
                    {
                        var mssRow = new System.Dynamic.ExpandoObject();
                        var mssDict = (IDictionary<string, object>)mssRow;
                        mssDict["isMss"] = true;
                        mssDict["ExtraAttached"] = false;
                        mssDict["CatchNo"] = catchNo;
                        mssDict["CenterCode"] = "";
                        mssDict["CourseName"] = courseName;
                        mssDict["ExamTime"] = examTime;
                        mssDict["ExamDate"] = examDate;
                        mssDict["Quantity"] = 0;
                        mssDict["EnvQuantity"] = mss.MssType;
                        mssDict["NodalCode"] = "";
                        mssDict["CenterEnv"] = 0;
                        mssDict["TotalEnv"] = 0;
                        mssDict["Env"] = "";
                        mssDict["NRQuantity"] = 0;
                        mssDict["CenterSort"] = 0;
                        mssDict["NodalSort"] = 0.0;
                        mssDict["Route"] = "";
                        mssDict["RouteSort"] = 0;
                        mssDict["NrDataId"] = 0;
                        mssDict["ExtraId"] = (int?)null;
                        rows.Add(mssRow);
                    }
                    return rows;
                }

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
                    extraCenterEnvCounter = 0;

                    for (int j = 1; j <= totalEnv; j++)
                    {
                        int envQuantity;
                        if (j == 1 && totalEnv > 1)
                            envQuantity = extra.Quantity - (envCapacity * (totalEnv - 1));
                        else
                            envQuantity = Math.Min(extra.Quantity - (envCapacity * (j - 1)), envCapacity);

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

                        extraDict["isMss"] = false;
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
                        extraDict["Route"] = extra.ExtraId==1 ? Route : "";
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
                        ? envData.TotalEnvelope : 1;

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
                                            capacity = cap;
                                        else
                                        {
                                            var match = System.Text.RegularExpressions.Regex.Match(kvp.Key, @"\d+");
                                            if (match.Success) capacity = int.Parse(match.Value);
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
                                int envQty = remainingQty > capacity ? capacity : remainingQty;

                                var nrRow = new System.Dynamic.ExpandoObject();
                                var nrDict = (IDictionary<string, object>)nrRow;

                                nrDict["isMss"] = false;
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
                                if (remainingQty <= 0) break;
                            }
                            if (remainingQty <= 0) break;
                        }
                    }

                    prevNodalCode = current.NodalCode;
                    prevNodalSort = current.NodalSort;
                    prevRouteSort = current.RouteSort;
                    prevRoute = current.Route;
                    prevCatchNo = current.CatchNo;
                }

                // Final extras for last CatchNo
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

                // ✅ STEP 1: Sort only non-MSS rows
                var sortFields = await _context.Fields
                    .Where(f => projectconfig.EnvelopeMakingCriteria.Contains(f.FieldId))
                    .ToListAsync();

                var sortingFieldNames = sortFields
                    .OrderBy(f => projectconfig.EnvelopeMakingCriteria.IndexOf(f.FieldId))
                    .Select(f => f.Name)
                    .ToList();

                var nonMssRows = resultList
                    .Where(r =>
                    {
                        var d = (IDictionary<string, object>)r;
                        return d.ContainsKey("isMss") && !(bool)d["isMss"];
                    }).ToList();

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
                        ? nonMssRows.OrderBy(keySelector)
                        : sortedResultList.ThenBy(keySelector);
                }

                var sortedNonMss = sortedResultList?.Cast<dynamic>().ToList() ?? nonMssRows;

                // ✅ STEP 2: Re-insert MSS rows at correct positions after sorting
                var finalResultList = new List<dynamic>();
                string lastCatchForMss = null;
                var buffer = new List<dynamic>();

                foreach (var item in sortedNonMss)
                {
                    var itemDict = (IDictionary<string, object>)item;
                    string catchNo = itemDict["CatchNo"]?.ToString();

                    if (catchNo != lastCatchForMss && lastCatchForMss != null)
                    {
                        if (mssMode == "end")
                        {
                            finalResultList.AddRange(buffer);
                            var last = (IDictionary<string, object>)buffer.Last();
                            finalResultList.AddRange(CreateMssRows(
                                lastCatchForMss,
                                last["ExamDate"]?.ToString(),
                                last["ExamTime"]?.ToString(),
                                last["CourseName"]?.ToString()
                            ));
                        }
                        else if (mssMode == "start")
                        {
                            finalResultList.AddRange(buffer);
                            finalResultList.AddRange(CreateMssRows(
                                catchNo,
                                itemDict["ExamDate"]?.ToString(),
                                itemDict["ExamTime"]?.ToString(),
                                itemDict["CourseName"]?.ToString()
                            ));
                        }
                        buffer.Clear();
                    }

                    // Very first catch in start mode
                    if (mssMode == "start" && lastCatchForMss == null)
                    {
                        finalResultList.AddRange(CreateMssRows(
                            catchNo,
                            itemDict["ExamDate"]?.ToString(),
                            itemDict["ExamTime"]?.ToString(),
                            itemDict["CourseName"]?.ToString()
                        ));
                    }

                    buffer.Add(item);
                    lastCatchForMss = catchNo;
                }

                // Flush last buffer
                if (buffer.Count > 0)
                {
                    finalResultList.AddRange(buffer);
                    if (mssMode == "end")
                    {
                        var last = (IDictionary<string, object>)buffer.Last();
                        finalResultList.AddRange(CreateMssRows(
                            lastCatchForMss,
                            last["ExamDate"]?.ToString(),
                            last["ExamTime"]?.ToString(),
                            last["CourseName"]?.ToString()
                        ));
                    }
                }

                // ✅ STEP 3: Assign serials - skip MSS rows
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
                    bool isMssRow = dict.ContainsKey("isMss") && (bool)dict["isMss"];
                    bool isExtra = (bool)dict["ExtraAttached"];
                    string catchNo = dict["CatchNo"]?.ToString();

                    // Reset serials on catch change - but not triggered by MSS rows
                    if (!isMssRow && prevCatchForSerial != null && catchNo != prevCatchForSerial)
                    {
                        serial = 1;

                        if (resetOmrSerialOnCatchChange)
                            omrStart = projectconfig.OmrSerialNumber;

                        if (resetBookletSerialOnCatchChange)
                            bookletStart = projectconfig?.BookletSerialNumber ?? 0;
                    }

                    // MSS rows: save with blank serials, no serial increment
                    if (isMssRow)
                    {
                        envelopeResults.Add(new EnvelopeBreakingResult
                        {
                            ProjectId = ProjectId,
                            NrDataId = 0,
                            ExtraId = null,
                            CatchNo = catchNo,
                            EnvQuantity = dict["EnvQuantity"]?.ToString(),
                            CenterEnv = 0,
                            TotalEnv = 0,
                            Env = "",
                            SerialNumber = 0,
                            BookletSerial = "",
                            OmrSerial = "",
                            CenterCode = dict["CenterCode"]?.ToString(),
                            CenterSort = 0,
                            ExamTime = dict["ExamTime"]?.ToString(),
                            ExamDate = dict["ExamDate"]?.ToString(),
                            Quantity = 0,
                            NodalCode = "",
                            NodalSort = 0,
                            Route = "",
                            RouteSort = 0,
                            CourseName = dict["CourseName"]?.ToString(),
                            UploadBatch = currentBatch,
                        });
                        // Don't update prevCatchForSerial for MSS rows
                        continue;
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

                    int centerSort = 0;
                    if (dict["CenterSort"] is double dCenterSort) centerSort = (int)dCenterSort;
                    else if (dict["CenterSort"] is int iCenterSort) centerSort = iCenterSort;
                    else int.TryParse(dict["CenterSort"]?.ToString(), out centerSort);

                    int routeSort = 0;
                    if (dict["RouteSort"] is double dRouteSort) routeSort = (int)dRouteSort;
                    else if (dict["RouteSort"] is int iRouteSort) routeSort = iRouteSort;
                    else int.TryParse(dict["RouteSort"]?.ToString(), out routeSort);

                    double nodalSort = 0.0;
                    if (dict["NodalSort"] is double dNodalSort) nodalSort = dNodalSort;
                    else if (dict["NodalSort"] is int iNodalSort) nodalSort = (double)iNodalSort;
                    else double.TryParse(dict["NodalSort"]?.ToString(), out nodalSort);

                    envelopeResults.Add(new EnvelopeBreakingResult
                    {
                        ProjectId = ProjectId,
                        NrDataId = nrDataId ?? 0,
                        ExtraId = extraId,
                        CatchNo = catchNo,
                        EnvQuantity = dict["EnvQuantity"]?.ToString(),
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
                        UploadBatch = currentBatch,
                    });

                    prevCatchForSerial = catchNo;
                }

                _context.EnvelopeBreakingResults.AddRange(envelopeResults);
                foreach (var nr in nrData)
                {
                    nr.Steps = 3; // Assuming NRData has a Step property
                }
                await _context.SaveChangesAsync();

                _loggerService.LogEvent(
                    $"Saved {envelopeResults.Count} envelope breaking results for ProjectId {ProjectId}, Batch {currentBatch}",
                    "EnvelopeBreakageProcessing",
                    User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0,
                    ProjectId);
                using var client = new HttpClient();
                var response = await client.GetAsync($"{_apiSettings.EnvelopeBreaking}?ProjectId={ProjectId}");
                if (!response.IsSuccessStatusCode)
                {
                    // Handle failure from GET call as needed
                    _loggerService.LogError("Failed to generate report", "", nameof(EnvelopeBreakageProcessingController));
                    return StatusCode((int)response.StatusCode, "Failed to get envelope breakages after configuration.");
                }

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
                        worksheet.Cells[row, 3].Value = item.BookletSerialRange;
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
        public async Task<IActionResult> GetEnvelopeBreakingReport(int ProjectId)
        {
            try
            {
                var projectconfig = await _context.ProjectConfigs
                    .FirstOrDefaultAsync(p => p.ProjectId == ProjectId);

                if (projectconfig == null)
                    return NotFound("Project config not found");

                var nrDataDict = await _context.NRDatas
                    .Where(p => p.ProjectId == ProjectId)
                    .ToDictionaryAsync(p => p.Id);

                var fields = await _context.Fields
                    .Where(f => projectconfig.EnvelopeMakingCriteria.Contains(f.FieldId))
                    .ToListAsync();

                var fieldNames = fields
                    .OrderBy(f => projectconfig.EnvelopeMakingCriteria.IndexOf(f.FieldId))
                    .Select(f => f.Name)
                    .ToList();

                IQueryable<EnvelopeBreakingResult> query = _context.EnvelopeBreakingResults
                    .Where(r => r.ProjectId == ProjectId);

                var maxBatch = await _context.EnvelopeBreakingResults
                    .Where(r => r.ProjectId == ProjectId)
                    .MaxAsync(r => (int?)r.UploadBatch);

                if (maxBatch.HasValue)
                    query = query.Where(r => r.UploadBatch == maxBatch.Value);

                var results = await query.OrderBy(r => r.Id).ToListAsync();

                if (!results.Any())
                    return NotFound("No envelope breaking results found");

                var fullData = new List<dynamic>();

                foreach (var result in results)
                {
                    var row = new ExpandoObject();
                    var rowDict = (IDictionary<string, object>)row;

                    bool isMssRow = result.NrDataId == 0 && result.SerialNumber == 0;
                    rowDict["isMss"] = isMssRow;

                    NRData nr = null;
                    if (!isMssRow && result.NrDataId != 0 && nrDataDict.TryGetValue(result.NrDataId, out var nrData))
                    {
                        nr = nrData;
                    }

                    // NR fields
                    if (nr != null)
                    {
                        rowDict["SubjectName"] = nr.SubjectName;
                        rowDict["Pages"] = nr.Pages;
                        rowDict["Symbol"] = nr.Symbol;
                        rowDict["Day"] = nr.Day;
                        rowDict["NRQuantity"] = nr.NRQuantity;
                    }

                    // JSON fields
                    if (nr != null && !string.IsNullOrEmpty(nr.NRDatas))
                    {
                        try
                        {
                            var extraFields = JsonSerializer.Deserialize<Dictionary<string, string>>(nr.NRDatas);
                            if (extraFields != null)
                            {
                                foreach (var kvp in extraFields)
                                {
                                    rowDict[kvp.Key] = kvp.Value;
                                }
                            }
                        }
                        catch { }
                    }

                    // Envelope fields
                    rowDict["SerialNo"] = result.SerialNumber;
                    rowDict["CatchNo"] = result.CatchNo ?? "";
                    rowDict["CenterCode"] = result.CenterCode ?? "";
                    rowDict["CenterSort"] = result.CenterSort;
                    rowDict["ExamTime"] = result.ExamTime ?? "";
                    rowDict["ExamDate"] = result.ExamDate ?? "";
                    rowDict["Quantity"] = result.Quantity;
                    rowDict["NodalCode"] = result.NodalCode ?? "";
                    rowDict["NodalSort"] = result.NodalSort;
                    rowDict["Route"] = result.Route ?? "";
                    rowDict["RouteSort"] = result.RouteSort;
                    rowDict["EnvQuantity"] = result.EnvQuantity;
                    rowDict["CenterEnv"] = result.CenterEnv;
                    rowDict["TotalEnv"] = result.TotalEnv;
                    rowDict["Env"] = result.Env ?? "";
                    rowDict["SerialNumber"] = result.SerialNumber;
                    rowDict["BookletSerial"] = result.BookletSerial ?? "";
                    rowDict["OmrSerial"] = result.OmrSerial ?? "";
                    rowDict["CourseName"] = result.CourseName ?? "";

                    fullData.Add(row);
                }

                // Separate MSS
                var nonMssData = fullData.Where(x =>
                {
                    var d = (IDictionary<string, object>)x;
                    return d.ContainsKey("isMss") && !(bool)d["isMss"];
                }).ToList();

                var mssRowsByCatch = fullData
                    .Where(x =>
                    {
                        var d = (IDictionary<string, object>)x;
                        return d.ContainsKey("isMss") && (bool)d["isMss"];
                    })
                    .GroupBy(x => ((IDictionary<string, object>)x)["CatchNo"]?.ToString())
                    .ToDictionary(g => g.Key, g => g.ToList());

                // Sorting
                IOrderedEnumerable<dynamic> ordered = null;

                foreach (var fieldName in fieldNames)
                {
                    Func<dynamic, object> keySelector = record =>
                    {
                        var dict = (IDictionary<string, object>)record;

                        if (!dict.ContainsKey(fieldName) || dict[fieldName] == null)
                            return "";

                        var val = dict[fieldName];

                        switch (fieldName)
                        {
                            case "RouteSort":
                            case "CenterSort":
                                return int.TryParse(val.ToString(), out int i) ? i : 0;

                            case "NodalSort":
                                return double.TryParse(val.ToString(), out double d) ? d : 0;

                            case "ExamDate":
                                return DateTime.TryParseExact(val.ToString(), "dd-MM-yyyy",
                                    CultureInfo.InvariantCulture,
                                    DateTimeStyles.None, out DateTime dt)
                                    ? dt : DateTime.MinValue;

                            default:
                                return val.ToString().Trim();
                        }
                    };

                    ordered = ordered == null
                        ? nonMssData.OrderBy(keySelector)
                        : ordered.ThenBy(keySelector);
                }

                var sortedNonMss = ordered?.ToList() ?? nonMssData;

                // Reinsert MSS
                string mssMode = projectconfig.MssAttached?.ToLower();
                var finalSortedList = new List<dynamic>();

                string lastCatch = null;
                var buffer = new List<dynamic>();

                foreach (var item in sortedNonMss)
                {
                    var dict = (IDictionary<string, object>)item;
                    string catchNo = dict["CatchNo"]?.ToString();

                    if (catchNo != lastCatch && lastCatch != null)
                    {
                        finalSortedList.AddRange(buffer);

                        if (mssMode == "end" && mssRowsByCatch.ContainsKey(lastCatch))
                            finalSortedList.AddRange(mssRowsByCatch[lastCatch]);

                        buffer.Clear();
                    }

                    if (mssMode == "start" && lastCatch == null && mssRowsByCatch.ContainsKey(catchNo))
                    {
                        finalSortedList.AddRange(mssRowsByCatch[catchNo]);
                    }

                    buffer.Add(item);
                    lastCatch = catchNo;
                }

                if (buffer.Count > 0)
                {
                    finalSortedList.AddRange(buffer);

                    if (mssMode == "end" && lastCatch != null && mssRowsByCatch.ContainsKey(lastCatch))
                        finalSortedList.AddRange(mssRowsByCatch[lastCatch]);
                }

                // Columns
                var allKeys = finalSortedList
                    .SelectMany(x => ((IDictionary<string, object>)x).Keys)
                    .Where(k => k != "isMss")
                    .Distinct()
                    .ToList();

                // SAFETY FIX
                if (!finalSortedList.Any() || !allKeys.Any())
                {
                    return BadRequest("No data available to generate Excel.");
                }

                // Excel
                var reportPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", ProjectId.ToString());
                Directory.CreateDirectory(reportPath);

                var filePath = Path.Combine(reportPath, "EnvelopeBreaking.xlsx");
                if (System.IO.File.Exists(filePath))
                    System.IO.File.Delete(filePath);

                using (var package = new ExcelPackage())
                {
                    var ws = package.Workbook.Worksheets.Add("Envelope Report");

                    // Headers
                    for (int i = 0; i < allKeys.Count; i++)
                    {
                        ws.Cells[1, i + 1].Value = allKeys[i];
                        ws.Cells[1, i + 1].Style.Font.Bold = true;
                    }

                    // Data
                    int rowIdx = 2;

                    foreach (var item in finalSortedList)
                    {
                        var dict = (IDictionary<string, object>)item;

                        for (int col = 0; col < allKeys.Count; col++)
                        {
                            var key = allKeys[col];
                            ws.Cells[rowIdx, col + 1].Value =
                                dict.ContainsKey(key) ? dict[key] : "";
                        }

                        rowIdx++;
                    }

                    // ✅ CRITICAL FIX
                    if (ws.Dimension != null)
                    {
                        ws.Cells[ws.Dimension.Address].AutoFitColumns();
                    }

                    ws.View.FreezePanes(2, 1);

                    package.SaveAs(new FileInfo(filePath));
                }

                return Ok(new
                {
                    message = "Report generated successfully",
                    filePath,
                    recordsCount = finalSortedList.Count
                });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new
                {
                    error = ex.Message,
                    stack = ex.StackTrace // helpful for debugging
                });
            }
        }
    }
    }
