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
using Tools.Middleware;

namespace Tools.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class EnvelopeBreakageProcessingController : ControllerBase
    {
        private readonly ERPToolsDbContext _context;
        private readonly ILoggerService _loggerService;
        private readonly ApiSettings _apiSettings;
        private readonly IDispatchService _dispatchService;

        public EnvelopeBreakageProcessingController(ERPToolsDbContext context, ILoggerService loggerService, IOptions<ApiSettings> apiSettings, IDispatchService dispatchService)
        {
            _context = context;
            _loggerService = loggerService;
            _apiSettings = apiSettings.Value;
            _dispatchService = dispatchService;
        }


        [HttpPost("ProcessEnvelopeBreaking")]
        public async Task<IActionResult> ProcessEnvelopeBreaking(int ProjectId, int triggeredBy = 0, [FromQuery] bool skipReset = false, [FromQuery] int? lotNo = null, [FromQuery] string? catchNo = null, [FromQuery] bool bypassDispatch = false, [FromQuery] int? batchNo = null)
        {
            try
            {
                var isNewModel = await _context.NrData1.AnyAsync(p => p.ProjectId == ProjectId);
                if (isNewModel)
                {
                    return await ProcessEnvelopeBreakingNew(ProjectId, triggeredBy, skipReset, lotNo, catchNo, bypassDispatch, batchNo);
                }

                if (!skipReset) await ResetReportStatus(ProjectId);

                // ✅ STEP 0: Validate dispatch status (mandatory backend validation unless bypassed)
                // Allow a global bypass via environment variable `BYPASS_DISPATCH_CHECK=true`
                // Temporarily ignore dispatch date checks (force bypass for now)
                var globalBypass = (Environment.GetEnvironmentVariable("BYPASS_DISPATCH_CHECK") ?? "").ToLowerInvariant() == "true";
                var dispatchBypassEffective = true;

                if (!dispatchBypassEffective)
                {
                    List<int> lotsToCheck = new List<int>();
                    if (lotNo.HasValue)
                    {
                        lotsToCheck.Add(lotNo.Value);
                    }
                    else if (!string.IsNullOrEmpty(catchNo))
                    {
                        var lot = await _context.NRDatas
                            .Where(p => p.ProjectId == ProjectId && p.CatchNo == catchNo && p.Batch == (batchNo ?? 1))
                            .Select(p => p.LotNo)
                            .FirstOrDefaultAsync();
                        if (lot != 0) lotsToCheck.Add(lot);
                    }
                    else
                    {
                        lotsToCheck = await _context.NRDatas
                            .Where(p => p.ProjectId == ProjectId && p.Status == true && p.Batch == (batchNo ?? 1))
                            .Select(p => p.LotNo)
                            .Distinct()
                            .ToListAsync();
                    }

                    if (lotsToCheck.Any())
                    {
                        var dispatchInfoDict = await _dispatchService.GetDispatchDatesAsync(ProjectId, lotsToCheck);
                        var dispatchedLots = dispatchInfoDict.Where(d => d.Value.IsDispatched).ToList();

                        if (dispatchedLots.Any())
                        {
                            var dispatchedLotNumbers = string.Join(", ", dispatchedLots.Select(d => d.Key));
                            //return BadRequest(new
                            //{
                            //    error = $"Lot(s) {dispatchedLotNumbers} already dispatched. Envelope Breaking not allowed.",
                            //    dispatchedLots = dispatchedLots.Select(d => new
                            //    {
                            //        lotNo = d.Key,
                            //        dispatchDate = d.Value.DispatchDate
                            //    })
                            //});
                        }
                    }
                }
                var envCaps = await _context.EnvelopesTypes
                    .Select(e => new { e.EnvelopeName, e.Capacity })
                    .ToListAsync();

                var eligibleSteps = Tools.Models.PipelineNavigator.GetEligiblePickupSteps(Tools.Models.PipelineNavigator.STEP_AWAITING_ENV);

                var nrQuery = _context.NRDatas
                    .Where(p => p.ProjectId == ProjectId && p.Status == true && eligibleSteps.Contains(p.Steps) && p.Batch == (batchNo ?? 1));

                if (lotNo.HasValue) nrQuery = nrQuery.Where(p => p.LotNo == lotNo.Value);
                if (!string.IsNullOrEmpty(catchNo)) nrQuery = nrQuery.Where(p => p.CatchNo == catchNo);

                var nrData = await nrQuery
                    .OrderBy(p => p.CatchNo)
                    .ThenBy(p => p.RouteSort)
                    .ThenBy(p => p.NodalSort)
                    .ThenBy(p => p.CenterSort)
                    .ToListAsync();

                var catchesToProcess = nrData.Select(n => n.CatchNo).Where(c => !string.IsNullOrEmpty(c)).Distinct().ToList();
                var nrDataIdsToProcess = nrData.Select(n => n.Id).ToList();



                var envBreaking = await _context.EnvelopeBreakages
                    .Where(p => p.ProjectId == ProjectId && (p.Status == true))
                    .ToListAsync();
                
                if (!envBreaking.Any())
                    return BadRequest("No Envelope Breakdown configuration found. Please run the Envelope Breakages (Inner/Outer) configuration first.");

                if (!nrData.Any())
                    return BadRequest("No valid NR Data found for Envelope Breaking processing.");

                var extrasQuery = _context.ExtrasEnvelope.Where(p => p.ProjectId == ProjectId && p.Status == 1);
                if (!string.IsNullOrEmpty(catchNo)) extrasQuery = extrasQuery.Where(e => e.CatchNo == catchNo);
                // Note: ExtraEnvelopes doesn't usually have LotNo, so we filter by catchNo if provided.
                
                var extras = await extrasQuery.ToListAsync();

                var projectconfig = await _context.ProjectConfigs
                    .Where(p => p.ProjectId == ProjectId)
                    .FirstOrDefaultAsync();

                if (projectconfig == null)
                    return NotFound("Project config not found");

                var envelopeCapacities = envCaps.ToDictionary(x => x.EnvelopeName, x => x.Capacity);
                var envDict = envBreaking.ToDictionary(
                    e => e.NrDataId,
                    e => new { e.TotalEnvelope, e.OuterEnvelope, e.InnerEnvelope }
                );

                var project = await _context.Projects.FirstOrDefaultAsync(p => p.ProjectId == ProjectId);
                int projectTypeId = project?.TypeId ?? 0;

                var mssTypes = projectconfig.MssTypes ?? new List<int>();
                string mssMode = projectconfig.MssAttached?.ToLower() ?? "end";
                var mssData = new List<Mss>();

                if (!string.Equals(mssMode, "none", StringComparison.OrdinalIgnoreCase))
                {
                    var mssQuery = _context.Mss.AsQueryable();

                    if (projectTypeId > 0)
                    {
                        mssQuery = mssQuery.Where(m => m.TypeId == projectTypeId);
                    }

                    if (mssTypes.Any())
                    {
                        mssQuery = mssQuery.Where(m => mssTypes.Contains(m.Id));
                    }

                    mssData = await mssQuery.ToListAsync();

                    if (!mssData.Any() && projectTypeId > 0)
                    {
                        mssData = await _context.Mss.Where(m => m.TypeId == projectTypeId).ToListAsync();
                    }
                }

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

                var nodalExtraConfig = extrasconfig.FirstOrDefault(e => e.ExtraType == 1);
                bool attachExtraForAllNodal = nodalExtraConfig?.AttachExtraForEachCatchForAllNodal == true ||
                                              extrasconfig.Any(e => e.AttachExtraForEachCatchForAllNodal);

                var allProjectNodals = await _context.NRDatas
                    .Where(n => n.ProjectId == ProjectId && n.Status == true && !string.IsNullOrEmpty(n.NodalCode))
                    .GroupBy(n => n.NodalCode)
                    .Select(g => new {
                        NodalCode = g.Key,
                        NodalSort = g.Max(n => n.NodalSort),
                        Route = g.Select(n => n.Route).FirstOrDefault(r => !string.IsNullOrEmpty(r)) ?? "",
                        RouteSort = g.Max(n => n.RouteSort),
                        District = g.Select(n => n.District).FirstOrDefault(d => !string.IsNullOrEmpty(d)) ?? "",
                        DistrictSort = g.Max(n => n.DistrictSort),
                        CenterCode = g.Select(n => n.CenterCode).FirstOrDefault() ?? "",
                        NrDataId = g.Select(n => n.Id).FirstOrDefault()
                    })
                    .OrderBy(n => n.RouteSort)
                    .ThenBy(n => n.NodalSort)
                    .ToListAsync();

                // ✅ Load sorting field names EARLY so they're available inside helpers
                var sortFields = await _context.Fields
                    .Where(f => projectconfig.EnvelopeMakingCriteria.Contains(f.FieldId))
                    .ToListAsync();

                var sortingFieldNames = sortFields
                    .OrderBy(f => projectconfig.EnvelopeMakingCriteria.IndexOf(f.FieldId))
                    .Select(f => f.Name)
                    .ToList();

                // ✅ Build a lookup: CatchNo -> parsed NRDatas JSON dictionary
                // This avoids re-parsing JSON on every row iteration
                var nrDataJsonLookup = nrData.ToDictionary(
                    nr => nr.Id,
                    nr =>
                    {
                        if (string.IsNullOrWhiteSpace(nr.NRDatas)) return null;
                        try { return JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(nr.NRDatas); }
                        catch { return null; }
                    }
                );

                // ✅ Helper: get a field value from parsed NRDatas JSON (case-insensitive)
                string GetNrField(Dictionary<string, JsonElement> nrDynamic, string fieldName)
                {
                    if (nrDynamic == null) return "";
                    var match = nrDynamic.FirstOrDefault(k =>
                        k.Key.Equals(fieldName, StringComparison.OrdinalIgnoreCase));
                    return match.Key != null ? (match.Value.GetString() ?? "") : "";
                }

                // ✅ Helper: fill any sorting fields missing from a row dict using NRDatas JSON
                void FillDynamicFields(IDictionary<string, object> rowDict, int nrDataId)
                {
                    if (!nrDataJsonLookup.TryGetValue(nrDataId, out var nrDynamic) || nrDynamic == null)
                        return;

                    foreach (var fieldName in sortingFieldNames)
                    {
                        if (!rowDict.ContainsKey(fieldName))
                        {
                            rowDict[fieldName] = GetNrField(nrDynamic, fieldName);
                        }
                    }
                }

                string GetPackingDenominationForQuantity(string innerEnvelopeJson, int targetQty, ref List<int> innerEnvelopesPool)
                {
                    if (innerEnvelopesPool == null)
                    {
                        innerEnvelopesPool = new List<int>();
                        if (!string.IsNullOrWhiteSpace(innerEnvelopeJson))
                        {
                            try
                            {
                                var dict = JsonSerializer.Deserialize<Dictionary<string, string>>(innerEnvelopeJson);
                                foreach (var kvp in dict)
                                {
                                    string envName = kvp.Key.Replace("E", "");
                                    if (int.TryParse(envName, out int capacity) && int.TryParse(kvp.Value, out int count))
                                    {
                                        for (int i = 0; i < count; i++)
                                            innerEnvelopesPool.Add(capacity);
                                    }
                                }
                            }
                            catch { }
                        }
                        innerEnvelopesPool = innerEnvelopesPool.OrderByDescending(x => x).ToList();
                    }

                    if (targetQty <= 0 || innerEnvelopesPool.Count == 0) return null;

                    var used = new Dictionary<int, int>();
                    int currentSum = 0;
                    var remainingPool = new List<int>();

                    foreach (var cap in innerEnvelopesPool)
                    {
                        if (currentSum < targetQty && currentSum + cap <= targetQty)
                        {
                            currentSum += cap;
                            if (!used.ContainsKey(cap)) used[cap] = 0;
                            used[cap]++;
                        }
                        else
                        {
                            remainingPool.Add(cap);
                        }
                    }

                    innerEnvelopesPool = remainingPool;
                    if (used.Count == 0) return null;
                    return "[" + string.Join(" & ", used.OrderByDescending(x => x.Key).Select(kvp => $"{kvp.Key}x{kvp.Value}")) + "]";
                }

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
                        mssDict["District"] = "";
                        mssDict["DistrictSort"] = 0;
                        mssDict["NrDataId"] = 0;
                        mssDict["ExtraId"] = (int?)null;
                        rows.Add(mssRow);
                    }
                    return rows;
                }

                void AddExtraWithEnv(ExtraEnvelopes extra, string examDate, string examTime, string course, int NrQuantity,
                    string NodalCode, string CenterCode, double CenterSort, double NodalSort, int RouteSort, string Route, int nrDataId, string District, int DistrictSort)
                {
                    var extraConfig = extrasconfig.FirstOrDefault(e => e.ExtraType == extra.ExtraId);
                    int envCapacity = 0;
                    string extraOuterTypeStr = null;

                    if (extraConfig != null && !string.IsNullOrEmpty(extraConfig.EnvelopeType))
                    {
                        var envType = JsonSerializer.Deserialize<Dictionary<string, string>>(extraConfig.EnvelopeType);
                        if (envType != null)
                        {
                            if (envType.TryGetValue("Outer", out string outerType))
                            {
                                if (envelopeCapacities.TryGetValue(outerType, out int cap))
                                {
                                    envCapacity = cap;
                                }
                                extraOuterTypeStr = outerType.Replace("E", "");
                            }
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

                        string extraPackingDenom = !string.IsNullOrEmpty(extraOuterTypeStr) ? $"[{extraOuterTypeStr}x1]" : null;
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

                        int modifiedDistrictSort = extra.ExtraId switch
                        {
                            1 => DistrictSort,
                            2 => 10000,
                            3 => 100000,
                            _ => DistrictSort
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
                        extraDict["PackingDenomination"] = extraPackingDenom;
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
                        extraDict["NodalCode"] = extra.ExtraId == 1 ? NodalCode : "-";
                        extraDict["Route"] = extra.ExtraId == 1 ? Route : "-";
                        extraDict["District"] = extra.ExtraId == 1 ? District : "-";
                        extraDict["NodalSort"] = modifiedNodalSort;
                        extraDict["CenterSort"] = modifiedCenterSort;
                        extraDict["RouteSort"] = modifiedRouteSort;
                        extraDict["DistrictSort"] = modifiedDistrictSort;
                        // ✅ Fill dynamic NRDatas fields for extra rows too
                        FillDynamicFields(extraDict, nrDataId);

                        resultList.Add(extraRow);
                    }
                }

                void AddMissingNodalExtrasForCatch(string targetCatchNo, NRData referenceNrData)
                {
                    if (!attachExtraForAllNodal || string.IsNullOrEmpty(targetCatchNo) || referenceNrData == null) return;

                    foreach (var nodal in allProjectNodals)
                    {
                        if (nodalExtrasAddedForNodalCatch.Contains((nodal.NodalCode, targetCatchNo)))
                            continue;

                        var extrasToAdd = extras.Where(e => e.ExtraId == 1 && e.CatchNo == targetCatchNo &&
                            (e.NodalCode == nodal.NodalCode || (string.IsNullOrEmpty(e.NodalCode) && !extras.Any(x => x.ExtraId == 1 && x.CatchNo == targetCatchNo && x.NodalCode == nodal.NodalCode)))).ToList();

                        if (extrasToAdd.Any())
                        {
                            foreach (var extra in extrasToAdd)
                            {
                                AddExtraWithEnv(extra, referenceNrData.ExamDate, referenceNrData.ExamTime, referenceNrData.CourseName,
                                    0, nodal.NodalCode, nodal.CenterCode, 10000,
                                    nodal.NodalSort, nodal.RouteSort, nodal.Route, nodal.NrDataId, nodal.District, nodal.DistrictSort);
                            }
                        }
                        else
                        {
                            var sampleExtra = extras.FirstOrDefault(e => e.ExtraId == 1 && e.CatchNo == targetCatchNo)
                                           ?? extras.FirstOrDefault(e => e.ExtraId == 1);

                            int calculatedQty = 0;
                            if (nodalExtraConfig != null)
                            {
                                if (!string.IsNullOrWhiteSpace(nodalExtraConfig.nodalValue))
                                {
                                    try
                                    {
                                        var nConfigs = JsonSerializer.Deserialize<List<ExtraEnvelopesController.NodalValueConfig>>(
                                            nodalExtraConfig.nodalValue,
                                            new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
                                        var match = nConfigs?.FirstOrDefault(nc =>
                                            !string.IsNullOrEmpty(nc.NodalCodes) &&
                                            nc.NodalCodes.Split(',').Select(x => x.Trim()).Contains(nodal.NodalCode, StringComparer.OrdinalIgnoreCase));
                                        if (match != null && int.TryParse(match.Value, out int mv))
                                        {
                                            calculatedQty = mv;
                                        }
                                    }
                                    catch { }
                                }

                                if (calculatedQty == 0 && int.TryParse(nodalExtraConfig.Value, out int fv))
                                {
                                    calculatedQty = fv;
                                }
                            }

                            if (calculatedQty == 0 && sampleExtra != null)
                            {
                                calculatedQty = sampleExtra.Quantity;
                            }

                            if (calculatedQty > 0 || sampleExtra != null)
                            {
                                var fallbackExtra = new ExtraEnvelopes
                                {
                                    ProjectId = ProjectId,
                                    CatchNo = targetCatchNo,
                                    NodalCode = nodal.NodalCode,
                                    ExtraId = 1,
                                    Quantity = calculatedQty > 0 ? calculatedQty : (sampleExtra?.Quantity ?? 0),
                                    InnerEnvelope = sampleExtra?.InnerEnvelope ?? "0",
                                    OuterEnvelope = sampleExtra?.OuterEnvelope ?? "0",
                                    Status = 1
                                };

                                AddExtraWithEnv(fallbackExtra, referenceNrData.ExamDate, referenceNrData.ExamTime, referenceNrData.CourseName,
                                    0, nodal.NodalCode, nodal.CenterCode, 10000,
                                    nodal.NodalSort, nodal.RouteSort, nodal.Route, nodal.NrDataId, nodal.District, nodal.DistrictSort);
                            }
                        }

                        nodalExtrasAddedForNodalCatch.Add((nodal.NodalCode, targetCatchNo));
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
                            var extrasToAdd = extras.Where(e => e.ExtraId == 1 && e.CatchNo == prevCatchNo &&
                                (e.NodalCode == prevNrData.NodalCode || (string.IsNullOrEmpty(e.NodalCode) && !extras.Any(x => x.ExtraId == 1 && x.CatchNo == prevCatchNo && x.NodalCode == prevNrData.NodalCode)))).ToList();
                            foreach (var extra in extrasToAdd)
                            {
                                AddExtraWithEnv(extra, prevNrData.ExamDate, prevNrData.ExamTime, prevNrData.CourseName,
                                    prevNrData.NRQuantity, prevNrData.NodalCode, prevNrData.CenterCode, prevNrData.CenterSort,
                                    prevNrData.NodalSort, prevNrData.RouteSort, prevNrData.Route, prevNrData.Id, prevNrData.District, prevNrData.DistrictSort);
                            }
                            nodalExtrasAddedForNodalCatch.Add((prevNrData.NodalCode, prevCatchNo));
                        }

                        if (attachExtraForAllNodal)
                        {
                            AddMissingNodalExtrasForCatch(prevCatchNo, prevNrData);
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
                                        prevNrData.NodalSort, prevNrData.RouteSort, prevNrData.Route, prevNrData.Id, prevNrData.District, prevNrData.DistrictSort);
                                }
                                catchExtrasAdded.Add((extraId, prevCatchNo));
                            }
                        }
                    }

                    if (!catchNoChanged && prevNodalCode != null && current.NodalCode != prevNodalCode)
                    {
                        if (!nodalExtrasAddedForNodalCatch.Contains((prevNodalCode, current.CatchNo)))
                        {
                            var extrasToAdd = extras.Where(e => e.ExtraId == 1 && e.CatchNo == current.CatchNo &&
                                (e.NodalCode == prevNodalCode || (string.IsNullOrEmpty(e.NodalCode) && !extras.Any(x => x.ExtraId == 1 && x.CatchNo == current.CatchNo && x.NodalCode == prevNodalCode)))).ToList();
                            foreach (var extra in extrasToAdd)
                            {
                                AddExtraWithEnv(extra, current.ExamDate, current.ExamTime, current.CourseName,
                                    current.NRQuantity, prevNodalCode, current.CenterCode, current.CenterSort,
                                    prevNodalSort, prevRouteSort, prevRoute, current.Id, current.District, current.DistrictSort);
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
                        List<int> innerPool = null;
                        string innerJson = envDict.TryGetValue(current.Id, out var eInfo) ? eInfo.InnerEnvelope : null;
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
                                nrDict["District"] = current.District;
                                nrDict["DistrictSort"] = current.DistrictSort;
                                nrDict["NrDataId"] = current.Id;
                                nrDict["ExtraId"] = (int?)null;
                                
                                nrDict["PackingDenomination"] = GetPackingDenominationForQuantity(innerJson, envQty, ref innerPool);

                                // ✅ Fill any sorting fields from NRDatas JSON
                                FillDynamicFields(nrDict, current.Id);

                                resultList.Add(nrRow);

                                remainingQty -= envQty;
                                envelopeIndex++;
                                if (remainingQty <= 0) break;
                            }
                            if (remainingQty <= 0) break;
                        }
                    }
                    else
                    {
                        centerEnvCounter++;
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
                        nrDict["EnvQuantity"] = current.Quantity;
                        nrDict["NodalCode"] = current.NodalCode;
                        nrDict["CenterEnv"] = centerEnvCounter;
                        nrDict["TotalEnv"] = 1;
                        nrDict["Env"] = "1/1";
                        nrDict["NRQuantity"] = current.NRQuantity;
                        nrDict["CenterSort"] = current.CenterSort;
                        nrDict["NodalSort"] = current.NodalSort;
                        nrDict["Route"] = current.Route;
                        nrDict["RouteSort"] = current.RouteSort;
                        nrDict["District"] = current.District;
                        nrDict["DistrictSort"] = current.DistrictSort;
                        nrDict["NrDataId"] = current.Id;
                        nrDict["ExtraId"] = (int?)null;

                        List<int> innerPool = null;
                        string innerJson = envDict.TryGetValue(current.Id, out var eInfo2) ? eInfo2.InnerEnvelope : null;
                        nrDict["PackingDenomination"] = GetPackingDenominationForQuantity(innerJson, current.Quantity, ref innerPool);

                        // ✅ Fill any sorting fields from NRDatas JSON
                        FillDynamicFields(nrDict, current.Id);

                        resultList.Add(nrRow);
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
                            var extrasToAdd = extras.Where(e => e.ExtraId == 1 && e.CatchNo == prevCatchNo &&
                                (e.NodalCode == lastNrData.NodalCode || (string.IsNullOrEmpty(e.NodalCode) && !extras.Any(x => x.ExtraId == 1 && x.CatchNo == prevCatchNo && x.NodalCode == lastNrData.NodalCode)))).ToList();
                            foreach (var extra in extrasToAdd)
                            {
                                AddExtraWithEnv(extra, lastNrData.ExamDate, lastNrData.ExamTime, lastNrData.CourseName,
                                    lastNrData.NRQuantity, lastNrData.NodalCode, lastNrData.CenterCode, lastNrData.CenterSort,
                                    lastNrData.NodalSort, lastNrData.RouteSort, lastNrData.Route, lastNrData.Id, lastNrData.District, lastNrData.DistrictSort);
                            }
                            nodalExtrasAddedForNodalCatch.Add((lastNrData.NodalCode, prevCatchNo));
                        }

                        if (attachExtraForAllNodal)
                        {
                            AddMissingNodalExtrasForCatch(prevCatchNo, lastNrData);
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
                                        lastNrData.NodalSort, lastNrData.RouteSort, lastNrData.Route, lastNrData.Id, lastNrData.District, lastNrData.DistrictSort);
                                }
                            }
                        }
                    }
                }

              

                // sortFields and sortingFieldNames already loaded above ✅

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

                        // Explicitly defined fields with specific types
                        if (fieldName.Equals("NodalSort", StringComparison.OrdinalIgnoreCase))
                            return double.TryParse(val.ToString(), out double n) ? n : 0.0;

                        if (fieldName.Equals("ExamDate", StringComparison.OrdinalIgnoreCase))
                            return DateTime.TryParseExact(val.ToString(), "dd-MM-yyyy",
                                CultureInfo.InvariantCulture, DateTimeStyles.None, out var d)
                                ? (object)d : DateTime.MinValue;

                        // ✅ Any field with "sort" in its name (covers RouteSort, CenterSort, 
                        // District Sort, and any other dynamic sort field) → parse as double
                        if (fieldName.Contains("sort", StringComparison.OrdinalIgnoreCase))
                            return double.TryParse(val.ToString(), out double d) ? d : 0.0;

                        // All other fields → string comparison
                        return val.ToString().Trim().ToLowerInvariant();
                    };

                    sortedResultList = sortedResultList == null
                        ? nonMssRows.OrderBy(keySelector)
                        : sortedResultList.ThenBy(keySelector);
                }

                var sortedNonMss = sortedResultList?.Cast<dynamic>().ToList() ?? nonMssRows;

                // STEP 2: Re-insert MSS rows at correct positions after sorting
                var finalResultList = new List<dynamic>();
                string lastCatchForMss = null;
                var buffer = new List<dynamic>();

                foreach (var item in sortedNonMss)
                {
                    var itemDict = (IDictionary<string, object>)item;
                    string cNo = itemDict["CatchNo"]?.ToString();

                    if (cNo != lastCatchForMss && lastCatchForMss != null)
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
                                cNo,
                                itemDict["ExamDate"]?.ToString(),
                                itemDict["ExamTime"]?.ToString(),
                                itemDict["CourseName"]?.ToString()
                            ));
                        }
                        buffer.Clear();
                    }

                    if (mssMode == "start" && lastCatchForMss == null)
                    {
                        finalResultList.AddRange(CreateMssRows(
                            cNo,
                            itemDict["ExamDate"]?.ToString(),
                            itemDict["ExamTime"]?.ToString(),
                            itemDict["CourseName"]?.ToString()
                        ));
                    }

                    buffer.Add(item);
                    lastCatchForMss = cNo;
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

                // STEP 3: Assign serials
                int bookletStart = projectconfig?.BookletSerialNumber ?? 0;
                int omrStart = projectconfig.OmrSerialNumber;
                bool resetOmrSerialOnCatchChange = projectconfig.ResetOmrSerialOnCatchChange;
                bool resetBookletSerialOnCatchChange = projectconfig?.ResetBookletSerialOnCatchChange ?? false;
                int serial = 1;

                // ✅ Continuity: if not resetting on catch change and we are doing a targeted run, 
                // find the last used serials in the project to continue the sequence.
                if (!string.IsNullOrEmpty(catchNo) && !resetOmrSerialOnCatchChange)
                {
                    var lastResult = await _context.EnvelopeBreakingResults
                        .Where(r => r.ProjectId == ProjectId && r.CatchNo != catchNo && r.Status)
                        .OrderByDescending(r => r.SerialNumber)
                        .FirstOrDefaultAsync();

                    if (lastResult != null)
                    {
                        serial = lastResult.SerialNumber + 1;

                        if (!string.IsNullOrEmpty(lastResult.OmrSerial))
                        {
                            var parts = lastResult.OmrSerial.Split('-');
                            if (parts.Length == 2 && int.TryParse(parts[1], out int lastOmr)) omrStart = lastOmr + 1;
                            else if (int.TryParse(lastResult.OmrSerial, out int lastOmrSing)) omrStart = lastOmrSing + 1;
                        }

                        if (!string.IsNullOrEmpty(lastResult.BookletSerial))
                        {
                            var parts = lastResult.BookletSerial.Split('-');
                            if (parts.Length == 2 && int.TryParse(parts[1], out int lastBooklet)) bookletStart = lastBooklet + 1;
                            else if (int.TryParse(lastResult.BookletSerial, out int lastBookletSing)) bookletStart = lastBookletSing + 1;
                        }
                    }
                }

                bool assignBookletSerial = bookletStart > 0;
                bool assignOmrSerial = omrStart > 0;
                string prevCatchForSerial = null;

                var envelopeResults = new List<EnvelopeBreakingResult>();

                foreach (var item in finalResultList)
                {
                    var dict = (IDictionary<string, object>)item;
                    bool isMssRow = dict.ContainsKey("isMss") && (bool)dict["isMss"];
                    bool isExtra = (bool)dict["ExtraAttached"];
                    string cNo = dict["CatchNo"]?.ToString();

                    if (!isMssRow && prevCatchForSerial != null && cNo != prevCatchForSerial)
                    {
                        serial = 1;
                        if (resetOmrSerialOnCatchChange) omrStart = projectconfig.OmrSerialNumber;
                        if (resetBookletSerialOnCatchChange) bookletStart = projectconfig?.BookletSerialNumber ?? 0;
                    }

                    if (isMssRow)
                    {
                        envelopeResults.Add(new EnvelopeBreakingResult
                        {
                            ProjectId = ProjectId,
                            NrDataId = 0,
                            ExtraId = null,
                            CatchNo = cNo,
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
                            DistrictSort = 0,
                            CourseName = dict["CourseName"]?.ToString(),
                            PackingDenomination = null,
                        });
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

                    double centerSort = 0;
                    if (dict["CenterSort"] is double dCenterSort) centerSort = dCenterSort;
                    else if (dict["CenterSort"] is int iCenterSort) centerSort = iCenterSort;
                    else double.TryParse(dict["CenterSort"]?.ToString(), out centerSort);

                    int routeSort = 0;
                    if (dict["RouteSort"] is double dRouteSort) routeSort = (int)dRouteSort;
                    else if (dict["RouteSort"] is int iRouteSort) routeSort = iRouteSort;
                    else int.TryParse(dict["RouteSort"]?.ToString(), out routeSort);

                    int districtSort = 0;
                    if (dict["DistrictSort"] is double dSort) districtSort = (int)dSort;
                    else if (dict["DistrictSort"] is int idSort) districtSort = idSort;
                    else int.TryParse(dict["DistrictSort"]?.ToString(), out districtSort);

                    double nodalSort = 0.0;
                    if (dict["NodalSort"] is double dNodalSort) nodalSort = dNodalSort;
                    else if (dict["NodalSort"] is int iNodalSort) nodalSort = (double)iNodalSort;
                    else double.TryParse(dict["NodalSort"]?.ToString(), out nodalSort);

                    envelopeResults.Add(new EnvelopeBreakingResult
                    {
                        ProjectId = ProjectId,
                        NrDataId = nrDataId ?? 0,
                        ExtraId = extraId,
                        CatchNo = cNo,
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
                        District = dict["District"]?.ToString(),
                        DistrictSort = districtSort,
                        PackingDenomination = dict.ContainsKey("PackingDenomination") ? dict["PackingDenomination"]?.ToString() : null,
                    });

                    prevCatchForSerial = cNo;
                }

                // Soft deactivate existing records for the catches being processed
                if (catchesToProcess.Any())
                {
                    var existingEnvelopeBreakingIds = await _context.EnvelopeBreakingResults
                        .Where(r => r.ProjectId == ProjectId && r.CatchNo != null && catchesToProcess.Contains(r.CatchNo) && r.Status)
                        .Select(r => r.Id)
                        .ToListAsync();

                    if (existingEnvelopeBreakingIds.Any())
                    {
                        await _context.EnvelopeBreakingResults
                            .Where(r => existingEnvelopeBreakingIds.Contains(r.Id))
                            .ExecuteUpdateAsync(s => s.SetProperty(r => r.Status, false));

                        await _context.BoxBreakingResults
                            .Where(b => b.ProjectId == ProjectId && b.EnvelopeBreakingResultId.HasValue && existingEnvelopeBreakingIds.Contains(b.EnvelopeBreakingResultId.Value) && b.Status)
                            .ExecuteUpdateAsync(s => s.SetProperty(b => b.Status, false));
                    }
                }

                _context.EnvelopeBreakingResults.AddRange(envelopeResults);
                foreach (var nr in nrData)
                   nr.Steps = Tools.Models.PipelineNavigator.STEP_AWAITING_ENV; ;

                await _context.SaveChangesAsync();

                await _loggerService.LogEventAsync(
                    $"Saved {envelopeResults.Count} envelope breaking results for ProjectId {ProjectId}",
                    "EnvelopeBreakageProcessing",
                    triggeredBy,
                    ProjectId);

                // Call the report generation method directly instead of using HttpClient
                try
                {
                    var reportResponse = await GetEnvelopeBreakingReport(ProjectId, lotNo);
                    var isSuccess = true;
                    int statusCode = 500;
                    string exactError = "Failed to get envelope breakages after configuration.";

                    if (reportResponse is ObjectResult objResult && objResult.StatusCode >= 400)
                    {
                        isSuccess = false;
                        statusCode = objResult.StatusCode ?? 500;
                        exactError = System.Text.Json.JsonSerializer.Serialize(objResult.Value);
                    }
                    else if (reportResponse is StatusCodeResult statusResult && statusResult.StatusCode >= 400)
                    {
                        isSuccess = false;
                        statusCode = statusResult.StatusCode;
                    }

                    if (!isSuccess)
                    {
                        await _loggerService.LogErrorAsync("Error during Excel report generation", exactError, nameof(EnvelopeBreakageProcessingController));
                        return StatusCode(statusCode, new { message = "Failed to get envelope breakages after configuration.", exactError = exactError });
                    }
                }
                catch (Exception ex)
                {
                    var innerEx = ex.InnerException?.Message ?? "No inner exception";
                    var fullError = $"{ex.Message} | Inner: {innerEx}";
                    await _loggerService.LogErrorAsync("Error during Excel report generation", fullError, nameof(EnvelopeBreakageProcessingController));
                    return StatusCode(500, new { message = "Failed to get envelope breakages after configuration.", exactError = fullError, stackTrace = ex.StackTrace });
                }

                return Ok(new
                {
                    message = "Envelope breaking data saved to database",
                    recordsCount = envelopeResults.Count,
                });
            }
            catch (Exception ex)
            {
                var innerException = ex.InnerException?.Message ?? "No inner exception";
                var fullError = $"{ex.Message} | Inner: {innerException}";
                await _loggerService.LogErrorAsync("Error processing envelope breaking", fullError, nameof(EnvelopeBreakageProcessingController));
                return StatusCode(500, new { 
                    error = ex.Message,
                    innerException = innerException,
                    stackTrace = ex.StackTrace
                });
            }
        }


        [HttpGet("GetEnvelopeBreakingReport")]
        public async Task<IActionResult> GetEnvelopeBreakingReport(int ProjectId, [FromQuery] int? lotNo = null)
        {
            try
            {
                var isNewModel = await _context.NrData1.AnyAsync(p => p.ProjectId == ProjectId);
                if (isNewModel)
                {
                    return await GetEnvelopeBreakingReportNew(ProjectId, lotNo);
                }

                var projectconfig = await _context.ProjectConfigs
                    .FirstOrDefaultAsync(p => p.ProjectId == ProjectId);

                if (projectconfig == null)
                    return NotFound("Project config not found");

                var nrDataDict = await _context.NRDatas
                    .Where(p => p.ProjectId == ProjectId && p.Status==true)
                    .ToDictionaryAsync(p => p.Id);

                // Catch-wise NR mapping (for MSS rows)
                var nrDataByCatch = await _context.NRDatas
                    .Where(p => p.ProjectId == ProjectId && p.Status==true)
                    .GroupBy(p => p.CatchNo)
                    .ToDictionaryAsync(g => g.Key, g => g.First());

                var fields = await _context.Fields
                    .Where(f => projectconfig.EnvelopeMakingCriteria.Contains(f.FieldId))
                    .ToListAsync();

                var fieldNames = fields
                    .OrderBy(f => projectconfig.EnvelopeMakingCriteria.IndexOf(f.FieldId))
                    .Select(f => f.Name)
                    .ToList();

                IQueryable<EnvelopeBreakingResult> query = _context.EnvelopeBreakingResults
                    .Where(r => r.ProjectId == ProjectId && r.Status);

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

                    // Normal rows
                    if (result.NrDataId != 0 && nrDataDict.TryGetValue(result.NrDataId, out var nrData))
                    {
                        nr = nrData;
                    }
                    // MSS rows → fallback using CatchNo
                    else if (isMssRow && !string.IsNullOrEmpty(result.CatchNo) &&
                             nrDataByCatch.TryGetValue(result.CatchNo, out var catchNr))
                    {
                        nr = catchNr;
                    }

                    if (lotNo.HasValue && lotNo.Value > 0)
                    {
                        if (nr == null || nr.LotNo != lotNo.Value)
                            continue;
                    }

                    // NR fields (now works for MSS too)
                    if (nr != null)
                    {
                        rowDict["SubjectName"] = nr.SubjectName;
                        rowDict["Pages"] = nr.Pages;
                        rowDict["Symbol"] = nr.Symbol;
                        rowDict["Day"] = nr.Day;
                        rowDict["NRQuantity"] = nr.NRQuantity;
                    }

                    // JSON dynamic fields
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
                    rowDict["District"] = result.District;
                    rowDict["DistrictSort"] = result.DistrictSort;
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
                    .GroupBy(x => ((IDictionary<string, object>)x)["CatchNo"]?.ToString() ?? "")
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
                            case "DistrictSort":
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

                // ---------------------------------------------------------------
                // Reinsert MSS — rewritten to be safe regardless of how sorting
                // scatters identical CatchNo values, and to guarantee each
                // catch's MSS rows are inserted EXACTLY ONCE.
                // ---------------------------------------------------------------
                string mssMode = projectconfig.MssAttached?.ToLower();
                var finalSortedList = new List<dynamic>();
                var insertedMssCatches = new HashSet<string>();

                // Group sortedNonMss into contiguous runs by CatchNo, preserving order.
                // (A run = consecutive rows sharing the same CatchNo in the *current* sort order.)
                var runs = new List<(string CatchNo, List<dynamic> Rows)>();
                string currentCatch = null;
                List<dynamic> currentRun = null;

                foreach (var item in sortedNonMss)
                {
                    var dict = (IDictionary<string, object>)item;
                    string catchNo = dict["CatchNo"]?.ToString() ?? "";

                    if (currentRun == null || catchNo != currentCatch)
                    {
                        currentRun = new List<dynamic>();
                        runs.Add((catchNo, currentRun));
                        currentCatch = catchNo;
                    }

                    currentRun.Add(item);
                }

                // Walk the runs and insert MSS rows at most once per catch.
                for (int r = 0; r < runs.Count; r++)
                {
                    var (catchNo, rows) = runs[r];

                    bool isFirstRunForCatch = !runs.Take(r).Any(x => x.CatchNo == catchNo);
                    bool isLastRunForCatch = !runs.Skip(r + 1).Any(x => x.CatchNo == catchNo);

                    if (mssMode == "start" && isFirstRunForCatch &&
                        mssRowsByCatch.TryGetValue(catchNo, out var mssStart) &&
                        insertedMssCatches.Add(catchNo))
                    {
                        finalSortedList.AddRange(mssStart);
                    }

                    finalSortedList.AddRange(rows);

                    if (mssMode == "end" && isLastRunForCatch &&
                        mssRowsByCatch.TryGetValue(catchNo, out var mssEnd) &&
                        insertedMssCatches.Add(catchNo))
                    {
                        finalSortedList.AddRange(mssEnd);
                    }
                }

                // Safety net: if any catch's MSS rows were never inserted
                // (e.g. catch present only in MSS data, not in nonMssData), append them.
                foreach (var kvp in mssRowsByCatch)
                {
                    if (insertedMssCatches.Add(kvp.Key))
                    {
                        finalSortedList.AddRange(kvp.Value);
                    }
                }

                // Columns
                var allKeys = finalSortedList
                    .SelectMany(x => ((IDictionary<string, object>)x).Keys)
                    .Where(k => k != "isMss")
                    .Distinct()
                    .ToList();

                if (!finalSortedList.Any() || !allKeys.Any())
                    return BadRequest("No data available to generate Excel.");

                // --- Guard rails before writing to Excel ---
                const int ExcelMaxRows = 1048576;
                const int ExcelMaxCols = 16384;

                if (finalSortedList.Count + 1 > ExcelMaxRows) // +1 for header row
                    return BadRequest(new
                    {
                        error = $"Row count {finalSortedList.Count} exceeds Excel's row limit ({ExcelMaxRows}).",
                        hint = "Check MSS reinsertion logic or duplicate source data."
                    });

                if (allKeys.Count > ExcelMaxCols)
                    return BadRequest(new
                    {
                        error = $"Column count {allKeys.Count} exceeds Excel's column limit ({ExcelMaxCols}).",
                        hint = "Check NRDatas JSON fields for inconsistent/unexpected keys."
                    });

                // Excel
                var reportPath = FileStorageHelper.GetProjectFolder(ProjectId);

                var distinctLots = (lotNo.HasValue && lotNo.Value > 0)
                    ? new List<int> { lotNo.Value }
                    : nrDataDict.Values.Where(r => r.LotNo > 0).Select(r => r.LotNo).Distinct().OrderBy(l => l).ToList();
                var lotStr = distinctLots.Any() ? string.Join("_", distinctLots) : "All";

                var fileName = ReportVersionHelper.GetNextVersionFileName(reportPath, $"EnvelopeBreaking_{lotStr}.xlsx");
                var filePath = Path.Combine(reportPath, fileName);

                ExcelPackage.License.SetNonCommercialPersonal("Tools");
                using (var package = new ExcelPackage())
                {
                    var ws = package.Workbook.Worksheets.Add("Envelope Report");

                    for (int i = 0; i < allKeys.Count; i++)
                    {
                        ws.Cells[1, i + 1].Value = allKeys[i];
                        ws.Cells[1, i + 1].Style.Font.Bold = true;
                    }

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

                    if (ws.Dimension != null)
                        ws.Cells[ws.Dimension.Address].AutoFitColumns();

                    ws.View.FreezePanes(2, 1);

                    package.SaveAs(new FileInfo(filePath));
                }

                await ToolsAPI.Helpers.ExcelReportHelper.RecordExcelReportAsync(
                    _context,
                    ProjectId,
                    4, // Module 4 (Envelope Breaking)
                    1,
                    lotNo,
                    filePath,
                    true,
                    Tools.Services.LogHelper.GetTriggeredBy(User, Request)
                );

                return Ok(new
                {
                    message = "Report generated successfully",
                    filePath,
                    fileName,
                    recordsCount = finalSortedList.Count
                });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new
                {
                    error = ex.Message,
                    stack = ex.StackTrace
                });
            }
        }

        [HttpGet("CatchWithOmrSerialing")]
        public async Task<IActionResult> ProcessSerialingReport(int ProjectId, int? uploadId = null, [FromQuery] int? lotNo = null)
        {
            try
            {
                var isNewModel = await _context.NrData1.AnyAsync(p => p.ProjectId == ProjectId);
                var grouped = new List<dynamic>();

                if (isNewModel)
                {
                    var queryNew = _context.NewEnvelopeBreakingResults.Where(x => x.ProjectId == ProjectId && x.Status == true);

                    if (uploadId.HasValue)
                    {
                        var nrDatas = await _context.NrData1.Where(n => n.ProjectId == ProjectId && n.Status == true).ToListAsync();
                        var validNrDataIds = nrDatas
                            .Select(n => n.Id)
                            .ToList();

                        queryNew = queryNew.Where(x => validNrDataIds.Contains(x.NrDataId));
                    }
                    else
                    {
                        queryNew = queryNew.Where(x => x.BookletSerial != null || x.OmrSerial != null);
                    }

                    var dataNew = await queryNew.OrderBy(x => x.Id).ToListAsync();
                    var nrDataIds = dataNew.Select(x => x.NrDataId).Distinct().ToList();
                    var nrDatasMap = await _context.NrData1.Where(n => nrDataIds.Contains(n.Id)).ToDictionaryAsync(n => n.Id);

                    grouped = dataNew
                        .GroupBy(x => nrDatasMap.ContainsKey(x.NrDataId) ? nrDatasMap[x.NrDataId].CatchNo : "Unknown")
                        .Select(g =>
                        {
                            var firstSerial = g.First().OmrSerial ?? "";
                            var lastSerial = g.Last().OmrSerial ?? "";
                            var firstBooklet = g.First().BookletSerial ?? "";
                            var lastBooklet = g.Last().BookletSerial ?? "";
                            var start = string.IsNullOrEmpty(firstSerial) ? "" : firstSerial.Split('-')[0];
                            var end = string.IsNullOrEmpty(lastSerial) ? "" : lastSerial.Split('-').Last();
                            var first = string.IsNullOrEmpty(firstBooklet) ? "" : firstBooklet.Split('-')[0];
                            var last = string.IsNullOrEmpty(lastBooklet) ? "" : lastBooklet.Split('-').Last();

                            var firstRow = g.First();
                            var nr = nrDatasMap.ContainsKey(firstRow.NrDataId) ? nrDatasMap[firstRow.NrDataId] : null;

                            return (dynamic)new
                            {
                                CatchNo = g.Key,
                                OmrSerialRange = $"{start}-{end}",
                                BookletSerialRange = $"{first}-{last}",
                                ExamDate = nr?.ExamDate,
                                ExamTime = nr?.ExamTime
                            };
                        })
                        .ToList();
                }
                else
                {
                    var query = _context.EnvelopeBreakingResults.Where(x => x.ProjectId == ProjectId);
                    
                    if (uploadId.HasValue)
                    {
                        // Filter by uploadId via NRData relationship
                        var nrDatas = await _context.NRDatas
                            .Where(n => n.ProjectId == ProjectId)
                            .ToListAsync();
                        var validNrDataIds = nrDatas
                            .Where(n => n.UploadList != null && n.UploadList.Contains(uploadId.Value))
                            .Select(n => n.Id)
                            .ToList();
                        
                        query = query.Where(x => validNrDataIds.Contains(x.NrDataId));
                    }
                    else
                    {
                        query = query.Where(x => x.BookletSerial != null || x.OmrSerial != null);
                    }

                    var data = await query.OrderBy(x => x.Id).ToListAsync();

                    grouped = data
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
                            return (dynamic)new
                            {
                                CatchNo = g.Key,
                                OmrSerialRange = $"{start}-{end}",
                                BookletSerialRange = $"{first}-{last}",
                                ExamDate = g.First().ExamDate,
                                ExamTime = g.First().ExamTime
                            };
                        })
                        .ToList();
                }

                var reportPath = FileStorageHelper.GetProjectFolder(ProjectId);

                var fileName = uploadId.HasValue ? $"CatchWiseBookletAndOmrSerialing_v{uploadId}.xlsx" : ReportVersionHelper.GetNextVersionFileName(reportPath, "CatchWiseBookletAndOmrSerialing.xlsx");
                var filePath = Path.Combine(reportPath, fileName);
                ExcelPackage.License.SetNonCommercialPersonal("Tools");
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

                await ToolsAPI.Helpers.ExcelReportHelper.RecordExcelReportAsync(
                    _context,
                    ProjectId,
                    4, // Module 4 (Envelope Breaking Serialing)
                    uploadId ?? 1,
                    lotNo.HasValue && lotNo.Value > 0 ? lotNo : null,
                    filePath,
                    true,
                    Tools.Services.LogHelper.GetTriggeredBy(User, Request)
                );
                return Ok(new { message = "Excel generated successfully", path = filePath });
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        private async Task<IActionResult> ProcessEnvelopeBreakingNew(int ProjectId, int triggeredBy = 0, bool skipReset = false, int? lotNo = null, string? catchNo = null, bool bypassDispatch = false, int? batchNo = null)
        {
            try
            {
                if (!skipReset) await ResetReportStatus(ProjectId);

                var envCaps = await _context.EnvelopesTypes
                    .Select(e => new { e.EnvelopeName, e.Capacity })
                    .ToListAsync();
                var envelopeCapacities = envCaps.ToDictionary(x => x.EnvelopeName, x => x.Capacity);

                var eligibleSteps = Tools.Models.PipelineNavigator.GetEligiblePickupSteps(Tools.Models.PipelineNavigator.STEP_AWAITING_ENV);

                var nrQuery = _context.NrData1
                    .Where(p => p.ProjectId == ProjectId && p.Status == true && eligibleSteps.Contains(p.Steps) && p.Batch == (batchNo ?? 1));

                if (lotNo.HasValue && lotNo.Value > 0)
                    nrQuery = nrQuery.Where(p => p.LotNo == lotNo.Value);
                if (!string.IsNullOrEmpty(catchNo))
                    nrQuery = nrQuery.Where(p => p.CatchNo == catchNo);

                var nrDataList = await nrQuery
                    .OrderBy(p => p.CatchNo)
                    .ToListAsync();

                if (!nrDataList.Any())
                {
                    var fallbackQuery = _context.NrData1
                        .Where(p => p.ProjectId == ProjectId && p.Status == true && p.Batch == (batchNo ?? 1));
                    if (lotNo.HasValue && lotNo.Value > 0)
                        fallbackQuery = fallbackQuery.Where(p => p.LotNo == lotNo.Value);
                    if (!string.IsNullOrEmpty(catchNo))
                        fallbackQuery = fallbackQuery.Where(p => p.CatchNo == catchNo);
                    nrDataList = await fallbackQuery.OrderBy(p => p.CatchNo).ToListAsync();
                }

                if (!nrDataList.Any())
                    return BadRequest("No valid NR Data found for Envelope Breaking processing.");

                var nrDataIds = nrDataList.Select(n => n.Id).ToList();
                var centers = await _context.CenterList
                    .Where(c => c.ProjectId == ProjectId && nrDataIds.Contains(c.NRDataId) && c.Status)
                    .OrderBy(c => c.RouteSort)
                    .ThenBy(c => c.NodalSort)
                    .ThenBy(c => c.CenterSort)
                    .ToListAsync();

                if (!centers.Any())
                    return BadRequest("No centers found for the selected NR data.");

                var centerIds = centers.Select(c => c.Id).ToList();
                var newBreakages = await _context.NewEnvelopeBreakages
                    .Where(b => b.ProjectId == ProjectId && centerIds.Contains(b.CenterListId) && b.Status)
                    .ToListAsync();

                if (!newBreakages.Any())
                    return BadRequest("No Envelope Breakdown configuration found. Please run the Envelope Breakages (Inner/Outer) configuration first.");

                var envDict = newBreakages.ToDictionary(e => e.CenterListId);

                var projectconfig = await _context.ProjectConfigs
                    .FirstOrDefaultAsync(p => p.ProjectId == ProjectId);
                if (projectconfig == null)
                    return NotFound("Project config not found");

                var project = await _context.Projects
                    .FirstOrDefaultAsync(p => p.ProjectId == ProjectId);
                int projectTypeId = project?.TypeId ?? 0;

                var mssTypes = projectconfig.MssTypes ?? new List<int>();
                string mssMode = projectconfig.MssAttached?.ToLower() ?? "end";
                var mssData = new List<Mss>();

                if (!string.Equals(mssMode, "none", StringComparison.OrdinalIgnoreCase))
                {
                    var mssQuery = _context.Mss.AsQueryable();

                    if (projectTypeId > 0)
                    {
                        mssQuery = mssQuery.Where(m => m.TypeId == projectTypeId);
                    }

                    if (mssTypes.Any())
                    {
                        mssQuery = mssQuery.Where(m => mssTypes.Contains(m.Id));
                    }

                    mssData = await mssQuery.ToListAsync();

                    if (!mssData.Any() && projectTypeId > 0)
                    {
                        mssData = await _context.Mss.Where(m => m.TypeId == projectTypeId).ToListAsync();
                    }
                }

                var sortFieldIds = projectconfig.EnvelopeMakingCriteria ?? new List<int>();
                var sortFields = await _context.Fields
                    .Where(f => sortFieldIds.Contains(f.FieldId))
                    .ToListAsync();

                var sortingFieldNames = sortFields
                    .OrderBy(f => sortFieldIds.IndexOf(f.FieldId))
                    .Select(f => f.Name)
                    .ToList();

                var nrDataLookup = nrDataList.ToDictionary(nr => nr.Id);
                var nrDataJsonLookup = nrDataList.ToDictionary(
                    nr => nr.Id,
                    nr =>
                    {
                        if (string.IsNullOrWhiteSpace(nr.NRDatas)) return null;
                        try { return JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(nr.NRDatas); }
                        catch { return null; }
                    }
                );

                string GetNrField(Dictionary<string, JsonElement>? nrDynamic, string fieldName)
                {
                    if (nrDynamic == null) return "";
                    var match = nrDynamic.FirstOrDefault(k =>
                        k.Key.Equals(fieldName, StringComparison.OrdinalIgnoreCase));
                    return match.Key != null ? (match.Value.GetString() ?? "") : "";
                }

                void FillDynamicFields(IDictionary<string, object> rowDict, int nrDataId)
                {
                    if (!nrDataJsonLookup.TryGetValue(nrDataId, out var nrDynamic) || nrDynamic == null)
                        return;

                    foreach (var fieldName in sortingFieldNames)
                    {
                        if (!rowDict.ContainsKey(fieldName))
                        {
                            rowDict[fieldName] = GetNrField(nrDynamic, fieldName);
                        }
                    }
                }

                string? GetPackingDenominationForQuantity(string? innerEnvelopeJson, int targetQty, ref List<int>? innerEnvelopesPool)
                {
                    if (innerEnvelopesPool == null)
                    {
                        innerEnvelopesPool = new List<int>();
                        if (!string.IsNullOrWhiteSpace(innerEnvelopeJson))
                        {
                            try
                            {
                                var dict = JsonSerializer.Deserialize<Dictionary<string, string>>(innerEnvelopeJson);
                                if (dict != null)
                                {
                                    foreach (var kvp in dict)
                                    {
                                        string envName = kvp.Key.Replace("E", "");
                                        if (int.TryParse(envName, out int capacity) && int.TryParse(kvp.Value, out int count))
                                        {
                                            for (int i = 0; i < count; i++)
                                                innerEnvelopesPool.Add(capacity);
                                        }
                                    }
                                }
                            }
                            catch { }
                        }
                        innerEnvelopesPool = innerEnvelopesPool.OrderByDescending(x => x).ToList();
                    }

                    if (targetQty <= 0 || innerEnvelopesPool.Count == 0) return null;

                    var used = new Dictionary<int, int>();
                    int currentSum = 0;
                    var remainingPool = new List<int>();

                    foreach (var cap in innerEnvelopesPool)
                    {
                        if (currentSum < targetQty && currentSum + cap <= targetQty)
                        {
                            currentSum += cap;
                            if (!used.ContainsKey(cap)) used[cap] = 0;
                            used[cap]++;
                        }
                        else
                        {
                            remainingPool.Add(cap);
                        }
                    }

                    innerEnvelopesPool = remainingPool;
                    if (used.Count == 0) return null;
                    return "[" + string.Join(" & ", used.OrderByDescending(x => x.Key).Select(kvp => $"{kvp.Key}x{kvp.Value}")) + "]";
                }

                List<dynamic> CreateMssRows(string catchNo, string examDate, string examTime, string courseName, int nrDataId)
                {
                    var rows = new List<dynamic>();
                    foreach (var mss in mssData)
                    {
                        var mssRow = new ExpandoObject();
                        var mssDict = (IDictionary<string, object>)mssRow;
                        mssDict["isMss"] = true;
                        mssDict["CatchNo"] = catchNo;
                        mssDict["EnvQuantity"] = mss.MssType;
                        mssDict["NodalCode"] = "";
                        mssDict["CenterEnv"] = 0;
                        mssDict["TotalEnv"] = 0;
                        mssDict["Env"] = "";
                        mssDict["NrDataId"] = nrDataId;
                        mssDict["CenterListId"] = 0;
                        mssDict["PackingDenomination"] = null;
                        rows.Add(mssRow);
                    }
                    return rows;
                }

                var resultList = new List<dynamic>();
                string? prevMergeField = null;
                int centerEnvCounter = 0;

                foreach (var center in centers)
                {
                    if (!nrDataLookup.TryGetValue(center.NRDataId, out var parentNr)) continue;

                    envDict.TryGetValue(center.Id, out var envInfo);
                    int totalEnv = envInfo != null && envInfo.TotalEnvelope > 0 ? envInfo.TotalEnvelope : 1;

                    string currentMergeField = $"{parentNr.CatchNo}-{center.CenterCode}";
                    if (currentMergeField != prevMergeField)
                    {
                        centerEnvCounter = 0;
                        prevMergeField = currentMergeField;
                    }

                    var envelopeBreakdown = new List<(string EnvType, int Count, int Capacity)>();

                    if (envInfo != null && !string.IsNullOrEmpty(envInfo.OuterEnvelope))
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
                        catch { }
                    }

                    envelopeBreakdown = envelopeBreakdown.OrderBy(x => x.Capacity).ToList();
                    int centerQty = center.Quantity > 0 ? center.Quantity : center.NRQuantity;

                    if (envelopeBreakdown.Count > 0)
                    {
                        List<int>? innerPool = null;
                        string? innerJson = envInfo?.InnerEnvelope;
                        int envelopeIndex = 1;
                        int remainingQty = centerQty;

                        foreach (var (envType, count, capacity) in envelopeBreakdown)
                        {
                            for (int k = 0; k < count; k++)
                            {
                                centerEnvCounter++;
                                int envQty = remainingQty > capacity ? capacity : remainingQty;

                                var row = new ExpandoObject();
                                var rowDict = (IDictionary<string, object>)row;

                                rowDict["isMss"] = false;
                                rowDict["CatchNo"] = parentNr.CatchNo ?? "";
                                rowDict["CenterCode"] = center.CenterCode ?? "";
                                rowDict["CourseName"] = parentNr.CourseName ?? "";
                                rowDict["ExamTime"] = parentNr.ExamTime ?? "";
                                rowDict["ExamDate"] = parentNr.ExamDate ?? "";
                                rowDict["Quantity"] = centerQty;
                                rowDict["EnvQuantity"] = envQty;
                                rowDict["NodalCode"] = center.NodalCode ?? "";
                                rowDict["CenterEnv"] = centerEnvCounter;
                                rowDict["TotalEnv"] = totalEnv;
                                rowDict["Env"] = $"{envelopeIndex}/{totalEnv}";
                                rowDict["NRQuantity"] = center.NRQuantity;
                                rowDict["CenterSort"] = center.CenterSort;
                                rowDict["NodalSort"] = center.NodalSort;
                                rowDict["Route"] = center.Route ?? "";
                                rowDict["RouteSort"] = center.RouteSort;
                                rowDict["District"] = center.District ?? "";
                                rowDict["DistrictSort"] = center.DistrictSort;
                                rowDict["NrDataId"] = parentNr.Id;
                                rowDict["CenterListId"] = center.Id;

                                // Compute packing denomination: inner breakdown or fallback to envelope capacity
                                string? packingDenom = GetPackingDenominationForQuantity(innerJson, envQty, ref innerPool);
                                if (string.IsNullOrEmpty(packingDenom))
                                {
                                    int denomCap = capacity > 0 ? capacity : envQty;
                                    packingDenom = $"[{denomCap}x1]";
                                }
                                rowDict["PackingDenomination"] = packingDenom;

                                FillDynamicFields(rowDict, parentNr.Id);
                                resultList.Add(row);

                                remainingQty -= envQty;
                                envelopeIndex++;
                                if (remainingQty <= 0) break;
                            }
                            if (remainingQty <= 0) break;
                        }
                    }
                    else
                    {
                        centerEnvCounter++;
                        var row = new ExpandoObject();
                        var rowDict = (IDictionary<string, object>)row;

                        rowDict["isMss"] = false;
                        rowDict["CatchNo"] = parentNr.CatchNo ?? "";
                        rowDict["CenterCode"] = center.CenterCode ?? "";
                        rowDict["CourseName"] = parentNr.CourseName ?? "";
                        rowDict["ExamTime"] = parentNr.ExamTime ?? "";
                        rowDict["ExamDate"] = parentNr.ExamDate ?? "";
                        rowDict["Quantity"] = centerQty;
                        rowDict["EnvQuantity"] = centerQty;
                        rowDict["NodalCode"] = center.NodalCode ?? "";
                        rowDict["CenterEnv"] = centerEnvCounter;
                        rowDict["TotalEnv"] = 1;
                        rowDict["Env"] = "1/1";
                        rowDict["NRQuantity"] = center.NRQuantity;
                        rowDict["CenterSort"] = center.CenterSort;
                        rowDict["NodalSort"] = center.NodalSort;
                        rowDict["Route"] = center.Route ?? "";
                        rowDict["RouteSort"] = center.RouteSort;
                        rowDict["District"] = center.District ?? "";
                        rowDict["DistrictSort"] = center.DistrictSort;
                        rowDict["NrDataId"] = parentNr.Id;
                        rowDict["CenterListId"] = center.Id;

                        List<int>? innerPool = null;
                        string? packingDenom = GetPackingDenominationForQuantity(envInfo?.InnerEnvelope, centerQty, ref innerPool);
                        if (string.IsNullOrEmpty(packingDenom))
                        {
                            packingDenom = $"[{centerQty}x1]";
                        }
                        rowDict["PackingDenomination"] = packingDenom;

                        FillDynamicFields(rowDict, parentNr.Id);
                        resultList.Add(row);
                    }
                }

                // Sorting
                var nonMssRows = resultList
                    .Where(r =>
                    {
                        var d = (IDictionary<string, object>)r;
                        return d.ContainsKey("isMss") && !(bool)d["isMss"];
                    }).ToList();

                IOrderedEnumerable<dynamic>? sortedResultList = null;
                foreach (var fieldName in sortingFieldNames)
                {
                    Func<dynamic, object> keySelector = x =>
                    {
                        var dict = (IDictionary<string, object>)x;
                        if (!dict.ContainsKey(fieldName)) return null!;
                        var val = dict[fieldName];
                        if (val == null) return null!;

                        if (fieldName.Equals("NodalSort", StringComparison.OrdinalIgnoreCase))
                            return double.TryParse(val.ToString(), out double n) ? n : 0.0;

                        if (fieldName.Equals("ExamDate", StringComparison.OrdinalIgnoreCase))
                            return DateTime.TryParseExact(val.ToString(), "dd-MM-yyyy",
                                CultureInfo.InvariantCulture, DateTimeStyles.None, out var d)
                                ? (object)d : DateTime.MinValue;

                        if (fieldName.Contains("sort", StringComparison.OrdinalIgnoreCase))
                            return double.TryParse(val.ToString(), out double d2) ? d2 : 0.0;

                        return val.ToString()?.Trim().ToLowerInvariant() ?? "";
                    };

                    sortedResultList = sortedResultList == null
                        ? nonMssRows.OrderBy(keySelector)
                        : sortedResultList.ThenBy(keySelector);
                }

                var sortedNonMss = sortedResultList?.Cast<dynamic>().ToList() ?? nonMssRows;

                // Re-insert MSS rows at correct positions
                var finalResultList = new List<dynamic>();
                var insertedMssCatches = new HashSet<string>();

                var runs = new List<(string CatchNo, List<dynamic> Rows)>();
                string? currentCatch = null;
                List<dynamic>? currentRun = null;

                foreach (var item in sortedNonMss)
                {
                    var dict = (IDictionary<string, object>)item;
                    string cNo = dict["CatchNo"]?.ToString() ?? "";

                    if (currentRun == null || cNo != currentCatch)
                    {
                        currentRun = new List<dynamic>();
                        runs.Add((cNo, currentRun));
                        currentCatch = cNo;
                    }

                    currentRun.Add(item);
                }

                for (int r = 0; r < runs.Count; r++)
                {
                    var (cNo, rows) = runs[r];
                    bool isFirstRunForCatch = !runs.Take(r).Any(x => x.CatchNo == cNo);
                    bool isLastRunForCatch = !runs.Skip(r + 1).Any(x => x.CatchNo == cNo);

                    if (mssMode == "start" && isFirstRunForCatch && insertedMssCatches.Add(cNo))
                    {
                        var firstRow = (IDictionary<string, object>)rows.First();
                        int nrId = (int)(firstRow["NrDataId"] ?? 0);
                        if (nrId == 0) nrId = nrDataList.FirstOrDefault(n => n.CatchNo == cNo)?.Id ?? 0;
                        finalResultList.AddRange(CreateMssRows(
                            cNo,
                            firstRow["ExamDate"]?.ToString() ?? "",
                            firstRow["ExamTime"]?.ToString() ?? "",
                            firstRow["CourseName"]?.ToString() ?? "",
                            nrId
                        ));
                    }

                    finalResultList.AddRange(rows);

                    if (mssMode == "end" && isLastRunForCatch && insertedMssCatches.Add(cNo))
                    {
                        var lastRow = (IDictionary<string, object>)rows.Last();
                        int nrId = (int)(lastRow["NrDataId"] ?? 0);
                        if (nrId == 0) nrId = nrDataList.FirstOrDefault(n => n.CatchNo == cNo)?.Id ?? 0;
                        finalResultList.AddRange(CreateMssRows(
                            cNo,
                            lastRow["ExamDate"]?.ToString() ?? "",
                            lastRow["ExamTime"]?.ToString() ?? "",
                            lastRow["CourseName"]?.ToString() ?? "",
                            nrId
                        ));
                    }
                }

                // Safety net: add MSS for any catches in nrDataList that had no rows in sortedNonMss
                foreach (var nr in nrDataList)
                {
                    if (!string.IsNullOrEmpty(nr.CatchNo) && insertedMssCatches.Add(nr.CatchNo))
                    {
                        finalResultList.AddRange(CreateMssRows(
                            nr.CatchNo,
                            nr.ExamDate ?? "",
                            nr.ExamTime ?? "",
                            nr.CourseName ?? "",
                            nr.Id
                        ));
                    }
                }

                // Serials
                int bookletStart = projectconfig?.BookletSerialNumber ?? 0;
                int omrStart = projectconfig?.OmrSerialNumber ?? 0;
                bool resetOmrSerialOnCatchChange = projectconfig?.ResetOmrSerialOnCatchChange ?? false;
                bool resetBookletSerialOnCatchChange = projectconfig?.ResetBookletSerialOnCatchChange ?? false;
                int serial = 1;

                bool assignBookletSerial = bookletStart > 0;
                bool assignOmrSerial = omrStart > 0;
                string? prevCatchForSerial = null;

                var newEnvelopeResults = new List<NewEnvelopeBreakingResult>();

                foreach (var item in finalResultList)
                {
                    var dict = (IDictionary<string, object>)item;
                    bool isMssRow = dict.ContainsKey("isMss") && (bool)dict["isMss"];
                    string cNo = dict["CatchNo"]?.ToString() ?? "";

                    if (!isMssRow && prevCatchForSerial != null && cNo != prevCatchForSerial)
                    {
                        serial = 1;
                        if (resetOmrSerialOnCatchChange) omrStart = projectconfig?.OmrSerialNumber ?? 0;
                        if (resetBookletSerialOnCatchChange) bookletStart = projectconfig?.BookletSerialNumber ?? 0;
                    }

                    if (isMssRow)
                    {
                        int mssNrId = (int)(dict["NrDataId"] ?? 0);
                        if (mssNrId == 0)
                        {
                            mssNrId = nrDataList.FirstOrDefault(n => n.CatchNo == cNo)?.Id ?? 0;
                        }

                        newEnvelopeResults.Add(new NewEnvelopeBreakingResult
                        {
                            ProjectId = ProjectId,
                            NrDataId = mssNrId,
                            CenterListId = 0,
                            EnvQuantity = dict["EnvQuantity"]?.ToString() ?? "MSS",
                            CenterEnv = 0,
                            TotalEnv = 0,
                            Env = "",
                            SerialNumber = 0,
                            BookletSerial = "",
                            OmrSerial = "",
                            PackingDenomination = null,
                            CreatedAt = DateTime.UtcNow,
                            Status = true
                        });
                        continue;
                    }

                    int centerListId = (int)(dict["CenterListId"] ?? 0);
                    int nrDataId = (int)(dict["NrDataId"] ?? 0);
                    int envQuantity = (int)(dict["EnvQuantity"] ?? 0);

                    string bookletSerial = "";
                    string omrSerial = "";

                    if (assignBookletSerial)
                    {
                        bookletSerial = $"{bookletStart}-{bookletStart + envQuantity - 1}";
                        bookletStart += envQuantity;
                    }

                    if (assignOmrSerial)
                    {
                        omrSerial = $"{omrStart}-{omrStart + envQuantity - 1}";
                        omrStart += envQuantity;
                    }

                    newEnvelopeResults.Add(new NewEnvelopeBreakingResult
                    {
                        ProjectId = ProjectId,
                        NrDataId = nrDataId,
                        CenterListId = centerListId,
                        EnvQuantity = envQuantity.ToString(),
                        CenterEnv = (int)dict["CenterEnv"],
                        TotalEnv = (int)dict["TotalEnv"],
                        Env = dict["Env"]?.ToString(),
                        SerialNumber = serial++,
                        BookletSerial = bookletSerial,
                        OmrSerial = omrSerial,
                        PackingDenomination = dict.ContainsKey("PackingDenomination") ? dict["PackingDenomination"]?.ToString() : null,
                        CreatedAt = DateTime.UtcNow,
                        Status = true
                    });

                    prevCatchForSerial = cNo;
                }

                // Deactivate previous active records for these catches
                var existingNewBreakingIds = await _context.NewEnvelopeBreakingResults
                    .Where(r => r.ProjectId == ProjectId && nrDataIds.Contains(r.NrDataId) && r.Status)
                    .Select(r => r.Id)
                    .ToListAsync();

                if (existingNewBreakingIds.Any())
                {
                    await _context.NewEnvelopeBreakingResults
                        .Where(r => existingNewBreakingIds.Contains(r.Id))
                        .ExecuteUpdateAsync(s => s.SetProperty(r => r.Status, false));
                }

                await _context.NewEnvelopeBreakingResults.AddRangeAsync(newEnvelopeResults);

                foreach (var nr in nrDataList)
                {
                    nr.Steps = Tools.Models.PipelineNavigator.STEP_AWAITING_ENV;
                }

                await _context.SaveChangesAsync();

                await _loggerService.LogEventAsync(
                    $"Saved {newEnvelopeResults.Count} new envelope breaking results for ProjectId {ProjectId}",
                    "EnvelopeBreakageProcessing",
                    triggeredBy,
                    ProjectId);

                try
                {
                    await GetEnvelopeBreakingReport(ProjectId, lotNo);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[ProcessEnvelopeBreakingNew] Report error: {ex.Message}");
                }

                return Ok(new
                {
                    message = "Envelope breaking processed successfully for new models",
                    recordsCount = newEnvelopeResults.Count,
                    lotNo = lotNo,
                    catchesProcessed = nrDataList.Select(n => n.CatchNo).Distinct().ToList()
                });
            }
            catch (Exception ex)
            {
                await _loggerService.LogErrorAsync("Error during new envelope breaking processing", ex.ToString(), nameof(EnvelopeBreakageProcessingController));
                return StatusCode(500, new { message = "Error during new envelope breaking processing", error = ex.Message });
            }
        }

        private async Task<IActionResult> GetEnvelopeBreakingReportNew(int ProjectId, int? lotNo = null)
        {
            try
            {
                var projectconfig = await _context.ProjectConfigs
                    .FirstOrDefaultAsync(p => p.ProjectId == ProjectId);

                if (projectconfig == null)
                    return NotFound("Project config not found");

                var nrDataDict = await _context.NrData1
                    .Where(p => p.ProjectId == ProjectId && p.Status == true)
                    .ToDictionaryAsync(p => p.Id);

                var centerDict = await _context.CenterList
                    .Where(p => p.ProjectId == ProjectId && p.Status == true)
                    .ToDictionaryAsync(p => p.Id);

                var nrDataByCatch = await _context.NrData1
                    .Where(p => p.ProjectId == ProjectId && p.Status == true)
                    .GroupBy(p => p.CatchNo)
                    .ToDictionaryAsync(g => g.Key ?? "", g => g.First());

                var sortFieldIds = projectconfig.EnvelopeMakingCriteria ?? new List<int>();
                var fields = await _context.Fields
                    .Where(f => sortFieldIds.Contains(f.FieldId))
                    .ToListAsync();

                var fieldNames = fields
                    .OrderBy(f => sortFieldIds.IndexOf(f.FieldId))
                    .Select(f => f.Name)
                    .ToList();

                var results = await _context.NewEnvelopeBreakingResults
                    .Where(r => r.ProjectId == ProjectId && r.Status)
                    .OrderBy(r => r.Id)
                    .ToListAsync();

                if (!results.Any())
                    return NotFound("No envelope breaking results found");

                var fullData = new List<dynamic>();

                foreach (var result in results)
                {
                    var row = new ExpandoObject();
                    var rowDict = (IDictionary<string, object>)row;

                    bool isMssRow = result.CenterListId == 0 && result.SerialNumber == 0;
                    rowDict["isMss"] = isMssRow;

                    NrData1? nr = null;
                    CenterList? center = null;

                    if (result.NrDataId != 0 && nrDataDict.TryGetValue(result.NrDataId, out var foundNr))
                    {
                        nr = foundNr;
                    }

                    if (result.CenterListId != 0 && centerDict.TryGetValue(result.CenterListId, out var foundCenter))
                    {
                        center = foundCenter;
                    }

                    if (lotNo.HasValue && lotNo.Value > 0)
                    {
                        if (nr == null || (nr.LotNo != lotNo.Value && nr.EnvLotNo != lotNo.Value))
                            continue;
                    }

                    if (nr != null)
                    {
                        rowDict["SubjectName"] = nr.SubjectName ?? "";
                        rowDict["Pages"] = nr.Pages;
                        rowDict["Symbol"] = "";
                        rowDict["Day"] = nr.Day ?? "";
                        rowDict["NRQuantity"] = center?.NRQuantity ?? 0;
                    }

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

                    rowDict["SerialNo"] = isMssRow ? (object)"" : result.SerialNumber;
                    rowDict["CatchNo"] = nr?.CatchNo ?? "";
                    rowDict["CenterCode"] = isMssRow ? "" : (center?.CenterCode ?? "");
                    rowDict["CenterSort"] = center?.CenterSort ?? 0.0;
                    rowDict["ExamTime"] = nr?.ExamTime ?? "";
                    rowDict["ExamDate"] = nr?.ExamDate ?? "";
                    rowDict["Quantity"] = center?.Quantity ?? 0;
                    rowDict["NodalCode"] = center?.NodalCode ?? "";
                    rowDict["NodalSort"] = center?.NodalSort ?? 0.0;
                    rowDict["Route"] = center?.Route ?? "";
                    rowDict["RouteSort"] = center?.RouteSort ?? 0;
                    rowDict["District"] = center?.District ?? "";
                    rowDict["DistrictSort"] = center?.DistrictSort ?? 0;
                    rowDict["EnvQuantity"] = result.EnvQuantity ?? "";
                    rowDict["CenterEnv"] = result.CenterEnv;
                    rowDict["TotalEnv"] = result.TotalEnv;
                    rowDict["Env"] = result.Env ?? "";
                    rowDict["SerialNumber"] = isMssRow ? 0 : result.SerialNumber;
                    rowDict["BookletSerial"] = result.BookletSerial ?? "";
                    rowDict["OmrSerial"] = result.OmrSerial ?? "";
                    rowDict["CourseName"] = nr?.CourseName ?? "";
                    rowDict["PackingDenomination"] = result.PackingDenomination ?? "";

                    fullData.Add(row);
                }

                // Generate Excel using EPPlus and record in ExcelReport
                var reportPath = FileStorageHelper.GetProjectFolder(ProjectId);
                var distinctLots = (lotNo.HasValue && lotNo.Value > 0)
                    ? new List<int> { lotNo.Value }
                    : nrDataDict.Values.Where(r => r.LotNo > 0 || r.EnvLotNo > 0).Select(r => r.LotNo > 0 ? r.LotNo : r.EnvLotNo).Distinct().OrderBy(l => l).ToList();
                var lotStr = distinctLots.Any() ? string.Join("_", distinctLots) : "All";

                var fileName = ReportVersionHelper.GetNextVersionFileName(reportPath, $"EnvelopeBreaking_New_{lotStr}.xlsx");
                var filePath = Path.Combine(reportPath, fileName);

                ExcelPackage.License.SetNonCommercialPersonal("Tools");
                using (var package = new ExcelPackage())
                {
                    var ws = package.Workbook.Worksheets.Add("Envelope Report");
                    var allKeys = fullData.SelectMany(x => ((IDictionary<string, object>)x).Keys).Where(k => k != "isMss").Distinct().ToList();

                    for (int i = 0; i < allKeys.Count; i++)
                    {
                        ws.Cells[1, i + 1].Value = allKeys[i];
                        ws.Cells[1, i + 1].Style.Font.Bold = true;
                    }

                    int rowIdx = 2;
                    foreach (var item in fullData)
                    {
                        var dict = (IDictionary<string, object>)item;
                        for (int col = 0; col < allKeys.Count; col++)
                        {
                            var key = allKeys[col];
                            ws.Cells[rowIdx, col + 1].Value = dict.ContainsKey(key) ? dict[key]?.ToString() : "";
                        }
                        rowIdx++;
                    }

                    if (ws.Dimension != null)
                        ws.Cells[ws.Dimension.Address].AutoFitColumns();

                    package.SaveAs(new FileInfo(filePath));
                }

                await ToolsAPI.Helpers.ExcelReportHelper.RecordExcelReportAsync(
                    _context,
                    ProjectId,
                    4, // Module 4 (Envelope Breaking)
                    1,
                    lotNo,
                    filePath,
                    true,
                    Tools.Services.LogHelper.GetTriggeredBy(User, Request)
                );

                return Ok(new
                {
                    message = "Report generated successfully",
                    filePath,
                    fileName,
                    recordsCount = fullData.Count
                });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { message = "Error during Excel report generation", error = ex.Message });
            }
        }

        private async Task ResetReportStatus(int projectId)
        {
            try
            {
                await _context.RPTTemplates
                    .Where(t => t.ProjectId == projectId && t.ReportStatus)
                    .ExecuteUpdateAsync(setters => setters.SetProperty(t => t.ReportStatus, false));
            }
            catch (Exception ex)
            {
                await _loggerService.LogErrorAsync("Report Status Reset Error", ex.Message, nameof(EnvelopeBreakageProcessingController));
            }
        }
    }
}

