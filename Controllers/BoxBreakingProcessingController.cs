using ERPToolsAPI.Data;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Options;
using OfficeOpenXml;
using System.Reflection;
using System.Text.Json;
using Tools.Models;
using Tools.Services;

namespace Tools.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class BoxBreakingProcessingController : ControllerBase
    {
        private readonly ERPToolsDbContext _context;
        private readonly ILoggerService _loggerService;
        private readonly ApiSettings _apiSettings;

        public BoxBreakingProcessingController(ERPToolsDbContext context, ILoggerService loggerService, IOptions<ApiSettings> apiSettings)
        {
            _context = context;
            _loggerService = loggerService;
            _apiSettings = apiSettings.Value;
        }

        [HttpPost("ProcessBoxBreaking")]
        public async Task<IActionResult> ProcessBoxBreaking(int ProjectId, List<int> LotNo)
        {
            try
            {
                var maxBatch = await _context.EnvelopeBreakingResults
                    .Where(r => r.ProjectId == ProjectId)
                    .MaxAsync(r => (int?)r.UploadBatch);

                if (!maxBatch.HasValue)
                    return NotFound("No envelope breaking results found. Run ProcessEnvelopeBreaking first.");

                var projectconfig = await _context.ProjectConfigs
                    .Where(p => p.ProjectId == ProjectId)
                    .FirstOrDefaultAsync();

                if (projectconfig == null)
                    return NotFound("Project config not found");

                // Load config once - shared across all lots
                var boxIds = projectconfig.BoxBreakingCriteria;
                var sortingId = projectconfig.SortingBoxReport;
                var duplicatesFields = projectconfig.DuplicateRemoveFields;
                bool InnerBundling = projectconfig.IsInnerBundlingDone;

                var innerBundlingFieldNames = new List<string>();
                if (InnerBundling && projectconfig.InnerBundlingCriteria != null)
                {
                    innerBundlingFieldNames = await _context.Fields
                        .Where(f => projectconfig.InnerBundlingCriteria.Contains(f.FieldId))
                        .Select(f => f.Name)
                        .ToListAsync();
                }

                var fields = await _context.Fields.Where(f => boxIds.Contains(f.FieldId)).ToListAsync();
                var fieldsFromDb = await _context.Fields.Where(f => sortingId.Contains(f.FieldId)).ToListAsync();
                var fieldNames = fieldsFromDb.OrderBy(f => sortingId.IndexOf(f.FieldId)).Select(f => f.Name).ToList();
                var dupNames = await _context.Fields.Where(f => duplicatesFields.Contains(f.FieldId)).Select(f => f.Name).ToListAsync();

                var startBox = projectconfig.BoxNumber;
                bool resetOnSymbolChange = projectconfig.ResetOnSymbolChange;
                var capacity = await _context.BoxCapacity
                    .Where(c => c.BoxCapacityId == projectconfig.BoxCapacity)
                    .Select(c => c.Capacity)
                    .FirstOrDefaultAsync();

                var envelopeObj = JsonSerializer.Deserialize<Dictionary<string, string>>(projectconfig.Envelope);
                int envelopeSize = 0;
                if (envelopeObj != null && envelopeObj.ContainsKey("Outer"))
                {
                    var outerParts = envelopeObj["Outer"].Split(',', StringSplitOptions.RemoveEmptyEntries);
                    if (outerParts.Length == 1)
                    {
                        string digits = new string(outerParts[0].Trim().Where(char.IsDigit).ToArray());
                        if (int.TryParse(digits, out int parsedValue) && parsedValue > 0)
                            envelopeSize = parsedValue;
                    }
                }

                var allResults = new List<object>();
                int totalSaved = 0;

                // ✅ Process each lot independently
                foreach (var lot in LotNo)
                {
                    // Load NRData only for this lot
                    var nrData = await _context.NRDatas
                        .Where(p => p.ProjectId == ProjectId && p.Status == true && p.LotNo == lot)
                        .ToListAsync();

                    if (!nrData.Any())
                    {
                        allResults.Add(new { lot, skipped = true, reason = "No NRData found" });
                        continue;
                    }

                    var nrCatchNos = nrData.Select(n => n.CatchNo).ToList();

                    var envelopeResults = await _context.EnvelopeBreakingResults
                        .Where(r => r.ProjectId == ProjectId && r.UploadBatch == maxBatch.Value
                                 && r.SerialNumber != 0 && nrCatchNos.Contains(r.CatchNo))
                        .ToListAsync();

                    if (!envelopeResults.Any())
                    {
                        allResults.Add(new { lot, skipped = true, reason = "No envelope results found" });
                        continue;
                    }

                    // Build data rows for this lot
                    var breakingReportData = new List<dynamic>();
                    foreach (var result in envelopeResults)
                    {
                        dynamic row = new System.Dynamic.ExpandoObject();
                        var rowDict = (IDictionary<string, object>)row;

                        rowDict["CatchNo"] = result.CatchNo ?? "";
                        rowDict["CenterCode"] = result.CenterCode ?? "";
                        rowDict["CenterSort"] = result.CenterSort;
                        rowDict["ExamTime"] = result.ExamTime ?? "";
                        rowDict["ExamDate"] = result.ExamDate;
                        rowDict["Quantity"] = result.Quantity;
                        rowDict["TotalEnv"] = result.TotalEnv;
                        rowDict["NodalCode"] = result.NodalCode ?? "";
                        rowDict["NodalSort"] = result.NodalSort;
                        rowDict["Route"] = result.Route ?? "";
                        rowDict["RouteSort"] = result.RouteSort;
                        rowDict["CourseName"] = result.CourseName ?? "";
                        rowDict["BookletSerial"] = result.BookletSerial ?? "";
                        rowDict["OmrSerial"] = result.OmrSerial;
                        rowDict["EnvelopeBreakingResultId"] = result.Id;
                        rowDict["NrDataId"] = result.NrDataId;

                        var nrRow = nrData.FirstOrDefault(n => n.CatchNo == result.CatchNo);
                        rowDict["Symbol"] = nrRow?.Symbol ?? "";
                        rowDict["Pages"] = nrRow?.Pages ?? 0;
                        rowDict["NRQuantity"] = nrRow?.NRQuantity ?? 0;

                        breakingReportData.Add(row);
                    }

                    if (!breakingReportData.Any()) continue;

                    // Deduplicate
                    var uniqueRows = breakingReportData
                        .GroupBy(x =>
                        {
                            var keyParts = dupNames.Select(fieldName =>
                            {
                                var dict = (IDictionary<string, object>)x;
                                return dict.ContainsKey(fieldName) ? dict[fieldName]?.ToString()?.Trim() ?? "" : "";
                            });
                            return string.Join("_", keyParts);
                        })
                        .Select(g => g.First())
                        .ToList();

                    // Start/End/Serial
                    var enrichedList = new List<dynamic>();
                    string previousCatchNo = null;
                    int previousEnd = 0;

                    foreach (var row in uniqueRows)
                    {
                        var rowDict = (IDictionary<string, object>)row;
                        string catchNo = rowDict["CatchNo"]?.ToString();
                        int start = catchNo != previousCatchNo ? 1 : previousEnd + 1;
                        int end = start + (int)rowDict["TotalEnv"] - 1;
                        rowDict["Start"] = start;
                        rowDict["End"] = end;
                        rowDict["Serial"] = $"{start} to {end}";
                        enrichedList.Add(row);
                        previousCatchNo = catchNo;
                        previousEnd = end;
                    }

                    // Sorting
                    IOrderedEnumerable<dynamic> ordered = null;
                    foreach (var fieldName in fieldNames)
                    {
                        Func<dynamic, object> keySelector = x =>
                        {
                            var dict = (IDictionary<string, object>)x;
                            if (!dict.ContainsKey(fieldName)) return null;
                            var val = dict[fieldName];
                            if (val == null) return null;

                            if (fieldName.Equals("NodalSort", StringComparison.OrdinalIgnoreCase))
                                return double.TryParse(val.ToString(), out double n) ? (object)n : 0.0;
                            if (fieldName.Equals("CenterSort", StringComparison.OrdinalIgnoreCase))
                                return int.TryParse(val.ToString(), out int c) ? (object)c : 0;
                            if (fieldName.Equals("RouteSort", StringComparison.OrdinalIgnoreCase))
                                return int.TryParse(val.ToString(), out int r) ? (object)r : 0;
                            if (fieldName.Equals("ExamDate", StringComparison.OrdinalIgnoreCase))
                            {
                                string[] formats = { "dd-MM-yyyy", "d-M-yyyy", "dd/MM/yyyy", "yyyy-MM-dd" };
                                return DateTime.TryParseExact(val.ToString().Trim(), formats,
                                    System.Globalization.CultureInfo.InvariantCulture,
                                    System.Globalization.DateTimeStyles.None, out DateTime d)
                                    ? (object)d : DateTime.MinValue;
                            }
                            return val.ToString().Trim().ToLowerInvariant();
                        };

                        ordered = ordered == null ? enrichedList.OrderBy(keySelector) : ordered.ThenBy(keySelector);
                    }

                    var sortedList = ordered?.ToList() ?? enrichedList;

                    // ✅ Fresh counters per lot
                    var finalWithBoxes = new List<dynamic>();
                    int boxNo = startBox;
                    int runningPages = 0;
                    string prevMergeKey = null;
                    long runningOmrPointer = 0;
                    long runningBooklet = 0;
                    string previousCatchForOmr = null;
                    string previousCourse = null;
                    int innerBundlingSerial = 0;
                    string prevInnerBundlingKey = null;

                    foreach (var item in sortedList)
                    {
                        var itemDict = (IDictionary<string, object>)item;
                        string catchNo = itemDict["CatchNo"]?.ToString();
                        var nrRow = nrData.FirstOrDefault(n => n.CatchNo == catchNo);
                        int pages = nrRow?.Pages ?? 0;
                        int totalPages = ((int)itemDict["Quantity"]) * pages;
                        string currentSymbol = resetOnSymbolChange ? (nrRow?.Symbol ?? "") : "";
                        string currentCourseName = itemDict["CourseName"]?.ToString() ?? "";
                        string bookletSerial = itemDict["BookletSerial"]?.ToString() ?? "";
                        string omrSerial = itemDict["OmrSerial"]?.ToString();
                        bool hasOmr = !string.IsNullOrWhiteSpace(omrSerial) && omrSerial != "0";
                        bool hasBooklet = !string.IsNullOrWhiteSpace(bookletSerial) && bookletSerial != "0";

                        bool courseChanged = resetOnSymbolChange && previousCourse != null && currentCourseName != previousCourse;
                        if (courseChanged)
                        {
                            boxNo = startBox;
                            runningPages = 0;
                            innerBundlingSerial = 0;
                            prevInnerBundlingKey = null;
                            prevMergeKey = null;
                        }

                        if (previousCatchForOmr != catchNo)
                        {
                            runningOmrPointer = 0;
                            previousCatchForOmr = catchNo;
                            if (hasOmr && omrSerial.Contains("-"))
                            {
                                var parts = omrSerial.Split('-');
                                if (long.TryParse(parts[0], out long omrStart)) runningOmrPointer = omrStart;
                            }
                            if (hasBooklet && bookletSerial.Contains("-"))
                            {
                                var parts = bookletSerial.Split('-');
                                if (long.TryParse(parts[0], out long bStart)) runningBooklet = bStart;
                            }
                        }

                        string innerBundlingKey = null;
                        int currentInnerBundlingSerial = 0;
                        if (InnerBundling && innerBundlingFieldNames.Any())
                        {
                            innerBundlingKey = string.Join("_", innerBundlingFieldNames.Select(fn =>
                                itemDict.ContainsKey(fn) ? itemDict[fn]?.ToString()?.Trim() ?? "" : ""));
                            if (innerBundlingKey != prevInnerBundlingKey)
                            {
                                innerBundlingSerial++;
                                prevInnerBundlingKey = innerBundlingKey;
                            }
                            currentInnerBundlingSerial = innerBundlingSerial;
                        }

                        string mergeKey = "";
                        if (boxIds.Any())
                        {
                            mergeKey = string.Join("_", boxIds.Select(fieldId =>
                            {
                                var fieldName = fields.FirstOrDefault(f => f.FieldId == fieldId)?.Name;
                                return fieldName != null && itemDict.ContainsKey(fieldName)
                                    ? itemDict[fieldName]?.ToString() ?? "" : "";
                            }));
                        }

                        bool mergeChanged = !courseChanged && prevMergeKey != null && mergeKey != prevMergeKey;
                        if (mergeChanged) { boxNo++; runningPages = 0; }

                        bool overflow = runningPages + totalPages > capacity;

                        if (overflow)
                        {
                            if (envelopeSize < 50 && envelopeSize > 0)
                            {
                                int pagesPerUnit = pages;
                                int remainingQty = (int)itemDict["Quantity"];
                                var lastBoxForCatch = finalWithBoxes
                                    .Where(b => { var d = (IDictionary<string, object>)b; return d["CatchNo"]?.ToString() == catchNo; })
                                    .OrderBy(b => { var d = (IDictionary<string, object>)b; return (int)d["End"]; })
                                    .LastOrDefault();

                                int currentStart = lastBoxForCatch != null
                                    ? (((IDictionary<string, object>)lastBoxForCatch)["End"] as int? ?? 0) + 1
                                    : (int)itemDict["Start"];

                                while (remainingQty > 0)
                                {
                                    int remainingCapacity = capacity - runningPages;
                                    int maxFittingQty = (int)Math.Floor((double)remainingCapacity / pagesPerUnit);
                                    maxFittingQty = (maxFittingQty / envelopeSize) * envelopeSize;
                                    if (maxFittingQty <= 0) { boxNo++; runningPages = 0; if (InnerBundling) { innerBundlingSerial++; currentInnerBundlingSerial = innerBundlingSerial; } continue; }

                                    int chunkQty = Math.Min(maxFittingQty, remainingQty);
                                    chunkQty = (chunkQty / envelopeSize) * envelopeSize;
                                    if (chunkQty <= 0) { boxNo++; runningPages = 0; if (InnerBundling) { innerBundlingSerial++; currentInnerBundlingSerial = innerBundlingSerial; } continue; }

                                    int chunkPages = chunkQty * pagesPerUnit;
                                    int envelopesInBox = chunkQty / envelopeSize;
                                    int start = currentStart;
                                    int end = start + envelopesInBox - 1;
                                    currentStart = end + 1;

                                    string omrRange = "", bookletRange = "";
                                    if (hasOmr) { long os = runningOmrPointer; long oe = os + chunkQty - 1; omrRange = $"{os}-{oe}"; runningOmrPointer = oe + 1; }
                                    if (hasBooklet) { long bs = runningBooklet; long be = bs + chunkQty - 1; bookletRange = $"{bs}-{be}"; runningBooklet = be + 1; }

                                    object boxNoValue = resetOnSymbolChange && !string.IsNullOrEmpty(currentSymbol)
                                        ? (object)$"{currentSymbol}-{boxNo}" : boxNo;

                                    var boxItem = new System.Dynamic.ExpandoObject();
                                    var d2 = (IDictionary<string, object>)boxItem;
                                    d2["CatchNo"] = itemDict["CatchNo"]; d2["CenterCode"] = itemDict["CenterCode"];
                                    d2["CenterSort"] = itemDict["CenterSort"]; d2["ExamTime"] = itemDict["ExamTime"];
                                    d2["ExamDate"] = itemDict["ExamDate"]; d2["Quantity"] = chunkQty;
                                    d2["NodalCode"] = itemDict["NodalCode"]; d2["NodalSort"] = itemDict["NodalSort"];
                                    d2["Route"] = itemDict["Route"]; d2["RouteSort"] = itemDict["RouteSort"];
                                    d2["TotalEnv"] = itemDict["TotalEnv"]; d2["Start"] = start; d2["End"] = end;
                                    d2["Serial"] = $"{start} to {end}"; d2["TotalPages"] = chunkPages;
                                    d2["BoxNo"] = boxNoValue; d2["OmrSerial"] = omrRange; d2["BookletSerial"] = bookletRange;
                                    d2["InnerBundlingSerial"] = currentInnerBundlingSerial;
                                    d2["CourseName"] = currentCourseName;
                                    d2["EnvelopeBreakingResultId"] = itemDict["EnvelopeBreakingResultId"];

                                    finalWithBoxes.Add(boxItem);
                                    runningPages += chunkPages;
                                    remainingQty -= chunkQty;
                                }

                                prevMergeKey = mergeKey;
                                previousCourse = currentCourseName;
                                continue;
                            }

                            boxNo++;
                            runningPages = 0;
                        }

                        runningPages += totalPages;
                        string normalOmrRange = "", normalBookletRange = "";
                        if (hasOmr) { long ns = runningOmrPointer; long ne = ns + (int)itemDict["Quantity"] - 1; normalOmrRange = $"{ns}-{ne}"; runningOmrPointer = ne + 1; }
                        if (hasBooklet) { long nbs = runningBooklet; long nbe = nbs + (int)itemDict["Quantity"] - 1; normalBookletRange = $"{nbs}-{nbe}"; runningBooklet = nbe + 1; }

                        object normalBoxNoValue = resetOnSymbolChange ? (object)$"{currentSymbol}-{boxNo}" : boxNo;

                        var normalBoxItem = new System.Dynamic.ExpandoObject();
                        var nd = (IDictionary<string, object>)normalBoxItem;
                        nd["CatchNo"] = itemDict["CatchNo"]; nd["CenterCode"] = itemDict["CenterCode"];
                        nd["CenterSort"] = itemDict["CenterSort"]; nd["ExamTime"] = itemDict["ExamTime"];
                        nd["ExamDate"] = itemDict["ExamDate"]; nd["Quantity"] = itemDict["Quantity"];
                        nd["NodalCode"] = itemDict["NodalCode"]; nd["NodalSort"] = itemDict["NodalSort"];
                        nd["Route"] = itemDict["Route"]; nd["RouteSort"] = itemDict["RouteSort"];
                        nd["TotalEnv"] = itemDict["TotalEnv"]; nd["Start"] = itemDict["Start"];
                        nd["End"] = itemDict["End"]; nd["Serial"] = itemDict["Serial"];
                        nd["TotalPages"] = totalPages; nd["BoxNo"] = normalBoxNoValue;
                        nd["OmrSerial"] = normalOmrRange; nd["BookletSerial"] = normalBookletRange;
                        nd["InnerBundlingSerial"] = currentInnerBundlingSerial;
                        nd["CourseName"] = currentCourseName;
                        nd["EnvelopeBreakingResultId"] = itemDict["EnvelopeBreakingResultId"];

                        finalWithBoxes.Add(normalBoxItem);
                        prevMergeKey = mergeKey;
                        previousCourse = currentCourseName;
                    }

                    // ✅ Each lot gets its own batch number
                    var maxBoxBatch = await _context.BoxBreakingResults
                        .Where(r => r.ProjectId == ProjectId)
                        .MaxAsync(r => (int?)r.UploadBatch) ?? 0;
                    int currentBatch = maxBoxBatch + 1;

                    var boxResults = new List<BoxBreakingResult>();
                    foreach (var item in finalWithBoxes)
                    {
                        var itemDict = (IDictionary<string, object>)item;
                        boxResults.Add(new BoxBreakingResult
                        {
                            ProjectId = ProjectId,
                            EnvelopeBreakingResultId = itemDict.ContainsKey("EnvelopeBreakingResultId") && itemDict["EnvelopeBreakingResultId"] != null
                                ? (int?)itemDict["EnvelopeBreakingResultId"] : null,
                            Start = (int)itemDict["Start"],
                            End = (int)itemDict["End"],
                            Serial = itemDict["Serial"]?.ToString(),
                            TotalPages = (int)itemDict["TotalPages"],
                            BoxNo = itemDict["BoxNo"]?.ToString(),
                            OmrSerial = itemDict["OmrSerial"]?.ToString(),
                            BookletSerial = itemDict["BookletSerial"]?.ToString(),
                            InnerBundlingSerial = itemDict["InnerBundlingSerial"] as int?,
                            Quantity = (int)itemDict["Quantity"],
                            UploadBatch = currentBatch,
                        });
                    }

                    _context.BoxBreakingResults.AddRange(boxResults);
                    foreach (var nr in nrData)
                        nr.Steps = 4;

                    await _context.SaveChangesAsync();
                    totalSaved += boxResults.Count;

                    _loggerService.LogEvent(
                        $"Lot {lot}: Saved {boxResults.Count} records, Batch {currentBatch}",
                        "BoxBreakingProcessing", LogHelper.GetTriggeredBy(User), ProjectId);

                    // ✅ Call report for this specific lot only
                    using var client = new HttpClient();
                    var url = $"{_apiSettings.BoxBreaking}?ProjectId={ProjectId}&LotNo={lot}";
                    var response = await client.GetAsync(url);

                    if (!response.IsSuccessStatusCode)
                    {
                        _loggerService.LogError($"Failed to generate report for Lot {lot}", "", nameof(BoxBreakingProcessingController));
                        return StatusCode((int)response.StatusCode, $"Failed to generate report for Lot {lot}");
                    }

                    allResults.Add(new { lot, records = boxResults.Count, batch = currentBatch });
                }

                return Ok(new
                {
                    message = "Box breaking processed lot-wise successfully",
                    totalRecords = totalSaved,
                    lots = allResults
                });
            }
            catch (Exception ex)
            {
                _loggerService.LogError("Error processing box breaking", ex.Message, nameof(BoxBreakingProcessingController));
                return StatusCode(500, new { error = ex.Message });
            }
        }

        [HttpGet("GetBoxBreakingReport")]
        public async Task<IActionResult> GetBoxBreakingReport(int ProjectId, [FromQuery] int LotNo)
        {
            try
            {
                // ✅ Get NRData for this lot to know which CatchNos belong to it
                var nrData = await _context.NRDatas
                    .Where(p => p.ProjectId == ProjectId && p.Status == true && p.LotNo == LotNo)
                    .ToListAsync();

                if (!nrData.Any())
                    return NotFound($"No NRData found for Lot {LotNo}");

                var nrCatchNos = nrData.Select(n => n.CatchNo).ToHashSet();

                // ✅ Get envelope results for this lot's catches
                var envelopeResults = await _context.EnvelopeBreakingResults
                    .Where(nr => nrCatchNos.Contains(nr.CatchNo) && nr.ProjectId == ProjectId)
                    .ToListAsync();

                var envelopeResultIds = envelopeResults.Select(e => e.Id).ToHashSet();

                // ✅ Get the latest batch
                var maxBatch = await _context.BoxBreakingResults
                    .Where(r => r.ProjectId == ProjectId)
                    .MaxAsync(r => (int?)r.UploadBatch);

                // ✅ Filter box results by latest batch AND only envelope result IDs belonging to this lot
                var boxResults = await _context.BoxBreakingResults
                    .Where(r => r.ProjectId == ProjectId
                             && r.UploadBatch == maxBatch
                             && r.EnvelopeBreakingResultId.HasValue
                             && envelopeResultIds.Contains(r.EnvelopeBreakingResultId.Value))
                    .OrderBy(r => r.Id)
                    .ToListAsync();

                if (!boxResults.Any())
                    return NotFound($"No box breaking results found for Lot {LotNo}");

                // ✅ Get only envelope results for this lot's catches
             

                var projectconfig = await _context.ProjectConfigs
                    .FirstOrDefaultAsync(p => p.ProjectId == ProjectId);

                bool resetOnSymbolChange = projectconfig?.ResetOnSymbolChange ?? false;

                var envelopeDict = envelopeResults.ToDictionary(e => e.Id);
                var nrDict = nrData.ToDictionary(n => n.Id);

                var headers = new List<string>
        {
            "SerialNo","CatchNo","CenterCode","CenterSort","ExamTime","ExamDate","Quantity",
            "NodalCode","NodalSort","Route","RouteSort","TotalEnv","Start","End",
            "Serial","Pages","TotalPages","BoxNo","Beejak"
        };

                if (resetOnSymbolChange) { headers.Add("Symbol"); headers.Add("CourseName"); }
                if (projectconfig.BookletSerialNumber > 0) headers.Add("BookletSerial");
                if (projectconfig.OmrSerialNumber > 0) headers.Add("OmrSerial");
                if (projectconfig.IsInnerBundlingDone) headers.Add("InnerBundlingSerial");

                var nrProperties = typeof(NRData).GetProperties().Select(p => p.Name).ToList();
                var excludedColumns = new HashSet<string>(headers) { "Id", "ProjectId", "NRDatas" };
                var extraNRColumns = nrProperties.Where(p => !excludedColumns.Contains(p)).ToList();
                headers.AddRange(extraNRColumns);

                var jsonKeys = new HashSet<string>();
                foreach (var nr in nrData)
                {
                    if (!string.IsNullOrEmpty(nr.NRDatas))
                    {
                        var dict = JsonSerializer.Deserialize<Dictionary<string, object>>(nr.NRDatas);
                        foreach (var key in dict.Keys)
                            if (!headers.Contains(key)) jsonKeys.Add(key);
                    }
                }
                headers.AddRange(jsonKeys);

                var reportPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", ProjectId.ToString());
                if (!Directory.Exists(reportPath)) Directory.CreateDirectory(reportPath);

                // ✅ One file per lot
                var fileName = $"BoxBreaking_{LotNo}.xlsx";
                var filePath = Path.Combine(reportPath, fileName);
                if (System.IO.File.Exists(filePath)) System.IO.File.Delete(filePath);

                using (var package = new ExcelPackage())
                {
                    var ws = package.Workbook.Worksheets.Add("BoxBreaking");

                    for (int i = 0; i < headers.Count; i++)
                    {
                        ws.Cells[1, i + 1].Value = headers[i];
                        ws.Cells[1, i + 1].Style.Font.Bold = true;
                    }

                    int row = 2, serial = 1;
                    var processedBoxCenters = new HashSet<string>(); // ✅ Beejak: per box+center combo

                    foreach (var result in boxResults)
                    {
                        if (!result.EnvelopeBreakingResultId.HasValue) continue;
                        if (!envelopeDict.TryGetValue(result.EnvelopeBreakingResultId.Value, out var env)) continue;

                        nrDict.TryGetValue(env.NrDataId, out var nrRow);

                        int col = 1;
                        ws.Cells[row, col++].Value = serial++;
                        ws.Cells[row, col++].Value = env.CatchNo;
                        ws.Cells[row, col++].Value = env.CenterCode;
                        ws.Cells[row, col++].Value = env.CenterSort;
                        ws.Cells[row, col++].Value = env.ExamTime;
                        ws.Cells[row, col++].Value = env.ExamDate;
                        ws.Cells[row, col++].Value = result.Quantity;
                        ws.Cells[row, col++].Value = env.NodalCode;
                        ws.Cells[row, col++].Value = env.NodalSort;
                        ws.Cells[row, col++].Value = env.Route;
                        ws.Cells[row, col++].Value = env.RouteSort;
                        ws.Cells[row, col++].Value = env.TotalEnv;
                        ws.Cells[row, col++].Value = result.Start;
                        ws.Cells[row, col++].Value = result.End;
                        ws.Cells[row, col++].Value = result.Serial;
                        ws.Cells[row, col++].Value = nrRow?.Pages ?? 0;
                        ws.Cells[row, col++].Value = result.TotalPages;
                        ws.Cells[row, col++].Value = result.BoxNo;

                        // ✅ Beejak: first time a box+center combo appears
                        string beejakKey = $"{result.BoxNo}_{env.CenterCode}";
                        ws.Cells[row, col++].Value = processedBoxCenters.Add(beejakKey) ? "Beejak" : "";

                        if (resetOnSymbolChange)
                        {
                            ws.Cells[row, col++].Value = nrRow?.Symbol;
                            ws.Cells[row, col++].Value = nrRow?.CourseName;
                        }

                        if (projectconfig.BookletSerialNumber > 0) ws.Cells[row, col++].Value = result.BookletSerial;
                        if (projectconfig.OmrSerialNumber > 0) ws.Cells[row, col++].Value = result.OmrSerial;
                        if (projectconfig.IsInnerBundlingDone) ws.Cells[row, col++].Value = result.InnerBundlingSerial;

                        foreach (var prop in extraNRColumns)
                        {
                            var val = nrRow?.GetType().GetProperty(prop)?.GetValue(nrRow);
                            ws.Cells[row, col++].Value = val;
                        }

                        Dictionary<string, object> jsonDict = null;
                        if (!string.IsNullOrEmpty(nrRow?.NRDatas))
                            jsonDict = JsonSerializer.Deserialize<Dictionary<string, object>>(nrRow.NRDatas);

                        foreach (var key in jsonKeys)
                            ws.Cells[row, col++].Value = jsonDict != null && jsonDict.ContainsKey(key) ? jsonDict[key]?.ToString() : "";

                        row++;
                    }

                    ws.Cells[ws.Dimension.Address].AutoFitColumns();
                    ws.View.FreezePanes(2, 1);
                    package.SaveAs(new FileInfo(filePath));
                }

                return Ok(new { message = "Report generated successfully", filePath, lot = LotNo });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { error = ex.Message });
            }
        }



    }
}

