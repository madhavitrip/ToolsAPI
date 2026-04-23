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

                var envelopeResults = await _context.EnvelopeBreakingResults
                    .Where(r => r.ProjectId == ProjectId && r.UploadBatch == maxBatch.Value && r.SerialNumber != 0)
                    .ToListAsync();

                var nrData = await _context.NRDatas

                    .Where(p => p.ProjectId == ProjectId && p.Status == true && LotNo.Contains(p.LotNo))
                    .ToListAsync();

                var projectconfig = await _context.ProjectConfigs
                    .Where(p => p.ProjectId == ProjectId)
                    .FirstOrDefaultAsync();

                if (projectconfig == null)
                    return NotFound("Project config not found");

                var boxIds = projectconfig.BoxBreakingCriteria;
                var sortingId = projectconfig.SortingBoxReport;
                var duplicatesFields = projectconfig.DuplicateRemoveFields;
                bool InnerBundling = projectconfig.IsInnerBundlingDone;
                var innerBundlingFieldNames = new List<string>();

                if (InnerBundling && projectconfig.InnerBundlingCriteria != null)
                {
                    var InnerBCriteria = projectconfig.InnerBundlingCriteria;
                    innerBundlingFieldNames = await _context.Fields
                        .Where(f => InnerBCriteria.Contains(f.FieldId))
                        .Select(f => f.Name)
                        .ToListAsync();
                }

                var fields = await _context.Fields
                    .Where(f => boxIds.Contains(f.FieldId))
                    .ToListAsync();

                var fieldsFromDb = await _context.Fields
                    .Where(f => sortingId.Contains(f.FieldId))
                    .ToListAsync();

                var fieldNames = fieldsFromDb
                    .OrderBy(f => sortingId.IndexOf(f.FieldId))
                    .Select(f => f.Name)
                    .ToList();

                var dupNames = await _context.Fields
                    .Where(f => duplicatesFields.Contains(f.FieldId))
                    .Select(f => f.Name)
                    .ToListAsync();

                var startBox = projectconfig.BoxNumber;
                bool resetOnSymbolChange = projectconfig.ResetOnSymbolChange;

                var capacity = await _context.BoxCapacity
                    .Where(c => c.BoxCapacityId == projectconfig.BoxCapacity)
                    .Select(c => c.Capacity)
                    .FirstOrDefaultAsync();

                // ==============================
                // Read from EnvelopeBreakingResults and build data rows
                // ==============================
                var breakingReportData = new List<dynamic>();

                // ✅ HELPER FUNCTION — placed here so fieldNames, dupNames, fields are all in scope
                // Deserializes the NRDatas JSON column once per row and returns a lookup function
                Dictionary<string, JsonElement> ParseNrDatas(string nrDatasJson)
                {
                    if (string.IsNullOrWhiteSpace(nrDatasJson)) return null;
                    try
                    {
                        return JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(nrDatasJson);
                    }
                    catch
                    {
                        return null;
                    }
                }

                // ✅ Given a parsed NRDatas dictionary, get a field value by name (case-insensitive)
                string GetNrField(Dictionary<string, JsonElement> nrDynamic, string fieldName)
                {
                    if (nrDynamic == null) return "";
                    var match = nrDynamic.FirstOrDefault(k =>
                        k.Key.Equals(fieldName, StringComparison.OrdinalIgnoreCase));
                    return match.Key != null ? (match.Value.GetString() ?? "") : "";
                }

                // ✅ All field names that might live in NRDatas JSON (union of sorting + dup + box criteria)
                var boxFieldNames = fields.Select(f => f.Name).ToList();
                var allDynamicFieldNames = fieldNames
                    .Union(dupNames)
                    .Union(boxFieldNames)
                    .Union(innerBundlingFieldNames)
                    .Distinct()
                    .ToList();

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

                    // ✅ Parse NRDatas JSON ONCE here for this row
                    var nrDynamic = nrRow != null ? ParseNrDatas(nrRow.NRDatas) : null;

                    if (nrRow != null)
                    {
                        rowDict["Symbol"] = nrRow.Symbol ?? "";
                        rowDict["Pages"] = nrRow.Pages;
                        rowDict["NRQuantity"] = nrRow.NRQuantity;
                    }
                    else
                    {
                        rowDict["Symbol"] = "";
                        rowDict["Pages"] = 0;
                        rowDict["NRQuantity"] = 0;
                    }

                    // ✅ Fill in any missing fields (Remark, PaperCode, etc.) from NRDatas JSON
                    // This covers sorting fields, duplicate fields, box criteria fields, inner bundling fields
                    foreach (var fieldName in allDynamicFieldNames)
                    {
                        if (!rowDict.ContainsKey(fieldName))
                        {
                            rowDict[fieldName] = GetNrField(nrDynamic, fieldName);
                        }
                    }

                    breakingReportData.Add(row);
                }

                if (!breakingReportData.Any())
                    return NotFound("No data found in envelope breaking results");

                // Step 1: Remove duplicates
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

                // Step 2: Calculate Start, End, Serial (BEFORE sorting)
                var enrichedList = new List<dynamic>();
                string previousCatchNo = null;
                int previousEnd = 0;

                foreach (var row in uniqueRows)
                {
                    var rowDict = (IDictionary<string, object>)row;
                    string catchNo = rowDict["CatchNo"]?.ToString();

                    int start = catchNo != previousCatchNo ? 1 : previousEnd + 1;
                    int end = start + (int)rowDict["TotalEnv"] - 1;
                    string serial = $"{start} to {end}";

                    rowDict["Start"] = start;
                    rowDict["End"] = end;
                    rowDict["Serial"] = serial;

                    enrichedList.Add(row);

                    previousCatchNo = catchNo;
                    previousEnd = end;
                }

                // Step 3: Apply sorting
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
                        {
                            if (double.TryParse(val.ToString(), out double nodalNum)) return nodalNum;
                            return 0.0;
                        }
                        if (fieldName.Equals("CenterSort", StringComparison.OrdinalIgnoreCase))
                        {
                            if (int.TryParse(val.ToString(), out int centerNum)) return centerNum;
                            return 0;
                        }
                        if (fieldName.Equals("RouteSort", StringComparison.OrdinalIgnoreCase))
                        {
                            if (int.TryParse(val.ToString(), out int routeNum)) return routeNum;
                            return 0;
                        }
                        if (fieldName.Equals("ExamDate", StringComparison.OrdinalIgnoreCase))
                        {
                            if (DateTime.TryParseExact(val.ToString(), "dd-MM-yyyy",
                                System.Globalization.CultureInfo.InvariantCulture,
                                System.Globalization.DateTimeStyles.None, out DateTime parsedDate))
                                return parsedDate;
                        }

                        return val.ToString().Trim().ToLowerInvariant();
                    };

                    if (ordered == null)
                        ordered = enrichedList.OrderBy(keySelector);
                    else
                        ordered = ordered.ThenBy(keySelector);
                }

                var sortedList = ordered?.ToList() ?? enrichedList;

                // Get envelope size from config
                var envelopeObj = JsonSerializer.Deserialize<Dictionary<string, string>>(projectconfig.Envelope);
                int envelopeSize = 0;
                if (envelopeObj != null && envelopeObj.ContainsKey("Outer"))
                {
                    string outerEnvValue = envelopeObj["Outer"];
                    var outerParts = outerEnvValue.Split(',', StringSplitOptions.RemoveEmptyEntries);
                    if (outerParts.Length == 1)
                    {
                        string singleValue = outerParts[0].Trim();
                        string digits = new string(singleValue.Where(char.IsDigit).ToArray());
                        if (int.TryParse(digits, out int parsedValue) && parsedValue > 0)
                            envelopeSize = parsedValue;
                    }
                }

                // Box breaking logic
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

                    bool courseChanged = resetOnSymbolChange
                        && previousCourse != null
                        && currentCourseName != previousCourse;

                    if (courseChanged)
                    {
                        boxNo = startBox;
                        runningPages = 0;
                        innerBundlingSerial = 0;
                        prevInnerBundlingKey = null;
                        prevMergeKey = null;
                    }

                    if (previousCatchForOmr != itemDict["CatchNo"]?.ToString())
                    {
                        runningOmrPointer = 0;
                        previousCatchForOmr = itemDict["CatchNo"]?.ToString();

                        if (hasOmr && omrSerial.Contains("-"))
                        {
                            var parts = omrSerial.Split('-');
                            if (long.TryParse(parts[0], out long omrStart))
                                runningOmrPointer = omrStart;
                        }
                        if (hasBooklet && bookletSerial.Contains("-"))
                        {
                            var parts = bookletSerial.Split('-');
                            if (long.TryParse(parts[0], out long bookletStart))
                                runningBooklet = bookletStart;
                        }
                    }

                    string innerBundlingKey = null;
                    int currentInnerBundlingSerial = 0;
                    if (InnerBundling && innerBundlingFieldNames.Any())
                    {
                        innerBundlingKey = string.Join("_", innerBundlingFieldNames.Select(fieldName =>
                            itemDict.ContainsKey(fieldName) ? itemDict[fieldName]?.ToString()?.Trim() ?? "" : ""));

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
                            if (fieldName != null && itemDict.ContainsKey(fieldName))
                                return itemDict[fieldName]?.ToString() ?? "";
                            return "";
                        }));
                    }

                    bool mergeChanged = !courseChanged && (prevMergeKey != null && mergeKey != prevMergeKey);

                    try
                    {
                        if (mergeChanged)
                        {
                            boxNo++;
                            runningPages = 0;
                        }

                        bool overflow = (runningPages + totalPages > capacity);

                        if (overflow)
                        {
                            if (envelopeSize < 50 && envelopeSize > 0)
                            {
                                int pagesPerUnit = pages;
                                int remainingQty = (int)itemDict["Quantity"];

                                var lastBoxForCatch = finalWithBoxes
                                    .Where(b =>
                                    {
                                        var dict = (IDictionary<string, object>)b;
                                        return dict["CatchNo"]?.ToString() == itemDict["CatchNo"]?.ToString();
                                    })
                                    .OrderBy(b =>
                                    {
                                        var dict = (IDictionary<string, object>)b;
                                        return (int)dict["End"];
                                    })
                                    .LastOrDefault();

                                int currentStart = lastBoxForCatch != null
                                    ? (((IDictionary<string, object>)lastBoxForCatch)["End"] as int? ?? 0) + 1
                                    : (int)itemDict["Start"];

                                while (remainingQty > 0)
                                {
                                    int remainingCapacity = capacity - runningPages;
                                    int maxFittingQty = (int)Math.Floor((double)remainingCapacity / pagesPerUnit);
                                    maxFittingQty = (maxFittingQty / envelopeSize) * envelopeSize;

                                    if (maxFittingQty <= 0)
                                    {
                                        boxNo++;
                                        runningPages = 0;
                                        if (InnerBundling) { innerBundlingSerial++; currentInnerBundlingSerial = innerBundlingSerial; }
                                        continue;
                                    }

                                    int chunkQty = Math.Min(maxFittingQty, remainingQty);
                                    chunkQty = (chunkQty / envelopeSize) * envelopeSize;

                                    if (chunkQty <= 0)
                                    {
                                        boxNo++;
                                        runningPages = 0;
                                        if (InnerBundling) { innerBundlingSerial++; currentInnerBundlingSerial = innerBundlingSerial; }
                                        continue;
                                    }

                                    int chunkPages = chunkQty * pagesPerUnit;
                                    int envelopesInBox = chunkQty / envelopeSize;
                                    int start = currentStart;
                                    int end = start + envelopesInBox - 1;
                                    currentStart = end + 1;
                                    string serial = $"{start} to {end}";
                                    string omrRange = "";
                                    string bookletRange = "";

                                    if (hasOmr)
                                    {
                                        long omrStart = runningOmrPointer;
                                        long omrEnd = omrStart + chunkQty - 1;
                                        omrRange = $"{omrStart}-{omrEnd}";
                                        runningOmrPointer = omrEnd + 1;
                                    }
                                    if (hasBooklet)
                                    {
                                        long bookletStart = runningBooklet;
                                        long bookletEnd = bookletStart + chunkQty - 1;
                                        bookletRange = $"{bookletStart}-{bookletEnd}";
                                        runningBooklet = bookletEnd + 1;
                                    }

                                    object boxNoValue = (resetOnSymbolChange && !string.IsNullOrEmpty(currentSymbol))
                                        ? (object)$"{currentSymbol}-{boxNo}" : boxNo;

                                    var boxItem = new System.Dynamic.ExpandoObject();
                                    var boxItemDict = (IDictionary<string, object>)boxItem;

                                    boxItemDict["CatchNo"] = itemDict["CatchNo"];
                                    boxItemDict["CenterCode"] = itemDict["CenterCode"];
                                    boxItemDict["CenterSort"] = itemDict["CenterSort"];
                                    boxItemDict["ExamTime"] = itemDict["ExamTime"];
                                    boxItemDict["ExamDate"] = itemDict["ExamDate"];
                                    boxItemDict["Quantity"] = chunkQty;
                                    boxItemDict["NodalCode"] = itemDict["NodalCode"];
                                    boxItemDict["NodalSort"] = itemDict["NodalSort"];
                                    boxItemDict["Route"] = itemDict["Route"];
                                    boxItemDict["RouteSort"] = itemDict["RouteSort"];
                                    boxItemDict["TotalEnv"] = itemDict["TotalEnv"];
                                    boxItemDict["Start"] = start;
                                    boxItemDict["End"] = end;
                                    boxItemDict["Serial"] = serial;
                                    boxItemDict["TotalPages"] = chunkPages;
                                    boxItemDict["BoxNo"] = boxNoValue;
                                    boxItemDict["OmrSerial"] = omrRange;
                                    boxItemDict["BookletSerial"] = bookletRange;
                                    boxItemDict["InnerBundlingSerial"] = currentInnerBundlingSerial;
                                    boxItemDict["CourseName"] = currentCourseName;
                                    boxItemDict["EnvelopeBreakingResultId"] = itemDict["EnvelopeBreakingResultId"];

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

                        // Normal case
                        runningPages += totalPages;
                        string normalOmrRange = "";
                        string normalBookletRange = "";

                        if (hasOmr)
                        {
                            long normalOmrStart = runningOmrPointer;
                            long normalOmrEnd = normalOmrStart + (int)itemDict["Quantity"] - 1;
                            normalOmrRange = $"{normalOmrStart}-{normalOmrEnd}";
                            runningOmrPointer = normalOmrEnd + 1;
                        }
                        if (hasBooklet)
                        {
                            long normalBookletStart = runningBooklet;
                            long normalBookletEnd = normalBookletStart + (int)itemDict["Quantity"] - 1;
                            normalBookletRange = $"{normalBookletStart}-{normalBookletEnd}";
                            runningBooklet = normalBookletEnd + 1;
                        }

                        object normalBoxNoValue = resetOnSymbolChange
                            ? (object)$"{currentSymbol}-{boxNo}" : boxNo;

                        var normalBoxItem = new System.Dynamic.ExpandoObject();
                        var normalBoxItemDict = (IDictionary<string, object>)normalBoxItem;

                        normalBoxItemDict["CatchNo"] = itemDict["CatchNo"];
                        normalBoxItemDict["CenterCode"] = itemDict["CenterCode"];
                        normalBoxItemDict["CenterSort"] = itemDict["CenterSort"];
                        normalBoxItemDict["ExamTime"] = itemDict["ExamTime"];
                        normalBoxItemDict["ExamDate"] = itemDict["ExamDate"];
                        normalBoxItemDict["Quantity"] = itemDict["Quantity"];
                        normalBoxItemDict["NodalCode"] = itemDict["NodalCode"];
                        normalBoxItemDict["NodalSort"] = itemDict["NodalSort"];
                        normalBoxItemDict["Route"] = itemDict["Route"];
                        normalBoxItemDict["RouteSort"] = itemDict["RouteSort"];
                        normalBoxItemDict["TotalEnv"] = itemDict["TotalEnv"];
                        normalBoxItemDict["Start"] = itemDict["Start"];
                        normalBoxItemDict["End"] = itemDict["End"];
                        normalBoxItemDict["Serial"] = itemDict["Serial"];
                        normalBoxItemDict["TotalPages"] = totalPages;
                        normalBoxItemDict["BoxNo"] = normalBoxNoValue;
                        normalBoxItemDict["OmrSerial"] = normalOmrRange;
                        normalBoxItemDict["BookletSerial"] = normalBookletRange;
                        normalBoxItemDict["InnerBundlingSerial"] = currentInnerBundlingSerial;
                        normalBoxItemDict["CourseName"] = currentCourseName;
                        normalBoxItemDict["EnvelopeBreakingResultId"] = itemDict["EnvelopeBreakingResultId"];

                        finalWithBoxes.Add(normalBoxItem);

                        prevMergeKey = mergeKey;
                        previousCourse = currentCourseName;
                    }
                    catch (Exception ex)
                    {
                        _loggerService.LogError("Error in box breaking logic", ex.Message, nameof(BoxBreakingProcessingController));
                        throw;
                    }
                }

                var maxBoxBatch = await _context.BoxBreakingResults
                    .Where(r => r.ProjectId == ProjectId)
                    .MaxAsync(r => (int?)r.UploadBatch) ?? 0;
                int currentBatch = maxBoxBatch + 1;

                var boxResults = new List<BoxBreakingResult>();
                int serialNumber = 1;

                foreach (var item in finalWithBoxes)
                {
                    var itemDict = (IDictionary<string, object>)item;
                    string boxNoValue = itemDict["BoxNo"]?.ToString() ?? "";

                    boxResults.Add(new BoxBreakingResult
                    {
                        ProjectId = ProjectId,
                        EnvelopeBreakingResultId = itemDict.ContainsKey("EnvelopeBreakingResultId") && itemDict["EnvelopeBreakingResultId"] != null
                            ? (int?)itemDict["EnvelopeBreakingResultId"] : null,
                        Start = (int)itemDict["Start"],
                        End = (int)itemDict["End"],
                        Serial = itemDict["Serial"]?.ToString(),
                        TotalPages = (int)itemDict["TotalPages"],
                        BoxNo = boxNoValue,
                        OmrSerial = itemDict["OmrSerial"]?.ToString(),
                        BookletSerial = itemDict["BookletSerial"]?.ToString(),
                        InnerBundlingSerial = itemDict["InnerBundlingSerial"] as int?,
                        Quantity = (int)itemDict["Quantity"],
                        UploadBatch = currentBatch
                    });
                }

                _context.BoxBreakingResults.AddRange(boxResults);
                foreach (var nr in nrData)
                {
                    nr.Steps = 5; // Assuming NRData has a Step property
                }

                await _context.SaveChangesAsync();

                _loggerService.LogEvent(
                    $"Saved {boxResults.Count} box breaking results for ProjectId {ProjectId}, Batch {currentBatch}",
                    "BoxBreakingProcessing",
                    LogHelper.GetTriggeredBy(User),
                    ProjectId);

                using var client = new HttpClient();
                var query = string.Join("&", LotNo.Select(l => $"LotNo={Uri.EscapeDataString(l.ToString())}"));
                var url = $"{_apiSettings.BoxBreaking}?ProjectId={ProjectId}&{query}";
                var response = await client.GetAsync(url);

                _loggerService.LogError("Calling URL", url, nameof(BoxBreakingProcessingController));
                if (!response.IsSuccessStatusCode)
                {
                    _loggerService.LogError("Failed to generate report", "", nameof(BoxBreakingProcessingController));
                    return StatusCode((int)response.StatusCode, "Failed to get envelope breakages after configuration.");
                }

                return Ok(new
                {
                    message = "Box breaking data saved to database",
                    recordsCount = boxResults.Count,
                    uploadBatch = currentBatch
                });
            }
            catch (Exception ex)
            {
                _loggerService.LogError("Error processing box breaking", ex.Message, nameof(BoxBreakingProcessingController));
                return StatusCode(500, new { error = ex.Message });
            }
        }

        /// <summary>
        /// GET: Retrieve BoxBreaking data from database and generate Excel
        /// </summary>

        [HttpGet("GetBoxBreakingReport")]
        public async Task<IActionResult> GetBoxBreakingReport(int ProjectId, [FromQuery] List<int> LotNo)
        {
            try
            {
                if (LotNo == null || !LotNo.Any())
                {
                    return BadRequest("LotNo is empty");
                }
                var maxBatch = await _context.BoxBreakingResults
                   .Where(r => r.ProjectId == ProjectId)
                   .MaxAsync(r => (int?)r.UploadBatch);
                var boxResults = await _context.BoxBreakingResults
                    .Where(r => r.ProjectId == ProjectId && r.UploadBatch == maxBatch)
                    .OrderBy(r => r.Id)
                    .ToListAsync();

                if (!boxResults.Any())
                    return NotFound("No box breaking results found");

                var envelopeResults = await _context.EnvelopeBreakingResults.ToListAsync();
                var nrData = await _context.NRDatas
                    .Where(p => p.ProjectId == ProjectId && p.Status == true && LotNo.Contains(p.LotNo))
                    .ToListAsync();

                var projectconfig = await _context.ProjectConfigs
                    .FirstOrDefaultAsync(p => p.ProjectId == ProjectId);

                bool resetOnSymbolChange = projectconfig?.ResetOnSymbolChange ?? false;

                // ?? Optimize lookups
                var envelopeDict = envelopeResults.ToDictionary(e => e.Id);
                var nrDict = nrData.ToDictionary(n => n.Id);

                // ?? Fixed Columns
                var headers = new List<string>
        {
            "SerialNo","CatchNo","CenterCode","CenterSort","ExamTime","ExamDate","Quantity",
            "NodalCode","NodalSort","Route","RouteSort","TotalEnv","Start","End",
            "Serial","Pages","TotalPages","BoxNo","Beejak"
        };

                if (resetOnSymbolChange)
                {
                    headers.Add("Symbol");
                    headers.Add("CourseName");
                }

                if (projectconfig.BookletSerialNumber > 0)
                    headers.Add("BookletSerial");

                if (projectconfig.OmrSerialNumber > 0)
                    headers.Add("OmrSerial");

                if (projectconfig.IsInnerBundlingDone)
                    headers.Add("InnerBundlingSerial");

                // ?? Get ALL remaining NRData columns dynamically
                var nrProperties = typeof(NRData).GetProperties()
                    .Select(p => p.Name)
                    .ToList();

                // ?? Exclude already used + unwanted
                var excludedColumns = new HashSet<string>(headers)
        {
            "Id","ProjectId","NRDatas" // handled separately
        };

                var extraNRColumns = nrProperties
                    .Where(p => !excludedColumns.Contains(p))
                    .ToList();

                headers.AddRange(extraNRColumns);

                // ?? Extract JSON keys

                var jsonKeys = new HashSet<string>();

                foreach (var nr in nrData)
                {
                    if (!string.IsNullOrEmpty(nr.NRDatas))
                    {
                        var dict = JsonSerializer.Deserialize<Dictionary<string, object>>(nr.NRDatas);
                        foreach (var key in dict.Keys)
                        {
                            if (!headers.Contains(key))
                                jsonKeys.Add(key);
                        }
                    }
                }

                headers.AddRange(jsonKeys);

                // ?? Excel Generation
                var reportPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", ProjectId.ToString());
                if (!Directory.Exists(reportPath))
                    Directory.CreateDirectory(reportPath);
                var lotNoPart = string.Join("_", LotNo);
                var fileName = $"BoxBreaking_{lotNoPart}.xlsx";
                var filePath = Path.Combine(reportPath, fileName);
                if (System.IO.File.Exists(filePath))
                    System.IO.File.Delete(filePath);

                using (var package = new ExcelPackage())
                {
                    var ws = package.Workbook.Worksheets.Add("BoxBreaking");

                    // Header
                    for (int i = 0; i < headers.Count; i++)
                    {
                        ws.Cells[1, i + 1].Value = headers[i];
                        ws.Cells[1, i + 1].Style.Font.Bold = true;
                    }

                    int row = 2;
                    int serial = 1;
                    var processedCenters = new HashSet<string>();

                    foreach (var result in boxResults)
                    {
                        if (!result.EnvelopeBreakingResultId.HasValue)
                            continue;

                        if (!envelopeDict.TryGetValue(result.EnvelopeBreakingResultId.Value, out var env))
                            continue;


                        nrDict.TryGetValue(env.NrDataId, out var nrRow);

                        int col = 1;

                        // ?? Fixed values
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

                        bool isFirst = processedCenters.Add(env.CenterCode);
                        ws.Cells[row, col++].Value = isFirst ? "Beejak" : "";

                        if (resetOnSymbolChange)
                        {
                            ws.Cells[row, col++].Value = nrRow?.Symbol;
                            ws.Cells[row, col++].Value = nrRow?.CourseName;
                        }

                        if (projectconfig.BookletSerialNumber > 0)
                            ws.Cells[row, col++].Value = result.BookletSerial;

                        if (projectconfig.OmrSerialNumber > 0)
                            ws.Cells[row, col++].Value = result.OmrSerial;

                        if (projectconfig.IsInnerBundlingDone)
                            ws.Cells[row, col++].Value = result.InnerBundlingSerial;

                        // ?? Extra NRData columns
                        foreach (var prop in extraNRColumns)
                        {
                            var val = nrRow?.GetType().GetProperty(prop)?.GetValue(nrRow);
                            ws.Cells[row, col++].Value = val;
                        }

                        // ?? JSON columns
                        Dictionary<string, object> jsonDict = null;

                        if (!string.IsNullOrEmpty(nrRow?.NRDatas))
                        {
                            jsonDict = JsonSerializer.Deserialize<Dictionary<string, object>>(nrRow.NRDatas);
                        }

                        foreach (var key in jsonKeys)
                        {
                            ws.Cells[row, col++].Value =
                                (jsonDict != null && jsonDict.ContainsKey(key))
                                ? jsonDict[key]?.ToString()
                                : "";
                        }

                        row++;
                    }

                    ws.Cells[ws.Dimension.Address].AutoFitColumns();
                    ws.View.FreezePanes(2, 1);

                    package.SaveAs(new FileInfo(filePath));
                }

                return Ok(new { message = "Report generated successfully", filePath });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { error = ex.Message });
            }
        }
    }
}

