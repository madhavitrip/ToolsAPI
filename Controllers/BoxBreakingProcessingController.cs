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
        public async Task<IActionResult> ProcessBoxBreaking(int ProjectId)
        {
            try
            {
                // Read EnvelopeBreakingResults from DB (latest batch)
                var maxBatch = await _context.EnvelopeBreakingResults
                    .Where(r => r.ProjectId == ProjectId)
                    .MaxAsync(r => (int?)r.UploadBatch);

                if (!maxBatch.HasValue)
                    return NotFound("No envelope breaking results found. Run ProcessEnvelopeBreaking first.");

                var envelopeResults = await _context.EnvelopeBreakingResults
                    .Where(r => r.ProjectId == ProjectId && r.UploadBatch == maxBatch.Value)
                    .ToListAsync();

                var nrData = await _context.NRDatas
                    .Where(p => p.ProjectId == ProjectId && p.Status == true)
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

                foreach (var result in envelopeResults)
                {
                    dynamic row = new System.Dynamic.ExpandoObject();
                    var rowDict = (IDictionary<string, object>)row;

                    // Get data from EnvelopeBreakingResults first (all sorting fields are already there)
                    rowDict["CatchNo"] = result.CatchNo ?? "";
                    rowDict["CenterCode"] = result.CenterCode ?? "";
                    rowDict["CenterSort"] = result.CenterSort;
                    rowDict["ExamTime"] = result.ExamTime ?? "";
                    rowDict["ExamDate"] = result.ExamDate ?? "";
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

                    // Get Pages and Symbol from NRData if available
                    var nrRow = nrData.FirstOrDefault(n => n.CatchNo == result.CatchNo);

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
                    }
                  
                    breakingReportData.Add(row);
                }

                if (!breakingReportData.Any())
                    return NotFound("No data found in envelope breaking results");

                // Step 1: Remove duplicates - keep first occurrence of each duplicate key
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
              
               
                // Step 3: Apply sorting using dictionary access (not reflection)
                IOrderedEnumerable<dynamic> ordered = null;

                foreach (var fieldName in fieldNames)
                {
                    Func<dynamic, object> keySelector = x =>
                    {
                        var dict = (IDictionary<string, object>)x;
                        if (!dict.ContainsKey(fieldName))
                        {
                            return null;
                        }
                        var val = dict[fieldName];
                        if (val == null) return null;

                        if (fieldName.Equals("NodalSort", StringComparison.OrdinalIgnoreCase))
                        {
                            if (double.TryParse(val.ToString(), out double nodalNum))
                                return nodalNum;
                            return 0.0;
                        }

                        if (fieldName.Equals("CenterSort", StringComparison.OrdinalIgnoreCase))
                        {
                            if (int.TryParse(val.ToString(), out int centerNum))
                                return centerNum;
                            return 0;
                        }

                        if (fieldName.Equals("RouteSort", StringComparison.OrdinalIgnoreCase))
                        {
                            if (int.TryParse(val.ToString(), out int routeNum))
                                return routeNum;
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
                        {
                            envelopeSize = parsedValue;
                        }
                    }
                }

                // Box breaking logic - EXACT SAME AS REPLICATION ENDPOINT
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
                    // RIGHT AFTER: var sortedList = ordered?.ToList() ?? enrichedList;
                    var itemDict = (IDictionary<string, object>)item;

                    string catchNo = itemDict["CatchNo"]?.ToString();
                    var nrRow = nrData.FirstOrDefault(n => n.CatchNo == catchNo);
                    int pages = nrRow?.Pages ?? 0;
                    int totalPages = ((int)itemDict["Quantity"]) * pages;
                    string currentSymbol = resetOnSymbolChange ? (nrRow?.Symbol ?? "") : "";
                    string currentCourseName = itemDict["CourseName"]?.ToString() ?? "";

                    // Check if this row has OMR serial data (stored in BookletSerial field)
                    string bookletSerial = itemDict["BookletSerial"]?.ToString() ?? "";
                    string omrSerial = itemDict["OmrSerial"]?.ToString();
                    bool hasOmr = !string.IsNullOrWhiteSpace(omrSerial) && omrSerial != "0";
                    bool hasBooklet = !string.IsNullOrWhiteSpace(bookletSerial) && bookletSerial != "0";
                    // Reset boxNo when CourseName changes
                    bool courseChanged = resetOnSymbolChange
    && previousCourse != null
    && currentCourseName != previousCourse;

                    if (courseChanged)
                    {
                        boxNo = startBox;
                        runningPages = 0;
                        innerBundlingSerial = 0;
                        prevInnerBundlingKey = null;
                        prevMergeKey = null; // clears stale key so first row of new course doesn't trigger boxNo++
                    }

                    if (previousCatchForOmr != itemDict["CatchNo"]?.ToString())
                    {
                        runningOmrPointer = 0;
                        previousCatchForOmr = itemDict["CatchNo"]?.ToString();
                        
                        // If this row has OMR serial, extract the starting number from BookletSerial
                        if (hasOmr && omrSerial.Contains("-"))
                        {
                            var parts = omrSerial.Split('-');
                            if (long.TryParse(parts[0], out long omrStart))
                            {
                                runningOmrPointer = omrStart;
                            }
                        }
                        if (hasBooklet && bookletSerial.Contains("-"))
                        {
                            var parts = bookletSerial.Split('-');
                            if (long.TryParse(parts[0], out long bookletStart))
                            {
                                runningBooklet = bookletStart;
                            }
                        }
                    }

                    string innerBundlingKey = null;
                    int currentInnerBundlingSerial = 0;
                    if (InnerBundling && innerBundlingFieldNames.Any())
                    {
                        innerBundlingKey = string.Join("_", innerBundlingFieldNames.Select(fieldName =>
                        {
                            return itemDict.ContainsKey(fieldName) ? itemDict[fieldName]?.ToString()?.Trim() ?? "" : "";
                        }));

                        if (innerBundlingKey != prevInnerBundlingKey)
                        {
                            innerBundlingSerial++;
                            prevInnerBundlingKey = innerBundlingKey;
                        }
                        currentInnerBundlingSerial = innerBundlingSerial;
                    }

                    // Build merge key
                    string mergeKey = "";
                    if (boxIds.Any())
                    {
                        mergeKey = string.Join("_", boxIds.Select(fieldId =>
                        {
                            var fieldName = fields.FirstOrDefault(f => f.FieldId == fieldId)?.Name;
                            if (fieldName != null && itemDict.ContainsKey(fieldName))
                            {
                                return itemDict[fieldName]?.ToString() ?? "";
                            }
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
                                        if (InnerBundling)
                                        {
                                            innerBundlingSerial++;
                                            currentInnerBundlingSerial = innerBundlingSerial;
                                        }
                                        continue;
                                    }

                                    int chunkQty = Math.Min(maxFittingQty, remainingQty);
                                    chunkQty = (chunkQty / envelopeSize) * envelopeSize;

                                    if (chunkQty <= 0)
                                    {
                                        boxNo++;
                                        runningPages = 0;
                                        if (InnerBundling)
                                        {
                                            innerBundlingSerial++;
                                            currentInnerBundlingSerial = innerBundlingSerial;
                                        }
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
      ? (object)$"{currentSymbol}-{boxNo}"
      : boxNo;
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
    ? (object)$"{currentSymbol}-{boxNo}"   // e.g. "A-1"
    : boxNo;
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

                // Get current batch number
                // RIGHT AFTER: the foreach (var item in sortedList) loop ends
               
                var maxBoxBatch = await _context.BoxBreakingResults
                    .Where(r => r.ProjectId == ProjectId)
                    .MaxAsync(r => (int?)r.UploadBatch) ?? 0;
                int currentBatch = maxBoxBatch + 1;

                // Save to database
                var boxResults = new List<BoxBreakingResult>();
                int serialNumber = 1;

                foreach (var item in finalWithBoxes)
                {
                    var itemDict = (IDictionary<string, object>)item;

                    // Store BoxNo in symbol-boxno format when resetOnSymbolChange is true
                    string boxNoValue = itemDict["BoxNo"]?.ToString() ?? "";
                    string boxNoForStorage = boxNoValue;
                 
                    boxResults.Add(new BoxBreakingResult
                    {
                        ProjectId = ProjectId,
                        EnvelopeBreakingResultId = itemDict.ContainsKey("EnvelopeBreakingResultId") && itemDict["EnvelopeBreakingResultId"] != null
                            ? (int?)itemDict["EnvelopeBreakingResultId"]
                            : null,
                        Start = (int)itemDict["Start"],
                        End = (int)itemDict["End"],
                        Serial = itemDict["Serial"]?.ToString(),
                        TotalPages = (int)itemDict["TotalPages"],
                        BoxNo = boxNoForStorage,
                        OmrSerial = itemDict["OmrSerial"]?.ToString(),
                        BookletSerial = itemDict["BookletSerial"]?.ToString(),
                        InnerBundlingSerial = itemDict["InnerBundlingSerial"] as int?,
                        Quantity = (int)itemDict["Quantity"],
                        UploadBatch = currentBatch
                    });
                }
             
                _context.BoxBreakingResults.AddRange(boxResults);
                await _context.SaveChangesAsync();

                _loggerService.LogEvent(
                    $"Saved {boxResults.Count} box breaking results for ProjectId {ProjectId}, Batch {currentBatch}",
                    "BoxBreakingProcessing",
                    User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0,
                    ProjectId);

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
        public async Task<IActionResult> GetBoxBreakingReport(int ProjectId, int UploadBatch)
        {
            try
            {
                var boxResults = await _context.BoxBreakingResults
                    .Where(r => r.ProjectId == ProjectId && r.UploadBatch == UploadBatch)
                    .OrderBy(r => r.Id)
                    .ToListAsync();

                if (!boxResults.Any())
                    return NotFound("No box breaking results found");

                var envelopeResults = await _context.EnvelopeBreakingResults.ToListAsync();

                var nrData = await _context.NRDatas
                    .Where(p => p.ProjectId == ProjectId && p.Status == true)
                    .ToListAsync();

                var projectconfig = await _context.ProjectConfigs
                    .Where(p => p.ProjectId == ProjectId)
                    .FirstOrDefaultAsync();

                bool resetOnSymbolChange = projectconfig?.ResetOnSymbolChange ?? false;

                var reportPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", ProjectId.ToString());
                if (!Directory.Exists(reportPath))
                {
                    Directory.CreateDirectory(reportPath);
                }

                var filename = "BoxBreakingReport.xlsx";
                var filePath = Path.Combine(reportPath, filename);

                if (System.IO.File.Exists(filePath))
                {
                    System.IO.File.Delete(filePath);
                }

                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("BoxBreaking");

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

                    if(projectconfig.BookletSerialNumber > 0)
                    {
                        headers.Add("BookletSerial");
                    }

                    if(projectconfig.OmrSerialNumber > 0)
                    {
                        headers.Add("OmrSerial");
                    }

                    if (projectconfig.IsInnerBundlingDone)
                    {
                        headers.Add("InnerBundlingSerial");
                    }

                    // Write headers
                    for (int i = 0; i < headers.Count; i++)
                    {
                        worksheet.Cells[1, i + 1].Value = headers[i];
                        worksheet.Cells[1, i + 1].Style.Font.Bold = true;
                    }

                    int row = 2;
                    int serial = 1;

                    // Track first center box
                    var processedCenters = new HashSet<string>();

                    foreach (var result in boxResults)
                    {
                        var envResult = envelopeResults.FirstOrDefault(e => e.Id == result.EnvelopeBreakingResultId);
                        if (envResult == null)
                            continue;

                        var nrRow = nrData.FirstOrDefault(n => n.Id == envResult.NrDataId);

                        int col = 1;

                        worksheet.Cells[row, col++].Value = serial++;
                        worksheet.Cells[row, col++].Value = envResult.CatchNo;
                        worksheet.Cells[row, col++].Value = envResult.CenterCode;
                        worksheet.Cells[row, col++].Value = envResult.CenterSort;
                        worksheet.Cells[row, col++].Value = envResult.ExamTime;
                        worksheet.Cells[row, col++].Value = envResult.ExamDate;
                        worksheet.Cells[row, col++].Value = result.Quantity;
                        worksheet.Cells[row, col++].Value = envResult.NodalCode;
                        worksheet.Cells[row, col++].Value = envResult.NodalSort;
                        worksheet.Cells[row, col++].Value = envResult.Route;
                        worksheet.Cells[row, col++].Value = envResult.RouteSort;
                        worksheet.Cells[row, col++].Value = envResult.TotalEnv;
                        worksheet.Cells[row, col++].Value = result.Start;
                        worksheet.Cells[row, col++].Value = result.End;
                        worksheet.Cells[row, col++].Value = result.Serial;
                        worksheet.Cells[row, col++].Value = nrRow?.Pages ?? 0;
                        worksheet.Cells[row, col++].Value = result.TotalPages;

                        string boxNoDisplay = result.BoxNo;
                        worksheet.Cells[row, col++].Value = boxNoDisplay;


                        // Beejak logic
                        bool isFirstBoxOfCentre = false;

                        if (!processedCenters.Contains(envResult.CenterCode))
                        {
                            processedCenters.Add(envResult.CenterCode);
                            isFirstBoxOfCentre = true;
                        }

                        worksheet.Cells[row, col++].Value = isFirstBoxOfCentre ? "Beejak" : "";

                        if (resetOnSymbolChange)
                        {
                            worksheet.Cells[row, col++].Value = nrRow?.Symbol ?? "";
                            worksheet.Cells[row, col++].Value = nrRow?.CourseName ?? "";
                        }

                        if (projectconfig.BookletSerialNumber >0) {
                            worksheet.Cells[row, col++].Value = result.BookletSerial;
                        }

                        if (projectconfig.OmrSerialNumber > 0) {
                            worksheet.Cells[row, col++].Value = result.OmrSerial;
                        }

                        if (projectconfig.IsInnerBundlingDone)
                        {
                            worksheet.Cells[row, col++].Value = result.InnerBundlingSerial;
                        }

                        row++;
                    }

                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                    worksheet.View.FreezePanes(2, 1);

                    FileInfo fi = new FileInfo(filePath);
                    package.SaveAs(fi);
                }

                _loggerService.LogEvent(
                    $"Generated box breaking report for ProjectId {ProjectId}",
                    "BoxBreakingProcessing",
                    User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0,
                    ProjectId);

                return Ok(new { message = "Report generated successfully", filePath });
            }
            catch (Exception ex)
            {
                _loggerService.LogError("Error generating box breaking report", ex.Message, nameof(BoxBreakingProcessingController));
                return StatusCode(500, new { error = ex.Message });
            }
        }
    }
}
