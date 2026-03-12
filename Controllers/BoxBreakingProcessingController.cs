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

        /// <summary>
        /// POST: Process and store BoxBreaking data to database
        /// This mimics the Replication GET endpoint but saves to DB first
        /// </summary>
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
                    .Where(p => p.ProjectId == ProjectId)
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
                // Read from EnvelopeBreakingResults
                // ==============================
                var breakingReportData = new List<dynamic>();

                foreach (var result in envelopeResults)
                {
                    dynamic row = new System.Dynamic.ExpandoObject();
                    var rowDict = (IDictionary<string, object>)row;

                    rowDict["CatchNo"] = result.CatchNo;
                    rowDict["CenterCode"] = result.CenterCode;
                    rowDict["CenterSort"] = result.CenterSort;
                    rowDict["ExamTime"] = result.ExamTime;
                    rowDict["ExamDate"] = result.ExamDate;
                    rowDict["Quantity"] = result.Quantity;
                    rowDict["TotalEnv"] = result.TotalEnv;
                    rowDict["NRQuantity"] = result.NRQuantity;
                    rowDict["NodalCode"] = result.NodalCode;
                    rowDict["NodalSort"] = result.NodalSort;
                    rowDict["Route"] = result.Route;
                    rowDict["RouteSort"] = result.RouteSort;
                    rowDict["OmrSerial"] = "";
                    rowDict["CourseName"] = result.CourseName ?? "";
                    rowDict["EnvelopeBreakingResultId"] = result.Id;

                    breakingReportData.Add(row);
                }

                if (!breakingReportData.Any())
                    return NotFound("No data found in envelope breaking results");

                // Remove duplicates
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

                // Calculate Start, End, Serial
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

                // Apply sorting
                IOrderedEnumerable<dynamic> ordered = null;

                var properties = fieldNames
                    .Select(name => new
                    {
                        Name = name,
                        Property = enrichedList.First().GetType().GetProperty(name,
                            System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.IgnoreCase)
                    })
                    .Where(x => x.Property != null)
                    .ToList();

                for (int i = 0; i < properties.Count; i++)
                {
                    var prop = properties[i].Property;

                    Func<dynamic, object> keySelector = x =>
                    {
                        var dict = (IDictionary<string, object>)x;
                        var val = dict.ContainsKey(prop.Name) ? dict[prop.Name] : null;

                        if (val == null) return null;

                        if (prop.Name.Equals("NodalSort", StringComparison.OrdinalIgnoreCase))
                        {
                            if (double.TryParse(val.ToString(), out double nodalNum))
                                return nodalNum;
                            return 0.0;
                        }

                        if (prop.Name.Equals("CenterSort", StringComparison.OrdinalIgnoreCase))
                        {
                            if (int.TryParse(val.ToString(), out int centerNum))
                                return centerNum;
                            return 0;
                        }

                        if (prop.Name.Equals("RouteSort", StringComparison.OrdinalIgnoreCase))
                        {
                            if (int.TryParse(val.ToString(), out int routeNum))
                                return routeNum;
                            return 0;
                        }

                        if (prop.Name.Equals("ExamDate", StringComparison.OrdinalIgnoreCase))
                        {
                            if (DateTime.TryParseExact(val.ToString(), "dd-MM-yyyy",
                                System.Globalization.CultureInfo.InvariantCulture,
                                System.Globalization.DateTimeStyles.None, out DateTime parsedDate))
                                return parsedDate;
                        }

                        return val.ToString().Trim();
                    };

                    if (i == 0)
                        ordered = enrichedList.OrderBy(keySelector);
                    else
                        ordered = ordered.ThenBy(keySelector);
                }

                var sortedList = ordered?.ToList() ?? enrichedList;

                // Box breaking logic
                var finalWithBoxes = new List<dynamic>();
                int boxNo = startBox;
                int runningPages = 0;
                string prevMergeKey = null;
                long runningOmrPointer = 0;
                string previousCatchForOmr = null;
                string previousCourse = null;
                int innerBundlingSerial = 0;
                string prevInnerBundlingKey = null;

                foreach (var item in sortedList)
                {
                    var itemDict = (IDictionary<string, object>)item;
                    var nrRow = nrData.FirstOrDefault(n => n.CatchNo == itemDict["CatchNo"]?.ToString());
                    int pages = nrRow?.Pages ?? 0;
                    int quantity = (int)itemDict["Quantity"];
                    int totalPages = quantity * pages;
                    string currentSymbol = resetOnSymbolChange ? (nrRow?.Symbol ?? "") : "";
                    string currentCourseName = itemDict["CourseName"]?.ToString() ?? "";

                    if (previousCourse != null && currentCourseName != previousCourse)
                    {
                        boxNo = startBox;
                        runningPages = 0;
                    }

                    bool hasOmr = false;
                    if (previousCatchForOmr != itemDict["CatchNo"]?.ToString())
                    {
                        runningOmrPointer = 0;
                        previousCatchForOmr = itemDict["CatchNo"]?.ToString();
                    }

                    string innerBundlingKey = null;
                    int currentInnerBundlingSerial = 0;
                    if (InnerBundling && innerBundlingFieldNames.Any())
                    {
                        innerBundlingKey = string.Join("_", innerBundlingFieldNames.Select(fieldName =>
                        {
                            var dict = (IDictionary<string, object>)item;
                            return dict.ContainsKey(fieldName) ? dict[fieldName]?.ToString()?.Trim() ?? "" : "";
                        }));

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
                            if (fieldName != null)
                            {
                                var dict = (IDictionary<string, object>)item;
                                return dict.ContainsKey(fieldName) ? dict[fieldName]?.ToString() ?? "" : "";
                            }
                            return "";
                        }));
                    }

                    bool mergeChanged = (prevMergeKey != null && mergeKey != prevMergeKey);

                    if (mergeChanged)
                    {
                        boxNo++;
                        runningPages = 0;
                    }

                    bool overflow = (runningPages + totalPages > capacity);

                    if (overflow)
                    {
                        boxNo++;
                        runningPages = 0;
                    }

                    runningPages += totalPages;
                    string omrRange = "";

                    if (hasOmr)
                    {
                        long omrStart = runningOmrPointer;
                        long omrEnd = omrStart + quantity - 1;
                        omrRange = $"{omrStart}-{omrEnd}";
                        runningOmrPointer = omrEnd + 1;
                    }

                    object boxNoValue = resetOnSymbolChange
                        ? (object)$"{boxNo}{currentSymbol}"
                        : boxNo;

                    var boxItem = new System.Dynamic.ExpandoObject();
                    var boxItemDict = (IDictionary<string, object>)boxItem;

                    boxItemDict["CatchNo"] = itemDict["CatchNo"];
                    boxItemDict["CenterCode"] = itemDict["CenterCode"];
                    boxItemDict["CenterSort"] = itemDict["CenterSort"];
                    boxItemDict["ExamTime"] = itemDict["ExamTime"];
                    boxItemDict["ExamDate"] = itemDict["ExamDate"];
                    boxItemDict["Quantity"] = quantity;
                    boxItemDict["NodalCode"] = itemDict["NodalCode"];
                    boxItemDict["NodalSort"] = itemDict["NodalSort"];
                    boxItemDict["Route"] = itemDict["Route"];
                    boxItemDict["RouteSort"] = itemDict["RouteSort"];
                    boxItemDict["TotalEnv"] = itemDict["TotalEnv"];
                    boxItemDict["Start"] = itemDict["Start"];
                    boxItemDict["End"] = itemDict["End"];
                    boxItemDict["Serial"] = itemDict["Serial"];
                    boxItemDict["TotalPages"] = totalPages;
                    boxItemDict["BoxNo"] = boxNoValue;
                    boxItemDict["OmrSerial"] = omrRange;
                    boxItemDict["InnerBundlingSerial"] = currentInnerBundlingSerial;
                    boxItemDict["CourseName"] = currentCourseName;
                    boxItemDict["EnvelopeBreakingResultId"] = itemDict["EnvelopeBreakingResultId"];

                    finalWithBoxes.Add(boxItem);

                    prevMergeKey = mergeKey;
                    previousCourse = currentCourseName;
                }

                // Get current batch number
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

                    boxResults.Add(new BoxBreakingResult
                    {
                        ProjectId = ProjectId,
                        EnvelopeBreakingResultId = (int)itemDict["EnvelopeBreakingResultId"],
                        Start = (int)itemDict["Start"],
                        End = (int)itemDict["End"],
                        Serial = itemDict["Serial"]?.ToString(),
                        TotalPages = (int)itemDict["TotalPages"],
                        BoxNo = itemDict["BoxNo"]?.ToString(),
                        OmrSerial = itemDict["OmrSerial"]?.ToString(),
                        InnerBundlingSerial = itemDict["InnerBundlingSerial"] as int?,
                        SerialNumber = serialNumber++,
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
        public async Task<IActionResult> GetBoxBreakingReport(int ProjectId, int? uploadBatch = null)
        {
            try
            {
                var projectconfig = await _context.ProjectConfigs
                    .Where(p => p.ProjectId == ProjectId)
                    .FirstOrDefaultAsync();

                if (projectconfig == null)
                    return NotFound("Project config not found");

                bool resetOnSymbolChange = projectconfig.ResetOnSymbolChange;
                bool InnerBundling = projectconfig.IsInnerBundlingDone;

                var nrData = await _context.NRDatas
                    .Where(p => p.ProjectId == ProjectId)
                    .ToListAsync();

                // Get results from database
                IQueryable<BoxBreakingResult> query = _context.BoxBreakingResults
                    .Where(r => r.ProjectId == ProjectId);

                if (uploadBatch.HasValue)
                {
                    query = query.Where(r => r.UploadBatch == uploadBatch.Value);
                }
                else
                {
                    // Get latest batch
                    var maxBatch = await _context.BoxBreakingResults
                        .Where(r => r.ProjectId == ProjectId)
                        .MaxAsync(r => (int?)r.UploadBatch);
                    if (maxBatch.HasValue)
                    {
                        query = query.Where(r => r.UploadBatch == maxBatch.Value);
                    }
                }

                var results = await query.ToListAsync();

                if (!results.Any())
                    return NotFound("No box breaking results found");

                // Join with EnvelopeBreakingResults to get all details
                var envelopeBreakingResults = await _context.EnvelopeBreakingResults
                    .Where(x => x.ProjectId == ProjectId)
                    .ToListAsync();

                var resultsWithDetails = results.Select(r =>
                {
                    var envelopeResult = envelopeBreakingResults.FirstOrDefault(e => e.Id == r.EnvelopeBreakingResultId);
                    var nrRow = envelopeResult != null ? nrData.FirstOrDefault(n => n.CatchNo == envelopeResult.CatchNo) : null;
                    return new { Result = r, EnvelopeResult = envelopeResult, NrRow = nrRow };
                }).ToList();

                // Generate Excel
                var reportPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", ProjectId.ToString());
                Directory.CreateDirectory(reportPath);

                var filePath = Path.Combine(reportPath, "BoxBreakingFromDB.xlsx");
                if (System.IO.File.Exists(filePath))
                {
                    System.IO.File.Delete(filePath);
                }

                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("BoxBreaking");

                    bool hasAnyOmr = results.Any(x => !string.IsNullOrWhiteSpace(x.OmrSerial));

                    // Headers
                    var headers = new List<string>
                    {
                        "SerialNumber", "CatchNo", "CenterCode", "CenterSort", "ExamTime", "ExamDate",
                        "Quantity", "NodalCode", "NodalSort", "Route", "RouteSort", "TotalEnv",
                        "Start", "End", "Serial", "TotalPages"
                    };

                    int nextCol = 17;
                    int symbolCol = -1;
                    int courseCol = -1;

                    if (resetOnSymbolChange)
                    {
                        headers.Add("Symbol");
                        symbolCol = nextCol;
                        nextCol++;

                        headers.Add("CourseName");
                        courseCol = nextCol;
                        nextCol++;
                    }

                    headers.Add("BoxNo");
                    int boxCol = nextCol;
                    nextCol++;

                    int omrCol = -1;
                    if (hasAnyOmr)
                    {
                        headers.Add("OmrSerial");
                        omrCol = nextCol;
                        nextCol++;
                    }

                    int innerBundlingCol = -1;
                    if (InnerBundling)
                    {
                        headers.Add("InnerBundlingSerial");
                        innerBundlingCol = nextCol;
                        nextCol++;
                    }

                    for (int i = 0; i < headers.Count; i++)
                    {
                        worksheet.Cells[1, i + 1].Value = headers[i];
                        worksheet.Cells[1, i + 1].Style.Font.Bold = true;
                    }

                    int row = 2;
                    int serial = 1;

                    foreach (var item in resultsWithDetails)
                    {
                        var result = item.Result;
                        var envelopeResult = item.EnvelopeResult;
                        var nrRow = item.NrRow;

                        // Get values from EnvelopeBreakingResults first, fallback to NRData
                        var catchNo = envelopeResult?.CatchNo ?? nrRow?.CatchNo ?? "";
                        var centerCode = envelopeResult?.CenterCode ?? nrRow?.CenterCode ?? "";
                        var centerSort = envelopeResult?.CenterSort ?? nrRow?.CenterSort ?? 0;
                        var examTime = envelopeResult?.ExamTime ?? nrRow?.ExamTime ?? "";
                        var examDate = envelopeResult?.ExamDate ?? nrRow?.ExamDate ?? "";
                        var quantity = envelopeResult?.Quantity ?? nrRow?.Quantity ?? 0;
                        var nodalCode = envelopeResult?.NodalCode ?? nrRow?.NodalCode ?? "";
                        var nodalSort = envelopeResult?.NodalSort ?? nrRow?.NodalSort ?? 0;
                        var route = envelopeResult?.Route ?? nrRow?.Route ?? "";
                        var routeSort = envelopeResult?.RouteSort ?? nrRow?.RouteSort ?? 0;
                        var totalEnv = envelopeResult?.TotalEnv ?? 0;
                        var symbol = nrRow?.Symbol ?? "";
                        var courseName = envelopeResult?.CourseName ?? nrRow?.CourseName ?? "";

                        worksheet.Cells[row, 1].Value = serial++;
                        worksheet.Cells[row, 2].Value = catchNo;
                        worksheet.Cells[row, 3].Value = centerCode;
                        worksheet.Cells[row, 4].Value = centerSort;
                        worksheet.Cells[row, 5].Value = examTime;
                        worksheet.Cells[row, 6].Value = examDate;
                        worksheet.Cells[row, 7].Value = quantity;
                        worksheet.Cells[row, 8].Value = nodalCode;
                        worksheet.Cells[row, 9].Value = nodalSort;
                        worksheet.Cells[row, 10].Value = route;
                        worksheet.Cells[row, 11].Value = routeSort;
                        worksheet.Cells[row, 12].Value = totalEnv;
                        worksheet.Cells[row, 13].Value = result.Start;
                        worksheet.Cells[row, 14].Value = result.End;
                        worksheet.Cells[row, 15].Value = result.Serial;
                        worksheet.Cells[row, 16].Value = result.TotalPages;

                        int colIdx = 17;

                        if (symbolCol > 0)
                        {
                            worksheet.Cells[row, symbolCol].Value = symbol;
                        }

                        if (courseCol > 0)
                        {
                            worksheet.Cells[row, courseCol].Value = courseName;
                        }

                        worksheet.Cells[row, boxCol].Value = result.BoxNo;

                        if (omrCol > 0)
                        {
                            worksheet.Cells[row, omrCol].Value = result.OmrSerial;
                        }

                        if (innerBundlingCol > 0)
                        {
                            worksheet.Cells[row, innerBundlingCol].Value = result.InnerBundlingSerial;
                        }

                        row++;
                    }

                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                    package.SaveAs(new FileInfo(filePath));
                }

                _loggerService.LogEvent(
                    $"Generated box breaking report from DB for ProjectId {ProjectId}",
                    "BoxBreakingProcessing",
                    User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0,
                    ProjectId);

                return Ok(new
                {
                    message = "Report generated successfully",
                    filePath = filePath,
                    recordsCount = resultsWithDetails.Count
                });
            }
            catch (Exception ex)
            {
                _loggerService.LogError("Error generating box breaking report", ex.Message, nameof(BoxBreakingProcessingController));
                return StatusCode(500, new { error = ex.Message });
            }
        }
    }
}
