using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Options;
using OfficeOpenXml;
using System.Reflection;
using Tools.Data;
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
                // Read from EnvelopeBreakingResults and build data rows
                // ==============================
                var breakingReportData = new List<dynamic>();

                foreach (var result in envelopeResults)
                {
                    dynamic row = new System.Dynamic.ExpandoObject();
                    var rowDict = (IDictionary<string, object>)row;

                    // Get data from EnvelopeBreakingResults first
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
                    rowDict["CourseName"] = result.CourseName ?? "";
                    rowDict["EnvelopeBreakingResultId"] = result.Id;
                    rowDict["NrDataId"] = result.NrDataId;

                    // Fallback to NRData if needed
                    var nrRow = nrData.FirstOrDefault(n => n.Id == result.NrDataId);
                    if (nrRow != null)
                    {
                        rowDict["Symbol"] = nrRow.Symbol ?? "";
                        rowDict["Pages"] = nrRow.Pages;
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
                    int pages = (int)itemDict["Pages"];
                    int quantity = (int)itemDict["Quantity"];
                    int totalPages = quantity * pages;
                    string currentSymbol = resetOnSymbolChange ? (itemDict["Symbol"]?.ToString() ?? "") : "";
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
        public async Task<IActionResult> GetBoxBreakingReport(int ProjectId)
        {
            try
            {
                var boxResults = await _context.BoxBreakingResults
                    .Where(r => r.ProjectId == ProjectId)
                    .OrderByDescending(r => r.UploadBatch)
                    .ThenBy(r => r.SerialNumber)
                    .ToListAsync();

                if (!boxResults.Any())
                    return NotFound("No box breaking results found");

                var envelopeResults = await _context.EnvelopeBreakingResults
                    .ToListAsync();

                var nrData = await _context.NRDatas
                    .Where(p => p.ProjectId == ProjectId)
                    .ToListAsync();

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

                    // Headers
                    worksheet.Cells[1, 1].Value = "SerialNumber";
                    worksheet.Cells[1, 2].Value = "CatchNo";
                    worksheet.Cells[1, 3].Value = "CenterCode";
                    worksheet.Cells[1, 4].Value = "CenterSort";
                    worksheet.Cells[1, 5].Value = "ExamTime";
                    worksheet.Cells[1, 6].Value = "ExamDate";
                    worksheet.Cells[1, 7].Value = "Quantity";
                    worksheet.Cells[1, 8].Value = "NodalCode";
                    worksheet.Cells[1, 9].Value = "NodalSort";
                    worksheet.Cells[1, 10].Value = "Route";
                    worksheet.Cells[1, 11].Value = "RouteSort";
                    worksheet.Cells[1, 12].Value = "TotalEnv";
                    worksheet.Cells[1, 13].Value = "Start";
                    worksheet.Cells[1, 14].Value = "End";
                    worksheet.Cells[1, 15].Value = "Serial";
                    worksheet.Cells[1, 16].Value = "Pages";
                    worksheet.Cells[1, 17].Value = "TotalPages";
                    worksheet.Cells[1, 18].Value = "BoxNo";
                    worksheet.Cells[1, 19].Value = "OmrSerial";
                    worksheet.Cells[1, 20].Value = "InnerBundlingSerial";

                    int row = 2;
                    foreach (var result in boxResults)
                    {
                        // Get EnvelopeBreakingResult data
                        var envResult = envelopeResults.FirstOrDefault(e => e.Id == result.EnvelopeBreakingResultId);
                        if (envResult == null)
                            continue;

                        // Get NRData for additional fields
                        var nrRow = nrData.FirstOrDefault(n => n.Id == envResult.NrDataId);

                        worksheet.Cells[row, 1].Value = result.SerialNumber;
                        worksheet.Cells[row, 2].Value = envResult.CatchNo;
                        worksheet.Cells[row, 3].Value = envResult.CenterCode;
                        worksheet.Cells[row, 4].Value = envResult.CenterSort;
                        worksheet.Cells[row, 5].Value = envResult.ExamTime;
                        worksheet.Cells[row, 6].Value = envResult.ExamDate;
                        worksheet.Cells[row, 7].Value = envResult.Quantity;
                        worksheet.Cells[row, 8].Value = envResult.NodalCode;
                        worksheet.Cells[row, 9].Value = envResult.NodalSort;
                        worksheet.Cells[row, 10].Value = envResult.Route;
                        worksheet.Cells[row, 11].Value = envResult.RouteSort;
                        worksheet.Cells[row, 12].Value = envResult.TotalEnv;
                        worksheet.Cells[row, 13].Value = result.Start;
                        worksheet.Cells[row, 14].Value = result.End;
                        worksheet.Cells[row, 15].Value = result.Serial;
                        worksheet.Cells[row, 16].Value = nrRow?.Pages ?? 0;
                        worksheet.Cells[row, 17].Value = result.TotalPages;
                        worksheet.Cells[row, 18].Value = result.BoxNo;
                        worksheet.Cells[row, 19].Value = result.OmrSerial;
                        worksheet.Cells[row, 20].Value = result.InnerBundlingSerial;

                        row++;
                    }

                    FileInfo fi = new FileInfo(filePath);
                    package.SaveAs(fi);
                }

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
