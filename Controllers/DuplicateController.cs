using ERPToolsAPI.Data;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Reflection;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using Tools.Models;
using System.Text.Json;
using Humanizer;
using Newtonsoft.Json;
using Tools.Migrations;
using NRData = Tools.Models.NRData;
using Tools.Services;

namespace Tools.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class DuplicateController : ControllerBase
    {
        private readonly ERPToolsDbContext _context;
        private readonly ILoggerService _logger;

        public DuplicateController(ERPToolsDbContext context, ILoggerService loggerService)
        {
            _context = context;
            _logger = loggerService;
        }

        [HttpPost]
        public async Task<IActionResult> MergeFields(int ProjectId)
        {
            try
            {
                // Get project data
                var data = await _context.NRDatas
                    .Where(p => p.ProjectId == ProjectId)
                    .ToListAsync();

                var projectconfig = await _context.ProjectConfigs
                    .Where(p => p.ProjectId == ProjectId).FirstOrDefaultAsync();

                if (projectconfig == null)
                {
                    return NotFound("Project config not exists for this project");
                }
                var mergeFieldIds = projectconfig.DuplicateCriteria.ToList();

                var fieldNames = await _context.Fields
                .Where(f => mergeFieldIds.Contains(f.FieldId)) // Match with the field IDs
                .Select(f => f.Name) // Get the field names
                .ToListAsync();

                if (!data.Any())
                    return NotFound("Nr data not found for this project.");

                // Group the data based on the merge fields
                var grouped = data.GroupBy(d =>
                {
                    var key = new List<string>();
                    foreach (var field in fieldNames)
                    {
                        var value = d.GetType().GetProperty(field, BindingFlags.IgnoreCase | BindingFlags.Public | BindingFlags.Instance)
                                     ?.GetValue(d)?.ToString()?.Trim() ?? "";
                        key.Add(value);
                    }
                    return string.Join("|", key); // Composite key
                });

                int mergedCount = 0;
                var deletedRows = new List<NRData>();
                var reportRows = new List<NRData>();

                foreach (var group in grouped)
                {
                    reportRows.AddRange(group); // Include all rows for reporting

                    if (group.Count() <= 1)
                        continue;

                    var keep = group.First();
                    keep.Quantity = group.Sum(x => x.NRQuantity);
                    var subjectValues = group.Select(x => x.SubjectName?.Trim())
                                 .Where(v => !string.IsNullOrEmpty(v))
                                 .Distinct(StringComparer.OrdinalIgnoreCase)
                                 .ToList();

                    if (subjectValues.Count > 1)
                        keep.SubjectName = string.Join(" / ", subjectValues);
                    else if (subjectValues.Count == 1)
                        keep.SubjectName = subjectValues.First();

                    var courseValues = group.Select(x => x.CourseName?.Trim())
                                            .Where(v => !string.IsNullOrEmpty(v))
                                            .Distinct(StringComparer.OrdinalIgnoreCase)
                                            .ToList();

                    if (courseValues.Count > 1)
                        keep.CourseName = string.Join(" / ", courseValues);
                    else if (courseValues.Count == 1)
                        keep.CourseName = courseValues.First();



                    var duplicates = group.Skip(1).ToList();
                    _context.NRDatas.RemoveRange(duplicates);

                    mergedCount += duplicates.Count;
                    deletedRows.AddRange(duplicates);
                }

                int smallestInner = 0;
                var innerEnv = await _context.ProjectConfigs.
                    Where(s => s.ProjectId == ProjectId).Select(s => s.Envelope)
                    .FirstOrDefaultAsync();
                if (!string.IsNullOrEmpty(innerEnv))
                {
                    Console.WriteLine(innerEnv);
                    try
                    {
                        var envelopeDict = System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, string>>(innerEnv);
                        if (envelopeDict != null && envelopeDict.TryGetValue("Inner", out var innerValue) &&
                         !string.IsNullOrWhiteSpace(innerValue))
                        {
                            var innerSizes = envelopeDict["Inner"]
                                .Split(',', StringSplitOptions.RemoveEmptyEntries)
                                .Select(e => e.Trim().ToUpper().Replace("E", "")) // get number from E10 → 10
                                .Where(x => int.TryParse(x, out _))
                                .Select(int.Parse)
                                .OrderBy(x => x)
                                .ToList();

                            if (innerSizes.Any())
                            {
                                smallestInner = innerSizes.First();
                            }
                        }
                        else
                        {
                            Console.WriteLine("Entering in this loop");

                            var innerSizes = envelopeDict["Outer"]
                                .Split(',', StringSplitOptions.RemoveEmptyEntries)
                                .Select(e => e.Trim().ToUpper().Replace("E", "")) // get number from E10 → 10
                                .Where(x => int.TryParse(x, out _))
                                .Select(int.Parse)
                                .OrderBy(x => x)
                                .ToList();

                            if (innerSizes.Any())
                            {
                                smallestInner = innerSizes.First();
                                Console.WriteLine(smallestInner);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError("Error in serializing Envelope", ex.Message, nameof(DuplicateController));
                        return StatusCode(500, "Internal server error");
                    }
                }
               
                if (smallestInner > 0)
                {
                    Console.WriteLine(smallestInner);
                    if (projectconfig.Enhancement > 0)
                    {
                        foreach (var d in data)
                        {
                            if (d.NRQuantity > 0)
                            {
                                d.Quantity = d.NRQuantity + (int)Math.Round((projectconfig.Enhancement * d.NRQuantity) / 100.0);

                                // Round up to nearest multiple of smallestInner
                                d.Quantity = (int)Math.Ceiling(d.Quantity / (double)smallestInner) * smallestInner;
                            }
                        }
                    }
                    else
                    {
                        foreach (var d in data)
                        {
                            if (d.NRQuantity > 0)
                            {
                                Console.WriteLine(d.NRQuantity);
                                d.Quantity = (int)Math.Ceiling(d.NRQuantity / (double)smallestInner) * smallestInner;
                            }
                        }
                    }
                }
                _logger.LogEvent($"Duplicates has been deleted", "Duplicates", User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0, ProjectId);
                await _context.SaveChangesAsync();
                // Excel Report Path
                var reportPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", ProjectId.ToString());
                if (!Directory.Exists(reportPath))
                {
                    // Create the directory if it doesn't exist
                    Directory.CreateDirectory(reportPath);
                }

                var filename = "DuplicateTool.xlsx";
                var filePath = Path.Combine(reportPath, filename);
                if (System.IO.File.Exists(filePath))
                {
                    System.IO.File.Delete(filePath);
                }
                // Gather static properties (excluding NRDatas)
                var baseProperties = typeof(NRData).GetProperties(BindingFlags.Public | BindingFlags.Instance)
                                                   .Where(p => p.Name != "NRDatas" && p.Name != "Id" && p.Name != "ProjectId")
                                                   .Where(p => reportRows.Any(row => p.GetValue(row) != null && !string.IsNullOrEmpty(p.GetValue(row)?.ToString())))
                                                   .ToList();

                var extraHeaders = new HashSet<string>();
                var parsedRows = new List<(NRData row, Dictionary<string, string> extras)>();
                try
                {
                    foreach (var row in reportRows)
                    {
                        var extras = new Dictionary<string, string>();

                        if (!string.IsNullOrEmpty(row.NRDatas))
                        {
                            try
                            {
                                extras = System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, string>>(row.NRDatas);
                                if (extras != null)
                                {
                                    foreach (var key in extras.Keys)
                                        extraHeaders.Add(key);
                                }
                            }
                            catch(Exception ex) 
                            {
                                _logger.LogError("Error in deserializing NRData", ex.Message, nameof(DuplicateController));
                                return StatusCode(500, "Internal server error");
                            }
                        }

                        parsedRows.Add((row, extras));
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogError("Error in reportrows", ex.Message, nameof(DuplicateController));
                    return StatusCode(500, "Internal server error");

                }
                using (var package = new ExcelPackage())
                {
                    var ws = package.Workbook.Worksheets.Add("Merge Report");

                    // Write Headers
                    int col = 1;
                    foreach (var prop in baseProperties)
                    {
                        ws.Cells[1, col].Value = prop.Name;
                        ws.Cells[1, col].Style.Font.Bold = true;
                        col++;
                    }

                    var extraHeaderList = extraHeaders.OrderBy(k => k).ToList();
                    foreach (var key in extraHeaderList)
                    {
                        ws.Cells[1, col].Value = key;
                        ws.Cells[1, col].Style.Font.Bold = true;
                        col++;
                    }

                    // Write Rows
                    int rowIdx = 2;
                    foreach (var (item, extras) in parsedRows)
                    {
                        col = 1;

                        foreach (var prop in baseProperties)
                        {
                            var value = prop.GetValue(item);
                            ws.Cells[rowIdx, col++].Value = value?.ToString() ?? "";
                        }

                        foreach (var key in extraHeaderList)
                        {
                            extras.TryGetValue(key, out var val);
                            ws.Cells[rowIdx, col++].Value = val ?? "";
                        }

                        // Highlight deleted rows
                        if (deletedRows.Any(x => x.Id == item.Id))
                        {
                            using (var range = ws.Cells[rowIdx, 1, rowIdx, col - 1])
                            {
                                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                range.Style.Fill.BackgroundColor.SetColor(Color.Red);
                            }
                        }

                        rowIdx++;
                    }

                    ws.Cells[ws.Dimension.Address].AutoFitColumns();
                    package.SaveAs(new FileInfo(filePath));
                    _logger.LogEvent($"Duplicates report has been created", "Duplicates", User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0, ProjectId);
                }

                return Ok(new
                {
                MergedRows = mergedCount
                });
            }
            catch (Exception ex)
            {
                _logger.LogError("Error solving duplicates", ex.Message, nameof(DuplicateController));
                return StatusCode(500, "Internal server error");
            }
        }



        [HttpGet("MergeData")]
        public async Task<IActionResult> MergeEnvelope(int ProjectId)
        {
            // Helper function to parse JSON and sum the values (same as your existing one)
            int ParseJsonEnvelope(string value)
            {
                if (string.IsNullOrEmpty(value))
                {
                    return 0; // Return 0 if the value is null or empty
                }

                // Try parsing the string to an integer
                if (int.TryParse(value, out var result))
                {
                    return result;
                }
                else
                {
                    // Log invalid value if it cannot be parsed
                    Console.WriteLine($"Invalid value for envelope: {value}");
                    return 0;
                }
            }

            // Helper function to parse JSON and sum values for EnvelopeBreakages
            int ParseJsonEnvBreakage(string value)
            {
                if (string.IsNullOrEmpty(value))
                {
                    return 0; // Return 0 if the value is null or empty
                }

                // Sum values from JSON string like {\"E100\":\"1\",\"E50\":\"1\",\"E10\":\"3\"}
                try
                {
                    var envelopeData = Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, int>>(value);
                    return envelopeData?.Values.Sum() ?? 0;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error parsing envelope data: {ex.Message}");
                    return 0;
                }
            }

            // Step 1: Group NRData by CatchNo and sum the quantity
            var NRdataGrouped = await _context.NRDatas
                .Where(p => p.ProjectId == ProjectId)
                .GroupBy(p => p.CatchNo)
                .Select(g => new
                {
                    CatchNo = g.Key,
                    TotalQuantity = g.Sum(p => p.Quantity),
                    NRDataId = g.Select(p => p.Id).ToList()
                })
                .ToListAsync();

            // Step 2: Fetch ExtrasEnvelope data and then group/sum in-memory
            var extraEnvelopeData = await _context.ExtrasEnvelope
                .Where(p => p.ProjectId == ProjectId)
                .ToListAsync();

            var ExtraEnvGrouped = extraEnvelopeData
                .GroupBy(p => p.CatchNo)
                .Select(g =>
                {
                    // Sum the quantities and the envelope values
                    var totalQuantity = g.Sum(p => p.Quantity);
                    var innerEnvelopeSum = g.Sum(p => ParseJsonEnvelope(p.InnerEnvelope));
                    var outerEnvelopeSum = g.Sum(p => ParseJsonEnvelope(p.OuterEnvelope));

                    // Log the summed values
                    Console.WriteLine($"CatchNo: {g.Key}, TotalQuantity: {totalQuantity}, InnerEnvelopeSum: {innerEnvelopeSum}, OuterEnvelopeSum: {outerEnvelopeSum}");

                    return new
                    {
                        CatchNo = g.Key,
                        TotalQuantity = totalQuantity,
                        InnerEnvelopeSum = innerEnvelopeSum,
                        OuterEnvelopeSum = outerEnvelopeSum
                    };
                })
                .ToList();

            // Step 3: Fetch EnvelopeBreakages data and group by NRDataIds, then sum in-memory
            var innerEnvData = await _context.EnvelopeBreakages
                .Where(p => p.ProjectId == ProjectId)
                .ToListAsync();

            var processedInnerEnv = innerEnvData
                .GroupBy(p => p.NrDataId) // Grouping by NrDataId
                .Select(g => new
                {
                    NRDataId = g.Key,
                    TotalInnerEnvelopes = g.Sum(p => ParseJsonEnvBreakage(p.InnerEnvelope)),
                    TotalOuterEnvelopes = g.Sum(p => ParseJsonEnvBreakage(p.OuterEnvelope))
                })
                .ToList();

            // Step 4: Merge NRDataGrouped and processedInnerEnv in-memory
            var mergedResult = from nr in NRdataGrouped
                               join ee in ExtraEnvGrouped on nr.CatchNo equals ee.CatchNo into extraEnvGroup
                               from extraEnv in extraEnvGroup.DefaultIfEmpty()
                               select new
                               {
                                   CatchNo = nr.CatchNo,
                                   NRDataIds = nr.NRDataId,
                                   TotalNRDataQuantity = nr.TotalQuantity + (extraEnv?.TotalQuantity ?? 0),
                                   TotalInnerEnvEnvelopes = processedInnerEnv
                               .Where(env => nr.NRDataId.Contains(env.NRDataId))
                               .Sum(env => env.TotalInnerEnvelopes)
                               + (extraEnv?.InnerEnvelopeSum ?? 0),  // Add ExtraEnv inner envelope

                                   TotalOuterEnvEnvelopes = processedInnerEnv
                               .Where(env => nr.NRDataId.Contains(env.NRDataId))
                               .Sum(env => env.TotalOuterEnvelopes)
                               + (extraEnv?.OuterEnvelopeSum ?? 0)
                               };

            return Ok(mergedResult);
        }




    }
}