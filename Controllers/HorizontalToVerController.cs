using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System.Data;
using ClosedXML.Excel;
using ERPToolsAPI.Data;
using Tools.Services;

namespace Tools.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class HorizontalToVerController : ControllerBase
    {
        private readonly ERPToolsDbContext _context;
        private readonly ILoggerService _loggerService;
        public HorizontalToVerController(ERPToolsDbContext context, ILoggerService loggerService)
        {
            _context = context;
            _loggerService = loggerService;
        }
        /*  [HttpGet("process")]
          public IActionResult ProcessExcel()
          {
              try
              {
                  // Define root folder (application base directory)
                  string rootPath = AppContext.BaseDirectory;

                  // Define input and output file paths
                  string inputFile = Path.Combine(rootPath, "input.xlsx");
                  string outputFile = Path.Combine(rootPath, "output_hr_to_vr_all_sheets.xlsx");

                  if (!System.IO.File.Exists(inputFile))
                  {
                      return BadRequest($"❌ File not found in root folder: {inputFile}");
                  }

                  using var inputWorkbook = new XLWorkbook(inputFile);
                  using var outputWorkbook = new XLWorkbook();

                  foreach (var worksheet in inputWorkbook.Worksheets)
                  {
                      Console.WriteLine($"Processing sheet: {worksheet.Name}");

                      // Load worksheet into DataTable
                      var dt = worksheet.RangeUsed().AsTable().AsNativeDataTable();

                      // Define main columns
                      var mainCols = new[] { "EXAM_CENTER", "DISTRICT", "COLLEG" };

                      // Find numeric columns (like "101", "102", etc.)
                      var centerCols = dt.Columns.Cast<DataColumn>()
                          .Where(c => int.TryParse(c.ColumnName, out _))
                          .Select(c => c.ColumnName)
                          .ToList();

                      // Create output table
                      var outputTable = new DataTable();
                      outputTable.Columns.AddRange(new[]
                      {
                          new DataColumn("EXAM_CENTER"),
                          new DataColumn("DISTRICT"),
                          new DataColumn("COLLEG"),
                          new DataColumn("CENTER"),
                          new DataColumn("COUNT")
                      });

                      // Flatten data like Python script
                      foreach (DataRow row in dt.Rows)
                      {
                          foreach (var center in centerCols)
                          {
                              string countValue = row[center]?.ToString() ?? "";

                              outputTable.Rows.Add(
                                  row[mainCols[0]]?.ToString() ?? "",
                                  row[mainCols[1]]?.ToString() ?? "",
                                  row[mainCols[2]]?.ToString() ?? "",
                                  center,
                                  countValue
                              );
                          }
                      }

                      // Create a safe sheet name
                      string safeSheetName = new string(worksheet.Name.Select(c =>
                          char.IsLetterOrDigit(c) || c == ' ' || c == '_' || c == '-' ? c : '_'
                      ).ToArray());

                      // Add processed table to output workbook
                      var newSheet = outputWorkbook.Worksheets.Add(safeSheetName);
                      newSheet.Cell(1, 1).InsertTable(outputTable);

                      Console.WriteLine($"✅ Sheet '{worksheet.Name}' processed ({outputTable.Rows.Count} rows)");
                  }

                  // Save final output file
                  outputWorkbook.SaveAs(outputFile);

                  Console.WriteLine($"✅ All sheets processed successfully! Output: {outputFile}");

                  return Ok(new
                  {
                      message = "✅ Processing complete",
                      outputFile
                  });
              }
              catch (Exception ex)
              {
                  return StatusCode(500, $"❌ Error: {ex.Message}");
              }
          }

          [HttpPost("excel")]
          public async Task<IActionResult> UploadExcel([FromForm] IFormFile file, [FromForm] string fixedHeaders)
          {
              if (file == null || file.Length == 0)
                  return BadRequest("No file uploaded.");

              var uploadsFolder = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "uploads");
              if (!Directory.Exists(uploadsFolder))
                  Directory.CreateDirectory(uploadsFolder);

              var filePath = Path.Combine(uploadsFolder, file.FileName);

              using (var stream = new FileStream(filePath, FileMode.Create))
              {
                  await file.CopyToAsync(stream);
              }

              // Deserialize the fixed headers
              var headers = System.Text.Json.JsonSerializer.Deserialize<List<string>>(fixedHeaders);

              return Ok(new
              {
                  message = "File uploaded successfully.",
                  receivedHeaders = headers
              });
          }
  */


        [HttpPost("upload-and-process")]
        public async Task<IActionResult> UploadAndProcessExcel([FromForm] IFormFile file, [FromForm] string fixedHeaders, [FromForm] int ProjectId)
        {
            if (file == null || file.Length == 0)
                return BadRequest("No file uploaded.");

            if (string.IsNullOrWhiteSpace(fixedHeaders))
                return BadRequest("Fixed headers not provided.");

            if(ProjectId == 0)
            {
                return BadRequest("Project is invalid");
            }

        
            try
            {
                // Step 1: Deserialize fixed headers (from JSON string)
                var mainCols = System.Text.Json.JsonSerializer.Deserialize<List<string>>(fixedHeaders);
                if (mainCols == null || mainCols.Count == 0)
                    return BadRequest("Invalid fixed headers provided.");

                // Step 2: Save uploaded file
                string rootPath = Directory.GetCurrentDirectory();
                string uploadsFolder = Path.Combine(rootPath, "wwwroot", "uploads", ProjectId.ToString());

                if (!Directory.Exists(uploadsFolder))
                {
                    Directory.CreateDirectory(uploadsFolder);
                }

                string inputFilePath = Path.Combine(uploadsFolder, file.FileName);
                if (System.IO.File.Exists(inputFilePath))
                {
                    System.IO.File.Delete(inputFilePath);
                }
                using (var stream = new FileStream(inputFilePath, FileMode.Create))
                {
                    await file.CopyToAsync(stream);
                }

                // Step 3: Define output file
                string outputFile = Path.Combine(uploadsFolder, "Output.xlsx");

                // Step 4: Process Excel using fixed headers
                using var inputWorkbook = new XLWorkbook(inputFilePath);
                using var outputWorkbook = new XLWorkbook();

                foreach (var worksheet in inputWorkbook.Worksheets)
                {
                    Console.WriteLine($"Processing sheet: {worksheet.Name}");

                    // Load sheet into DataTable
                    var dt = worksheet.RangeUsed().AsTable().AsNativeDataTable();

                    // Find numeric columns (like "101", "102", etc.)
                    var centerCols = dt.Columns.Cast<DataColumn>()
                        .Where(c => int.TryParse(c.ColumnName, out _))
                        .Select(c => c.ColumnName)
                        .ToList();

                    // Create output table
                    var outputTable = new DataTable();
                    foreach (var col in mainCols)
                        outputTable.Columns.Add(new DataColumn(col));

                    outputTable.Columns.Add(new DataColumn("CENTER"));
                    outputTable.Columns.Add(new DataColumn("COUNT"));

                    // Flatten data
                    foreach (DataRow row in dt.Rows)
                    {
                        foreach (var center in centerCols)
                        {
                            string countValue = row[center]?.ToString() ?? "";
                            var rowValues = mainCols.Select(c => row[c]?.ToString() ?? "").ToList();
                            rowValues.Add(center);
                            rowValues.Add(countValue);
                            outputTable.Rows.Add(rowValues.ToArray());
                        }
                    }

                    // Create a safe sheet name
                    string safeSheetName = new string(worksheet.Name.Select(c =>
                        char.IsLetterOrDigit(c) || c == ' ' || c == '_' || c == '-' ? c : '_'
                    ).ToArray());

                    var newSheet = outputWorkbook.Worksheets.Add(safeSheetName);
                    newSheet.Cell(1, 1).InsertTable(outputTable);

                    Console.WriteLine($"✅ Sheet '{worksheet.Name}' processed ({outputTable.Rows.Count} rows)");
                }

                // Step 5: Save processed file
                outputWorkbook.SaveAs(outputFile);

                Console.WriteLine($"✅ File processed successfully! Output: {outputFile}");

                return Ok( "✅ File uploaded and processed successfully!"
                );
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"❌ Error processing Excel file: {ex.Message}");
            }
        }


        [HttpGet("check-file/{projectId}")]
        public IActionResult CheckFile(int projectId)
        {
            string uploadsFolder = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "uploads", projectId.ToString());
            string outputFile = Path.Combine(uploadsFolder, "Output.xlsx");

            if (System.IO.File.Exists(outputFile))
                return Ok(new { exists = true, fileUrl = $"/uploads/{projectId}/Output.xlsx" });

            return Ok(new { exists = false });
        }

    }
}
