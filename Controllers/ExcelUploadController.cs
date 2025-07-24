using ERPToolsAPI.Data;
using ERPToolsAPI.Models;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using System.Net.Http.Headers;

namespace Tools.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelUploadController : ControllerBase
    {
        private readonly ERPToolsDbContext _context;
        private readonly string _baseUrl;

        public ExcelUploadController(ERPToolsDbContext context, IConfiguration configuration)
        {
            _context = context;
            _baseUrl = configuration["ApiSettings:BaseUrl"];
        }


        [HttpGet("group/{groupId}")]
        public IActionResult GetByGroup(int groupId)
        {
            var data = _context.ExcelUploads
                        .Where(x => x.GroupId == groupId)
                        .ToList();

            return Ok(data);
        }


        [HttpGet]
        public IActionResult GetUploads()
        {
            var uploads = _context.ExcelUploads.ToList();
            return Ok(uploads);
        }


        /*[HttpPost("upload-mapped")]
        public async Task<IActionResult> UploadMapped([FromBody] List<ExcelUpload> data)
        {
            if (data == null || !data.Any())
                return BadRequest("No data received.");

            _context.ExcelUploads.AddRange(data);
            await _context.SaveChangesAsync();

            return Ok(new { message = "Data uploaded successfully.", count = data.Count });
        }*/


        [AllowAnonymous]
        [HttpPost("upload-mapped")]
        public async Task<IActionResult> UploadMapped([FromQuery] int groupId, [FromBody] List<ExcelUpload> data)
        {
            if (data == null || !data.Any())
                return BadRequest("No data received.");

            try
            {
                using var httpClient = new HttpClient();

                // ✅ Read the JWT token from incoming request header
                var accessToken = Request.Headers["Authorization"].ToString().Replace("Bearer ", "");
                httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                if (string.IsNullOrWhiteSpace(accessToken))
                    return Unauthorized("Missing token");

                httpClient.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", accessToken);

                // ✅ Call ERP's Groups API with the token
                var groupResponse = await httpClient.GetAsync($"{_baseUrl}/Groups/{groupId}");

                if (!groupResponse.IsSuccessStatusCode)
                    return BadRequest("Invalid GroupId or unauthorized to access group.");

                var groupContent = await groupResponse.Content.ReadAsStringAsync();

                // ✅ Read group name from JSON directly
                var groupJson = JObject.Parse(groupContent);
                var groupName = groupJson["name"]?.ToString() ?? "Unknown";

                // ✅ Assign GroupId to each row
                foreach (var item in data)
                {
                    item.GroupId = groupId;
                }

                _context.ExcelUploads.AddRange(data);
                await _context.SaveChangesAsync();

                return Ok(new
                {
                    message = $"Uploaded {data.Count} records under group '{groupName}'",
                    count = data.Count
                });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { message = "Upload failed", error = ex.Message });
            }
        }


    }

}
