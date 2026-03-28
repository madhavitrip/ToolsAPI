using ERPToolsAPI.Data;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Options;
using System;
using System.Reflection;
using System.Net.Http.Headers;
using System.Text.Json;
using Tools.Models;

namespace Tools.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    [Microsoft.AspNetCore.Authorization.AllowAnonymous]
    public class RPTTemplatesController : ControllerBase
    {
        private readonly ERPToolsDbContext _context;
        private readonly IWebHostEnvironment _env;
        private readonly ApiSettings _apiSettings;
        private readonly IHttpClientFactory _httpClientFactory;

        public RPTTemplatesController(
            ERPToolsDbContext context,
            IWebHostEnvironment env,
            IOptions<ApiSettings> apiSettings,
            IHttpClientFactory httpClientFactory)
        {
            _context = context;
            _env = env;
            _apiSettings = apiSettings.Value;
            _httpClientFactory = httpClientFactory;
        }

        // GET: api/RPTTemplates
        [HttpGet]
        public async Task<ActionResult<IEnumerable<RPTTemplate>>> GetRPTTemplate()
        {
            return await _context.RPTTemplates.ToListAsync();
        }

        // GET: api/RPTTemplates/by-group?groupId=1&typeId=2
        [HttpGet("by-group")]
        public async Task<ActionResult> GetByGroup(int groupId, int typeId)
        {
            var templates = await _context.RPTTemplates
                .Where(t => t.GroupId == groupId && t.TypeId == typeId && t.IsActive)
                .OrderBy(t => t.TemplateName)
                .ToListAsync();

            return Ok(templates);
        }

        // GET: api/RPTTemplates/mapping-options?groupId=1&typeId=2
        [HttpGet("mapping-options")]
        public async Task<ActionResult> GetMappingOptions(int groupId, int typeId)
        {
            var nrColumns = GetModelColumns<NRData>(exclude: new[] { "NRDatas" });
            var envColumns = GetModelColumns<EnvelopeBreakingResult>();
            var boxColumns = GetModelColumns<BoxBreakingResult>();

            var nrJsonKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            List<int> projectIds = new();
            if (groupId > 0 && typeId > 0)
            {
                projectIds = await _context.Projects
                    .Where(p => p.GroupId == groupId && p.TypeId == typeId)
                    .Select(p => p.ProjectId)
                    .ToListAsync();
            }

            var nrQuery = _context.NRDatas.AsQueryable();
            if (projectIds.Count > 0)
            {
                nrQuery = nrQuery.Where(n => projectIds.Contains(n.ProjectId));
            }

            var nrJsonRows = await nrQuery
                .Where(n => !string.IsNullOrWhiteSpace(n.NRDatas))
                .Select(n => n.NRDatas)
                .ToListAsync();

            foreach (var json in nrJsonRows)
            {
                try
                {
                    using var doc = JsonDocument.Parse(json);
                    if (doc.RootElement.ValueKind != JsonValueKind.Object) continue;
                    foreach (var prop in doc.RootElement.EnumerateObject())
                    {
                        if (!string.IsNullOrWhiteSpace(prop.Name))
                            nrJsonKeys.Add(prop.Name);
                    }
                }
                catch
                {
                    // ignore malformed JSON rows
                }
            }

            return Ok(new
            {
                nrColumns = nrColumns.OrderBy(x => x).ToList(),
                envColumns = envColumns.OrderBy(x => x).ToList(),
                boxColumns = boxColumns.OrderBy(x => x).ToList(),
                nrJsonKeys = nrJsonKeys.OrderBy(x => x).ToList()
            });
        }

        // GET: api/RPTTemplates/5
        [HttpGet("{id}")]
        public async Task<ActionResult<RPTTemplate>> GetRPTTemplate(int id)
        {
            var t = await _context.RPTTemplates.FindAsync(id);
            if (t == null) return NotFound();
            return t;
        }

        // POST: api/RPTTemplates/upload
        // multipart/form-data: file (.rpt), groupId, typeId, templateName
        [HttpPost("upload")]
        public async Task<ActionResult> Upload([FromForm] int groupId, [FromForm] int typeId,
            [FromForm] string templateName, IFormFile file)
        {
            if (file == null || file.Length == 0)
                return BadRequest("No file uploaded.");

            var ext = Path.GetExtension(file.FileName).ToLowerInvariant();
            if (ext != ".rpt")
                return BadRequest("Only .rpt files are accepted.");

            if (string.IsNullOrWhiteSpace(templateName))
                return BadRequest("templateName is required.");

            // Determine next version for this group+type+name combo
            var lastVersion = await _context.RPTTemplates
                .Where(t => t.GroupId == groupId && t.TypeId == typeId && t.TemplateName == templateName)
                .MaxAsync(t => (int?)t.Version) ?? 0;

            // Deactivate previous versions (never delete)
            var previous = await _context.RPTTemplates
                .Where(t => t.GroupId == groupId && t.TypeId == typeId && t.TemplateName == templateName && t.IsActive)
                .ToListAsync();
            previous.ForEach(t => t.IsActive = false);

            // Save file to wwwroot/rpt-templates/{groupId}/{typeId}/
            var webRoot = _env.WebRootPath ?? Path.Combine(Directory.GetCurrentDirectory(), "wwwroot");
            var folder = Path.Combine(webRoot, "rpt-templates", groupId.ToString(), typeId.ToString());
            Directory.CreateDirectory(folder);
            var fileName     = $"{templateName}_{groupId}_{typeId}_{lastVersion + 1}{ext}";
            var absolutePath = Path.Combine(folder, fileName);
            var relativePath = Path.Combine("rpt-templates", groupId.ToString(), typeId.ToString(), fileName);

            using (var stream = new FileStream(absolutePath, FileMode.Create))
                await file.CopyToAsync(stream);

            var (parsedFields, parseError) = await ParseFieldsAsync(absolutePath, file.FileName);
            if (!string.IsNullOrWhiteSpace(parseError))
            {
                Console.WriteLine($"[RPTTemplates] Parse warning: {parseError}");
            }
            var parsedFieldsJson = parsedFields.Count > 0
                ? JsonSerializer.Serialize(parsedFields)
                : null;

            var template = new RPTTemplate
            {
                GroupId      = groupId,
                TypeId       = typeId,
                TemplateName = templateName,
                RPTFilePath  = relativePath,   // store relative path only
                ParsedFieldsJson = parsedFieldsJson,
                Version      = lastVersion + 1,
                IsActive     = true,
                CreatedDate  = DateTime.Now,
                UpdatedDate  = DateTime.Now
            };

            _context.RPTTemplates.Add(template);
            await _context.SaveChangesAsync();

            return Ok(new
            {
                template.TemplateId,
                template.TemplateName,
                template.Version,
                template.RPTFilePath,
                parsedFields = parsedFields,
                parseError = parseError
            });
        }

        // POST: api/RPTTemplates/import-from-group
        // Copy all active templates from a source group+type to a new group+type
        [HttpPost("import-from-group")]
        public async Task<ActionResult> ImportFromGroup([FromBody] ImportGroupRequest req)
        {
            var sourceTemplates = await _context.RPTTemplates
                .Where(t => t.GroupId == req.SourceGroupId && t.TypeId == req.SourceTypeId && t.IsActive)
                .ToListAsync();

            if (!sourceTemplates.Any())
                return NotFound("No active templates found for source group/type.");

            var imported = new List<object>();

            foreach (var src in sourceTemplates)
            {
                // Deactivate any existing for target group+type+name
                var existing = await _context.RPTTemplates
                    .Where(t => t.GroupId == req.TargetGroupId && t.TypeId == req.TargetTypeId
                                && t.TemplateName == src.TemplateName && t.IsActive)
                    .ToListAsync();
                existing.ForEach(t => t.IsActive = false);

                var lastVersion = await _context.RPTTemplates
                    .Where(t => t.GroupId == req.TargetGroupId && t.TypeId == req.TargetTypeId
                                && t.TemplateName == src.TemplateName)
                    .MaxAsync(t => (int?)t.Version) ?? 0;

                var newTemplate = new RPTTemplate
                {
                    GroupId      = req.TargetGroupId,
                    TypeId       = req.TargetTypeId,
                    TemplateName = src.TemplateName,
                    RPTFilePath  = src.RPTFilePath,   // reuse same file
                    ParsedFieldsJson = src.ParsedFieldsJson,
                    Version      = lastVersion + 1,
                    IsActive     = true,
                    CreatedDate  = DateTime.Now,
                    UpdatedDate  = DateTime.Now
                };
                _context.RPTTemplates.Add(newTemplate);
                await _context.SaveChangesAsync();

                // Copy mapping if requested
                if (req.CopyMappings)
                {
                    var srcMapping = await _context.RPTMappings
                        .FirstOrDefaultAsync(m => m.TemplateId == src.TemplateId);
                    if (srcMapping != null)
                    {
                        _context.RPTMappings.Add(new RPTMapping
                        {
                            TemplateId  = newTemplate.TemplateId,
                            MappingJson = srcMapping.MappingJson
                        });
                        await _context.SaveChangesAsync();
                    }
                }

                imported.Add(new { newTemplate.TemplateId, newTemplate.TemplateName, newTemplate.Version });
            }

            return Ok(new { imported });
        }

        // GET: api/RPTTemplates/5/mapping
        [HttpGet("{id}/mapping")]
        public async Task<ActionResult> GetMapping(int id)
        {
            var mapping = await _context.RPTMappings.FirstOrDefaultAsync(m => m.TemplateId == id);
            if (mapping == null) return NotFound();
            return Ok(mapping);
        }

        // POST: api/RPTTemplates/{id}/mapping
        // body: { mappingJson: "..." }
        [HttpPost("{id}/mapping")]
        public async Task<ActionResult> SaveMapping(int id, [FromBody] SaveMappingRequest req)
        {
            if (!await _context.RPTTemplates.AnyAsync(t => t.TemplateId == id))
                return NotFound("Template not found.");

            var existing = await _context.RPTMappings.FirstOrDefaultAsync(m => m.TemplateId == id);
            if (existing != null)
            {
                existing.MappingJson = req.MappingJson;
            }
            else
            {
                _context.RPTMappings.Add(new RPTMapping { TemplateId = id, MappingJson = req.MappingJson });
            }

            await _context.SaveChangesAsync();
            return Ok(new { templateId = id, message = "Mapping saved." });
        }

        // POST: api/RPTTemplates/{id}/parse-fields
        // Re-parse existing template and store parsed fields in DB
        [HttpPost("{id}/parse-fields")]
        public async Task<ActionResult> ParseFields(int id)
        {
            var t = await _context.RPTTemplates.FindAsync(id);
            if (t == null || string.IsNullOrWhiteSpace(t.RPTFilePath))
                return NotFound("Template not found.");

            var webRoot = _env.WebRootPath ?? Path.Combine(Directory.GetCurrentDirectory(), "wwwroot");
            var absolutePath = Path.Combine(webRoot, t.RPTFilePath);
            if (!System.IO.File.Exists(absolutePath))
                return NotFound("File not found on disk.");

            var (parsedFields, parseError) = await ParseFieldsAsync(absolutePath, Path.GetFileName(absolutePath));
            if (!string.IsNullOrWhiteSpace(parseError))
                return StatusCode(502, parseError);
            t.ParsedFieldsJson = parsedFields.Count > 0
                ? JsonSerializer.Serialize(parsedFields)
                : null;
            t.UpdatedDate = DateTime.Now;
            await _context.SaveChangesAsync();

            return Ok(new { templateId = id, parsedFields });
        }

        // GET: api/RPTTemplates/5/download
        [HttpGet("{id}/download")]
        public async Task<ActionResult> Download(int id)
        {
            var t = await _context.RPTTemplates.FindAsync(id);
            if (t == null || string.IsNullOrWhiteSpace(t.RPTFilePath))
                return NotFound();

            // Resolve relative path against wwwroot
            var webRoot      = _env.WebRootPath ?? Path.Combine(Directory.GetCurrentDirectory(), "wwwroot");
            var absolutePath = Path.Combine(webRoot, t.RPTFilePath);
            if (!System.IO.File.Exists(absolutePath))
                return NotFound("File not found on disk.");

            var bytes = await System.IO.File.ReadAllBytesAsync(absolutePath);
            return File(bytes, "application/octet-stream", Path.GetFileName(absolutePath));
        }

        private static List<string> GetModelColumns<T>(IEnumerable<string>? exclude = null)
        {
            var excludeSet = new HashSet<string>(exclude ?? Array.Empty<string>(), StringComparer.OrdinalIgnoreCase);
            return typeof(T)
                .GetProperties(BindingFlags.Public | BindingFlags.Instance)
                .Select(p => p.Name)
                .Where(name => !excludeSet.Contains(name))
                .ToList();
        }

        private async Task<(List<string> fields, string error)> ParseFieldsAsync(string absolutePath, string originalFileName)
        {
            if (string.IsNullOrWhiteSpace(_apiSettings?.RptParserUrl))
                return (new List<string>(), "RptParserUrl is not configured.");

            try
            {
                HttpClient client;
                if (Uri.TryCreate(_apiSettings.RptParserUrl, UriKind.Absolute, out var parserUri)
                    && parserUri.Scheme == Uri.UriSchemeHttps
                    && parserUri.IsLoopback)
                {
                    // Allow self-signed certs for local parser service (dev only)
                    var handler = new HttpClientHandler
                    {
                        ServerCertificateCustomValidationCallback =
                            HttpClientHandler.DangerousAcceptAnyServerCertificateValidator
                    };
                    client = new HttpClient(handler);
                }
                else
                {
                    client = _httpClientFactory.CreateClient();
                }

                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                using var stream = System.IO.File.OpenRead(absolutePath);
                using var content = new MultipartFormDataContent();
                var fileContent = new StreamContent(stream);
                fileContent.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
                content.Add(fileContent, "file", originalFileName);

                var response = await client.PostAsync(_apiSettings.RptParserUrl, content);
                var json = await response.Content.ReadAsStringAsync();
                if (!response.IsSuccessStatusCode)
                    return (new List<string>(), $"Parser error {(int)response.StatusCode}: {json}");

                using var doc = JsonDocument.Parse(json);
                if (!doc.RootElement.TryGetProperty("data", out var dataElement))
                    return (new List<string>(), "Parser response missing data.");

                if (dataElement.ValueKind == JsonValueKind.Array)
                {
                    var fields = new List<string>();
                    foreach (var item in dataElement.EnumerateArray())
                    {
                        if (item.ValueKind == JsonValueKind.String)
                            fields.Add(item.GetString());
                    }
                    return (fields, "");
                }

                return (new List<string>(), "Parser data was not an array.");
            }
            catch (Exception ex)
            {
                return (new List<string>(), $"Parser exception: {ex.Message}");
            }
        }
    }

    public class ImportGroupRequest
    {
        public int SourceGroupId  { get; set; }
        public int SourceTypeId   { get; set; }
        public int TargetGroupId  { get; set; }
        public int TargetTypeId   { get; set; }
        public bool CopyMappings  { get; set; } = true;
    }

    public class SaveMappingRequest
    {
        public string MappingJson { get; set; }
    }
}
