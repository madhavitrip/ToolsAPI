
using DocumentFormat.OpenXml.Bibliography;
using ERPToolsAPI.Data;
using Microsoft.AspNetCore.Mvc;
using Microsoft.CodeAnalysis;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Options;
using MySqlConnector;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.IdentityModel.Tokens.Jwt;
using System.Net.Http.Headers;
using System.Reflection;
using System.Security.Claims;
using System.Security.Cryptography;
using System.Text.Json;
using Tools.Models;
using Tools.Services;

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
        private readonly ILoggerService _loggerService;

        public RPTTemplatesController(
            ERPToolsDbContext context,
            IWebHostEnvironment env,
            IOptions<ApiSettings> apiSettings,
            IHttpClientFactory httpClientFactory,
            ILoggerService loggerService)
        {
            _context = context;
            _env = env;
            _apiSettings = apiSettings.Value;
            _httpClientFactory = httpClientFactory;
            _loggerService = loggerService;
        }

        // GET: api/RPTTemplates
        [HttpGet]
        public async Task<ActionResult<IEnumerable<RPTTemplate>>> GetRPTTemplate()
        {
            return await _context.RPTTemplates.ToListAsync();
        }

        // GET: api/RPTTemplates/by-group?typeId=2&groupId=1&projectId=10
        [HttpGet("by-group")]
        public async Task<ActionResult> GetByGroup(
     [FromQuery] int? typeId,
     [FromQuery] int? groupId,
     [FromQuery] int? projectId)
        {
            groupId = NormalizeNullableId(groupId);
            projectId = NormalizeNullableId(projectId);

            if (projectId.HasValue)
            {
                var project = await _context.Projects
                    .FirstOrDefaultAsync(p => p.ProjectId == projectId.Value);

                if (project == null)
                    return NotFound("Project not found.");

                // Only fill missing values
                if (!groupId.HasValue)
                    groupId = project.GroupId;

                if (!typeId.HasValue)
                    typeId = project.TypeId;

                // Still validate after attempting fill
                if (!groupId.HasValue || !typeId.HasValue)
                    return BadRequest("groupId and typeId could not be resolved.");

                var resolved = await ResolveTemplatesForContext(
                    typeId.Value,
                    groupId.Value,
                    projectId.Value);

                return Ok(resolved);
            }

            if (groupId.HasValue)
            {
                if (!typeId.HasValue)
                    return BadRequest("typeId is required when groupId is provided.");

                var groupTemplates = await _context.RPTTemplates
                    .Where(t => t.GroupId == groupId
                                && t.TypeId == typeId
                                && t.ProjectId == null
                                && t.IsActive)
                    .OrderBy(t => t.TemplateName)
                    .ToListAsync();

                return Ok(groupTemplates);
            }

            if (!typeId.HasValue)
                return BadRequest("typeId is required.");

            var standardTemplates = await _context.RPTTemplates
                .Where(t => t.GroupId == null
                            && t.ProjectId == null
                            && t.TypeId == typeId
                            && t.IsActive)
                .OrderBy(t => t.TemplateName)
                .ToListAsync();

            return Ok(standardTemplates);
        }

        // GET: api/RPTTemplates/versions?typeId=2&templateName=ABC&groupId=1&projectId=10
        [HttpGet("versions")]
        public async Task<ActionResult> GetVersions(
            [FromQuery] int typeId,
            [FromQuery] string templateName,
            [FromQuery] int? groupId,
            [FromQuery] int? projectId)
        {
            if (typeId <= 0)
                return BadRequest("typeId is required.");
            if (string.IsNullOrWhiteSpace(templateName))
                return BadRequest("templateName is required.");

            groupId = NormalizeNullableId(groupId);
            projectId = NormalizeNullableId(projectId);

            if (projectId.HasValue && !groupId.HasValue)
                return BadRequest("groupId is required when projectId is provided.");

            var query = _context.RPTTemplates
                .Where(t => t.TypeId == typeId && t.TemplateName == templateName);

            if (projectId.HasValue)
            {
                query = query.Where(t => t.GroupId == groupId && t.ProjectId == projectId);
            }
            else if (groupId.HasValue)
            {
                query = query.Where(t => t.GroupId == groupId && t.ProjectId == null);
            }
            else
            {
                query = query.Where(t => t.GroupId == null && t.ProjectId == null);
            }

            var versions = await query
                .OrderByDescending(t => t.Version)
                .ToListAsync();

            return Ok(versions);
        }

        // GET: api/RPTTemplates/mapping-options?groupId=1&typeId=2
        [HttpGet("mapping-options")]
        public async Task<ActionResult> GetMappingOptions(int groupId, int typeId, int? projectId = null)
        {
            var excludeColumns = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "id", "projectid", "envelopetype", "envelopebreakingresultid",
                "createdat", "uploadedbatch", "uploadbatch", "nrdataid",
                "extraid", "envelopid", "envelopeid", "status", "lotno"
            };

            List<string> FilterColumns(List<string> columns) =>
                columns.Where(c => !excludeColumns.Contains(c)).ToList();

            // Direct NRData model columns (excluding the JSON blob column itself)
            var nrDirectColumns = new HashSet<string>(
                GetModelColumns<NRData>(exclude: new[] { "NRDatas", "UploadList" }),
                StringComparer.OrdinalIgnoreCase);

            var nrColumns = FilterColumns(nrDirectColumns.ToList());
            var envColumns = FilterColumns(GetModelColumns<EnvelopeBreakingResult>());
            var envBreakageCols = FilterColumns(GetModelColumns<EnvelopeBreakage>());
            var boxColumns = FilterColumns(GetModelColumns<BoxBreakingResult>());
            var extraConfigCols = FilterColumns(GetModelColumns<ExtrasConfiguration>());

            var nrJsonKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var envBreakageJsonKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            List<int> projectIds = new();
            if (groupId > 0 && typeId > 0)
            {
                projectIds = await _context.Projects
                    .Where(p => p.GroupId == groupId && p.TypeId == typeId)
                    .Select(p => p.ProjectId)
                    .ToListAsync();
            }

            // NR JSON keys
            var nrQuery = _context.NRDatas.AsQueryable();
            if (projectIds.Count > 0) nrQuery = nrQuery.Where(n => projectIds.Contains(n.ProjectId));
            foreach (var json in await nrQuery.Where(n => !string.IsNullOrWhiteSpace(n.NRDatas)).Select(n => n.NRDatas).ToListAsync())
            {
                try
                {
                    using var doc = JsonDocument.Parse(json);
                    foreach (var p in doc.RootElement.EnumerateObject())
                        if (!string.IsNullOrWhiteSpace(p.Name)
                            && !p.Name.StartsWith("json:", StringComparison.OrdinalIgnoreCase))
                            nrJsonKeys.Add(p.Name);
                }
                catch { }
            }

            // EnvBreakage JSON keys
            var ebQuery = _context.EnvelopeBreakages.AsQueryable();
            if (projectIds.Count > 0) ebQuery = ebQuery.Where(e => projectIds.Contains(e.ProjectId));
            foreach (var json in await ebQuery.Where(e => !string.IsNullOrWhiteSpace(e.InnerEnvelope)).Select(e => e.InnerEnvelope).ToListAsync())
            {
                try { using var doc = JsonDocument.Parse(json); foreach (var p in doc.RootElement.EnumerateObject()) if (!string.IsNullOrWhiteSpace(p.Name)) envBreakageJsonKeys.Add(p.Name); }
                catch { }
            }

            // ExtraConfig � only expose Inner from EnvelopeType JSON, label = actual value (e.g. "E10")
            // Use projectId directly if provided, otherwise fall back to projectIds from group+type
            string innerEnvelopeValue = null;
            var extraProjectIds = projectId.HasValue && projectId.Value > 0
                ? new List<int> { projectId.Value }
                : projectIds;

            if (extraProjectIds.Count > 0)
            {
                var envelopeTypeJson = await _context.ExtraConfigurations
                    .Where(e => extraProjectIds.Contains(e.ProjectId) && !string.IsNullOrWhiteSpace(e.EnvelopeType))
                    .Select(e => e.EnvelopeType)
                    .FirstOrDefaultAsync();
                if (!string.IsNullOrWhiteSpace(envelopeTypeJson))
                {
                    try
                    {
                        using var doc = JsonDocument.Parse(envelopeTypeJson);
                        if (doc.RootElement.TryGetProperty("Inner", out var innerProp))
                            innerEnvelopeValue = innerProp.GetString()?.Trim();
                    }
                    catch { }
                }
            }

            // Load all fields from the Fields table
            var allFields = await _context.Fields.OrderBy(f => f.Name).ToListAsync();

            // Build single deduplicated flat list � priority: b ? e ? n ? eb ? x
            // Each column name appears exactly once, mapped to the highest-priority table prefix
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var result = new List<Dictionary<string, string>>();

            void AddColumns(IEnumerable<string> cols, string prefix)
            {
                foreach (var col in cols.OrderBy(x => x))
                {
                    if (excludeColumns.Contains(col)) continue;
                    if (!seen.Add(col)) continue;
                    result.Add(new Dictionary<string, string>
                    {
                        ["value"] = $"{prefix}{col}",
                        ["label"] = col
                    });
                }
            }

            AddColumns(boxColumns, "b.");
            AddColumns(envColumns, "e.");
            AddColumns(nrColumns, "n.");
            AddColumns(nrJsonKeys, "n.");
            AddColumns(envBreakageCols, "eb.");
            AddColumns(envBreakageJsonKeys.Where(k => !excludeColumns.Contains(k)), "eb.");

            // Add Fields table entries � always use n. prefix.
            // BuildSourceExpression handles the rest: direct column ? n.`col`, otherwise ? JSON_EXTRACT(n.NRDatas, ...)
            foreach (var field in allFields)
            {
                if (string.IsNullOrWhiteSpace(field.Name)) continue;
                if (excludeColumns.Contains(field.Name)) continue;
                if (seen.Contains(field.Name)) continue; // already present, skip duplicate

                seen.Add(field.Name);
                result.Add(new Dictionary<string, string>
                {
                    ["value"] = $"n.{field.Name}",
                    ["label"] = field.Name
                });
            }

            // Calculated fields for quantity sheet (sourced directly from SP result)
            result.Add(new Dictionary<string, string> { ["value"] = "c.TotalCenters", ["label"] = "TotalCenters" });
            result.Add(new Dictionary<string, string> { ["value"] = "c.TotalNodal", ["label"] = "TotalNodal" });

            // Virtual computed fields � not real columns, resolved to SQL expressions at report generation
            result.Add(new Dictionary<string, string>
            {
                ["value"] = "calc:BOX_RANGE",
                ["label"] = "Box Range"
            });
            result.Add(new Dictionary<string, string>
            {
                ["value"] = "calc:TOTAL_BOXES",
                ["label"] = "Total Boxes"
            });

            // x.Inner � store as eb.<actualValue> so the mapping saves the real field reference
            // e.g. if Inner = "E10", value = "eb.E10", label = "E10"
            if (!string.IsNullOrWhiteSpace(innerEnvelopeValue)
                && seen.Add(innerEnvelopeValue)
                && !result.Any(r => string.Equals(r["label"], innerEnvelopeValue, StringComparison.OrdinalIgnoreCase)))
            {
                result.Add(new Dictionary<string, string>
                {
                    ["value"] = $"eb.{innerEnvelopeValue}",
                    ["label"] = innerEnvelopeValue
                });
            }

            return Ok(result);
        }

        // GET: api/RPTTemplates/5
        [HttpGet("{id}")]
        public async Task<ActionResult<RPTTemplate>> GetRPTTemplate(int id)
        {
            var t = await _context.RPTTemplates.FindAsync(id);
            if (t == null) return NotFound();
            return t;
        }

        // PUT: api/RPTTemplates/5
        // body: { templateName, moduleIds, applyToAllVersions }
        [HttpPut("{id}")]
        public async Task<ActionResult> UpdateTemplate(int id, [FromBody] UpdateTemplateRequest req)
        {
            if (req == null)
                return BadRequest("Request body is required.");

            var template = await _context.RPTTemplates.FindAsync(id);
            if (template == null) return NotFound("Template not found.");

            var hasName = !string.IsNullOrWhiteSpace(req.TemplateName);
            if (req.TemplateName != null && !hasName)
                return BadRequest("templateName cannot be empty.");

            var newName = hasName ? req.TemplateName.Trim() : template.TemplateName;
            var updateAll = req.ApplyToAllVersions;

            if (updateAll)
            {
                var scopeQuery = _context.RPTTemplates
                    .Where(t => t.TypeId == template.TypeId
                                && t.TemplateName == template.TemplateName
                                && t.GroupId == template.GroupId
                                && t.ProjectId == template.ProjectId);

                var items = await scopeQuery.ToListAsync();
                foreach (var item in items)
                {
                    if (hasName) item.TemplateName = newName;
                    if (req.ModuleIds != null) item.ModuleIds = req.ModuleIds;
                    item.UpdatedDate = DateTime.Now;
                }
            }
            else
            {
                if (hasName) template.TemplateName = newName;
                if (req.ModuleIds != null) template.ModuleIds = req.ModuleIds;
                template.UpdatedDate = DateTime.Now;
            }

            await _context.SaveChangesAsync();
            _loggerService.LogEvent($"Updated template {id} (applyToAll={updateAll})", "RPTTemplate", LogHelper.GetTriggeredBy(User), 0);
            return Ok(new { templateId = template.TemplateId, message = "Template updated." });
        }

        // POST: api/RPTTemplates/{id}/activate
        // Marks a specific version as active for its scope (standard/group/project)
        [HttpPost("{id}/activate")]
        public async Task<ActionResult> ActivateTemplate(int id)
        {
            var template = await _context.RPTTemplates.FindAsync(id);
            if (template == null) return NotFound("Template not found.");

            var scopeQuery = _context.RPTTemplates
                .Where(t => t.TypeId == template.TypeId
                            && t.TemplateName == template.TemplateName
                            && t.GroupId == template.GroupId
                            && t.ProjectId == template.ProjectId);

            var items = await scopeQuery.ToListAsync();
            if (items.Count == 0)
                return NotFound("Template scope not found.");

            var now = DateTime.Now;
            foreach (var item in items)
            {
                item.IsActive = item.TemplateId == template.TemplateId;
                item.UpdatedDate = now;
            }

            await _context.SaveChangesAsync();
            _loggerService.LogEvent($"Activated template {id}", "RPTTemplate", LogHelper.GetTriggeredBy(User), 0);
            return Ok(new { templateId = template.TemplateId, message = "Template activated." });
        }

        // POST: api/RPTTemplates/import-from-group
        // Copy all active templates from a source group+type to a new group+type
        [HttpPost("import-from-group")]
        public async Task<ActionResult> ImportFromGroup([FromBody] ImportGroupRequest req)
        {
            var sourceScope = (req.SourceScope ?? "group").Trim().ToLowerInvariant();
            var sourceTypeId = NormalizeNullableId(req.SourceTypeId);
            var targetProjectId = NormalizeNullableId(req.TargetProjectId);
            var targetGroupId = NormalizeNullableId(req.TargetGroupId);
            var targetTypeId = NormalizeNullableId(req.TargetTypeId);
            var uploadedByUserId = LogHelper.GetTriggeredBy(User);

            if (targetProjectId.HasValue)
            {
                var targetProject = await _context.Projects
                    .FirstOrDefaultAsync(p => p.ProjectId == targetProjectId.Value);
                if (targetProject == null)
                    return NotFound("Target project not found.");

                targetGroupId = NormalizeNullableId(targetProject.GroupId);
                targetTypeId = NormalizeNullableId(targetProject.TypeId);

                if (!targetGroupId.HasValue || !targetTypeId.HasValue)
                    return BadRequest("Target project is missing GroupId/TypeId.");
            }
            else
            {
                if (!targetGroupId.HasValue || !targetTypeId.HasValue)
                    return BadRequest("TargetGroupId and TargetTypeId are required.");
            }

            List<RPTTemplate> sourceTemplates;
            if (sourceScope == "standard")
            {
                var query = _context.RPTTemplates
                    .Where(t => t.GroupId == null && t.ProjectId == null && t.IsActive);
                if (sourceTypeId.HasValue)
                    query = query.Where(t => t.TypeId == sourceTypeId);
                sourceTemplates = await query.ToListAsync();
            }
            else if (sourceScope == "project")
            {
                if (!req.SourceProjectId.HasValue || req.SourceProjectId.Value <= 0)
                    return BadRequest("SourceProjectId is required for project imports.");

                var project = await _context.Projects
                    .FirstOrDefaultAsync(p => p.ProjectId == req.SourceProjectId.Value);
                if (project == null)
                    return NotFound("Source project not found.");

                var effectiveTypeId = sourceTypeId ?? (project.TypeId > 0 ? project.TypeId : (int?)null);
                var effectiveGroupId = project.GroupId > 0 ? project.GroupId : (int?)null;
                if (!effectiveTypeId.HasValue || !effectiveGroupId.HasValue)
                    return BadRequest("Source project is missing GroupId/TypeId.");

                sourceTemplates = await ResolveTemplatesForContext(
                    effectiveTypeId.Value,
                    effectiveGroupId.Value,
                    req.SourceProjectId.Value);
            }
            else
            {
                if (req.SourceGroupId <= 0)
                    return BadRequest("SourceGroupId is required for group imports.");

                var groupQuery = _context.RPTTemplates
                    .Where(t => t.GroupId == req.SourceGroupId && t.ProjectId == null && t.IsActive);
                if (sourceTypeId.HasValue)
                    groupQuery = groupQuery.Where(t => t.TypeId == sourceTypeId);

                var sourceGroupTemplates = await groupQuery.ToListAsync();

                sourceTemplates = new List<RPTTemplate>(sourceGroupTemplates);
                if (req.IncludeStandard)
                {
                    var standardQuery = _context.RPTTemplates
                        .Where(t => t.GroupId == null && t.ProjectId == null && t.IsActive);
                    if (sourceTypeId.HasValue)
                        standardQuery = standardQuery.Where(t => t.TypeId == sourceTypeId);

                    var standardTemplates = await standardQuery.ToListAsync();

                    var existingNames = new HashSet<string>(
                        sourceGroupTemplates.Select(t => t.TemplateName ?? string.Empty),
                        StringComparer.OrdinalIgnoreCase);

                    foreach (var std in standardTemplates)
                    {
                        if (!existingNames.Contains(std.TemplateName ?? string.Empty))
                            sourceTemplates.Add(std);
                    }
                }
            }

            if (!sourceTemplates.Any())
                return NotFound("No active templates found for source group/type.");

            var imported = new List<object>();

            foreach (var src in sourceTemplates)
            {
                // Deactivate any existing for target group+type+name
                var existing = await _context.RPTTemplates
                    .Where(t => t.GroupId == targetGroupId && t.TypeId == targetTypeId
                                && t.TemplateName == src.TemplateName
                                && t.ProjectId == (targetProjectId.HasValue ? targetProjectId : null)
                                && t.IsActive)
                    .ToListAsync();
                existing.ForEach(t => t.IsActive = false);

                var lastVersion = await _context.RPTTemplates
                    .Where(t => t.GroupId == targetGroupId && t.TypeId == targetTypeId
                                && t.TemplateName == src.TemplateName
                                && t.ProjectId == (targetProjectId.HasValue ? targetProjectId : null))
                    .MaxAsync(t => (int?)t.Version) ?? 0;

                var newTemplate = new RPTTemplate
                {
                    GroupId = targetGroupId,
                    TypeId = targetTypeId.Value,
                    ProjectId = targetProjectId,
                    UploadedByUserId = uploadedByUserId,
                    ModuleIds = src.ModuleIds,
                    TemplateName = src.TemplateName,
                    RPTFilePath = src.RPTFilePath,   // reuse same file
                    ParsedFieldsJson = src.ParsedFieldsJson,
                    Version = lastVersion + 1,
                    IsActive = true,
                    CreatedDate = DateTime.Now,
                    UpdatedDate = DateTime.Now
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
                            TemplateId = newTemplate.TemplateId,
                            MappingJson = srcMapping.MappingJson
                        });
                        await _context.SaveChangesAsync();
                    }
                }

                imported.Add(new { newTemplate.TemplateId, newTemplate.TemplateName, newTemplate.Version });
            }

            _loggerService.LogEvent($"Imported {imported.Count} template(s) from group {req.SourceGroupId} to group {targetGroupId}", "RPTTemplate", LogHelper.GetTriggeredBy(User), 0);
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
            _loggerService.LogEvent($"Saved mapping for templateId {id}", "RPTMapping", LogHelper.GetTriggeredBy(User), 0);
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
                ? System.Text.Json.JsonSerializer.Serialize(parsedFields)
                : null;
            t.UpdatedDate = DateTime.Now;
            await _context.SaveChangesAsync();
            _loggerService.LogEvent($"Parsed fields for templateId {id}", "RPTTemplate", LogHelper.GetTriggeredBy(User), 0);
            return Ok(new { templateId = id, parsedFields });
        }


        [HttpPost("upload")]
        public async Task<ActionResult> Upload(
            [FromForm] int? templateId,
            [FromForm] int? projectId,
            [FromForm] string? templateName,
            [FromForm] int? groupId,
            [FromForm] int? typeId,
            [FromForm] IFormFile file,
            [FromForm] bool forceUpload = false)
        {
            try
            {
                if (file == null || file.Length == 0)
                    return BadRequest("No file uploaded.");

                var ext = Path.GetExtension(file.FileName).ToLowerInvariant();
                if (ext != ".rpt")
                    return BadRequest("Only .rpt files allowed.");

                if (string.IsNullOrWhiteSpace(templateName))
                    return BadRequest("templateName is required.");

                // STEP 1: Save TEMP file
                var tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".rpt");

                using (var stream = new FileStream(tempPath, FileMode.Create))
                {
                    await file.CopyToAsync(stream);
                }

                // STEP 2: Resolve templateId
                if (!templateId.HasValue)
                {
                    var baseQuery = _context.RPTTemplates
                        .Where(x => x.TypeId == typeId && x.TemplateName == templateName && x.IsActive);

                    if (projectId.HasValue)
                        baseQuery = baseQuery.Where(x => x.ProjectId == projectId && x.GroupId == groupId);
                    else if (groupId.HasValue)
                        baseQuery = baseQuery.Where(x => x.GroupId == groupId && x.ProjectId == null);
                    else
                        baseQuery = baseQuery.Where(x => x.GroupId == null && x.ProjectId == null);

                    templateId = await baseQuery
                        .OrderByDescending(x => x.Version)
                        .Select(x => (int?)x.TemplateId)
                        .FirstOrDefaultAsync();
                }

                // STEP 3: CALL DESIGN MICROSERVICE
                DesignCheckResponse designCheck = null;

                if (templateId.HasValue && templateId > 0)
                {
                    using (var client = new HttpClient())
                    using (var content = new MultipartFormDataContent())
                    using (var fileStream = new FileStream(tempPath, FileMode.Open, FileAccess.Read))
                    {
                        var fileContent = new StreamContent(fileStream);

                        content.Add(fileContent, "file", file.FileName);
                        content.Add(new StringContent(templateId.Value.ToString()), "templateId");

                        var response = await client.PostAsync(
                            $"{_apiSettings.RptServiceUrl}/check-design?templateId={templateId.Value}",
                            content
                        );

                        if (!response.IsSuccessStatusCode)
                        {
                            var error = await response.Content.ReadAsStringAsync();
                            return StatusCode(500, "Design check failed: " + error);
                        }

                        var json = await response.Content.ReadAsStringAsync();
                        var token = JsonConvert.DeserializeObject<JToken>(json);

                        if (token is JArray arr)
                        {
                            var first = arr.FirstOrDefault();
                            if (first != null)
                                designCheck = first.ToObject<DesignCheckResponse>();
                        }
                        else if (token is JObject obj)
                        {
                            if (obj["designCheck"] != null)
                                designCheck = obj["designCheck"].ToObject<DesignCheckResponse>();
                            else
                                designCheck = obj.ToObject<DesignCheckResponse>();
                        }

                        if (designCheck != null && !forceUpload)
                        {
                            return Ok(new
                            {
                                success = true,
                                requireConfirmation = true,
                                message = designCheck.message,
                                designCheck
                            });
                        }
                    }
                }

                // STEP 5: SAVE FINAL FILE
                var uploadFolder = Path.Combine(_env.WebRootPath, "uploads");

                if (!Directory.Exists(uploadFolder))
                    Directory.CreateDirectory(uploadFolder);

                var fileName = $"{Guid.NewGuid()}_{file.FileName}";
                var finalPath = Path.Combine(uploadFolder, fileName);

                System.IO.File.Copy(tempPath, finalPath, true);

                int nextVersion = 1;

                var baseScopeQuery = _context.RPTTemplates
                    .Where(x => x.TypeId == typeId && x.TemplateName == templateName);

                if (projectId.HasValue)
                    baseScopeQuery = baseScopeQuery.Where(x => x.ProjectId == projectId && x.GroupId == groupId);
                else if (groupId.HasValue)
                    baseScopeQuery = baseScopeQuery.Where(x => x.GroupId == groupId && x.ProjectId == null);
                else
                    baseScopeQuery = baseScopeQuery.Where(x => x.GroupId == null && x.ProjectId == null);

                var lastVersion = await baseScopeQuery
                    .OrderByDescending(x => x.Version)
                    .Select(x => x.Version)
                    .FirstOrDefaultAsync();

                if (lastVersion > 0)
                {
                    nextVersion = lastVersion + 1;
                }

                // Call /fields endpoint for ParsedFieldsJson
                string parsedFieldsJson = null;
                try
                {
                    using var client = new HttpClient();
                    using var content = new MultipartFormDataContent();
                    using var stream = System.IO.File.OpenRead(finalPath);
                    var fileContent = new StreamContent(stream);
                    content.Add(fileContent, "file", file.FileName);
                    var response = await client.PostAsync($"{_apiSettings.RptServiceUrl}/fields", content);
                    if (response.IsSuccessStatusCode)
                    {
                        var fieldsJson = await response.Content.ReadAsStringAsync();
                        var fieldsToken = JToken.Parse(fieldsJson);
                        if (fieldsToken is JObject fieldsObj && fieldsObj["data"] is JArray dataArr)
                        {
                            var fieldsList = dataArr.Select(x => x.ToString()).ToList();
                            parsedFieldsJson = JsonConvert.SerializeObject(fieldsList);
                        }
                        else if (fieldsToken is JArray)
                        {
                            parsedFieldsJson = fieldsJson;
                        }
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine("Warning: Failed to fetch fields from python microservice: " + e.Message);
                }

                // Fetch full snapshot from python microservice
                string snapshotJson = null;
                try
                {
                    snapshotJson = await GetDesignSnapshotJson(finalPath, file.FileName);
                }
                catch (Exception e)
                {
                    Console.WriteLine("Warning: Failed to fetch design snapshot from python microservice: " + e.Message);
                }

                var existingTemplates = await baseScopeQuery.Where(x => x.IsActive).ToListAsync();
                foreach (var item in existingTemplates)
                {
                    item.IsActive = false;
                }

                var template = new RPTTemplate
                {
                    TemplateName = templateName,
                    GroupId = groupId,
                    TypeId = typeId ?? 0,
                    ProjectId = projectId,
                    RPTFilePath = Path.Combine("uploads", fileName),
                    ParsedFieldsJson = parsedFieldsJson,
                    DesignSnapshotJson = snapshotJson,
                    Version = nextVersion,
                    IsActive = true,
                    CreatedDate = DateTime.Now,
                    UpdatedDate = DateTime.Now
                };

                _context.RPTTemplates.Add(template);
                await _context.SaveChangesAsync();

                using (var client = new HttpClient())
                using (var content = new MultipartFormDataContent())
                using (var fs = new FileStream(finalPath, FileMode.Open, FileAccess.Read))
                {
                    var fileContent = new StreamContent(fs);
                    content.Add(fileContent, "file", file.FileName);

                    await client.PostAsync(
                        $"{_apiSettings.RptServiceUrl}/upload-final?templateId={template.TemplateId}",
                        content
                    );
                }

                if (System.IO.File.Exists(tempPath))
                    System.IO.File.Delete(tempPath);

                return Ok(new
                {
                    success = true,
                    templateId = template.TemplateId,
                    filePath = template.RPTFilePath,
                    designCheck
                });
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message);
            }
        }
        //GET: api/RPTTemplates/5/download
        [HttpGet("{id}/download")]
        public async Task<ActionResult> Download(int id)
        {
            var t = await _context.RPTTemplates.FindAsync(id);
            if (t == null || string.IsNullOrWhiteSpace(t.RPTFilePath))
                return NotFound();

            // Resolve relative path against wwwroot
            var webRoot = _env.WebRootPath ?? Path.Combine(Directory.GetCurrentDirectory(), "wwwroot");
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

        private async Task<string> GetDesignSnapshotJson(string filePath, string fileName)
        {
            using var client = _httpClientFactory.CreateClient();

            using var stream = System.IO.File.OpenRead(filePath);
            using var content = new MultipartFormDataContent();

            var fileContent = new StreamContent(stream);
            fileContent.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");

            content.Add(fileContent, "file", fileName);

            var response = await client.PostAsync($"{_apiSettings.RptServiceUrl}/design-snapshot", content);

            var json = await response.Content.ReadAsStringAsync();

            if (!response.IsSuccessStatusCode)
                throw new Exception("Design snapshot failed: " + json);

            return json; // already JSON snapshot
        }

        private async Task<dynamic> CheckDesignChanges(string filePath, string fileName, int templateId)
        {
            using var client = _httpClientFactory.CreateClient();

            using var stream = System.IO.File.OpenRead(filePath);
            using var content = new MultipartFormDataContent();

            var fileContent = new StreamContent(stream);
            fileContent.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");

            content.Add(fileContent, "file", fileName);

            var response = await client.PostAsync(
                $"{_apiSettings.RptServiceUrl}/check-design?templateId={templateId}",
                content
            );
            var responseString = await response.Content.ReadAsStringAsync();
            Console.WriteLine("Response: " + responseString);

            var result = JsonConvert.DeserializeObject<DesignCheckResponse>(responseString);
            var json = await response.Content.ReadAsStringAsync();

            if (!response.IsSuccessStatusCode)
                throw new Exception("Design check failed: " + json);

            return JsonConvert.DeserializeObject(json);
        }
        private static string ComputeFileHash(string path)
        {
            using var sha = SHA256.Create();
            using var stream = System.IO.File.OpenRead(path);
            var hash = sha.ComputeHash(stream);
            return BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
        }

        private static string? TryGetPreviousHash(RPTTemplate? prev, string webRoot)
        {
            if (prev == null || string.IsNullOrWhiteSpace(prev.RPTFilePath)) return null;
            var absolutePath = Path.Combine(webRoot, prev.RPTFilePath);
            if (!System.IO.File.Exists(absolutePath)) return null;
            return ComputeFileHash(absolutePath);
        }

        private int GetNextVersion(int templateId)
        {
            var lastVersion = _context.RPTTemplates
                .Where(x => x.TemplateId == templateId)
                .Select(x => x.Version)
                .DefaultIfEmpty(0)
                .Max();

            return lastVersion + 1;
        }
        private static bool AreFieldsEqual(string? prevFieldsJson, string? newFieldsJson)
        {
            var prev = ParseFieldList(prevFieldsJson);
            var next = ParseFieldList(newFieldsJson);
            if (prev.Count == 0 && next.Count == 0) return true;
            if (prev.Count != next.Count) return false;
            var setPrev = new HashSet<string>(prev, StringComparer.OrdinalIgnoreCase);
            var setNext = new HashSet<string>(next, StringComparer.OrdinalIgnoreCase);
            return setPrev.SetEquals(setNext);
        }

        private static List<string> ParseFieldList(string? json)
        {
            if (string.IsNullOrWhiteSpace(json)) return new List<string>();
            try
            {
                return System.Text.Json.JsonSerializer.Deserialize<List<string>>(json) ?? new List<string>();
            }
            catch
            {
                return new List<string>();
            }
        }

        private static int? NormalizeNullableId(int? id)
        {
            if (!id.HasValue) return null;
            return id.Value > 0 ? id : null;
        }

        private int? GetUserIdFromToken()
        {
            var token = Request.Headers["Authorization"].ToString()?.Replace("Bearer ", "");
            if (string.IsNullOrWhiteSpace(token)) return null;

            try
            {
                var handler = new JwtSecurityTokenHandler();
                if (!handler.CanReadToken(token)) return null;
                var jwt = handler.ReadToken(token) as JwtSecurityToken;
                if (jwt == null) return null;

                var claim = jwt.Claims.FirstOrDefault(c =>
                    c.Type == "userid" ||
                    c.Type == "http://schemas.xmlsoap.org/ws/2005/05/identity/claims/name" ||
                    c.Type == "sub");

                if (claim == null) return null;
                return int.TryParse(claim.Value, out var id) ? id : null;
            }
            catch
            {
                return null;
            }
        }

        private static int GetScopeRank(RPTTemplate template)
        {
            if (template.ProjectId.HasValue) return 3;
            if (template.GroupId.HasValue) return 2;
            return 1;
        }

        private async Task<List<RPTTemplate>> ResolveTemplatesForContext(int typeId, int groupId, int projectId)
        {
            var candidates = await _context.RPTTemplates
                .Where(t => t.TypeId == typeId && t.IsActive &&
                            ((t.ProjectId == projectId && t.GroupId == groupId)
                             || (t.GroupId == groupId && t.ProjectId == null)
                             || (t.GroupId == null && t.ProjectId == null)))
                .ToListAsync();

            var resolved = candidates
                .GroupBy(t => t.TemplateName ?? string.Empty, StringComparer.OrdinalIgnoreCase)
                .Select(g => g
                    .OrderByDescending(t => GetScopeRank(t))
                    .ThenByDescending(t => t.Version)
                    .First())
                .OrderBy(t => t.TemplateName)
                .ToList();

            return resolved;
        }

        private int? GetLastActiveTemplateId(int groupId, int typeId)
        {
            return _context.RPTTemplates
                .Where(x => x.GroupId == groupId && x.TypeId == typeId)
                .OrderByDescending(x => x.UpdatedDate)
                .Select(x => (int?)x.TemplateId)
                .FirstOrDefault();
        }

    }

    public class ImportGroupRequest
    {
        public int SourceGroupId { get; set; }
        public int? SourceTypeId { get; set; }
        public int TargetGroupId { get; set; }
        public int TargetTypeId { get; set; }
        public int? TargetProjectId { get; set; }
        public int? SourceProjectId { get; set; }
        public string? SourceScope { get; set; }
        public bool CopyMappings { get; set; } = true;
        public bool IncludeStandard { get; set; } = true;
    }

    public class SaveMappingRequest
    {
        public string MappingJson { get; set; }
    }

    public class UpdateTemplateRequest
    {
        public string? TemplateName { get; set; }
        public List<int>? ModuleIds { get; set; }
        public bool ApplyToAllVersions { get; set; } = true;
    }

    public class RptDesignFieldDto
    {
        public string FieldName { get; set; }
        public string DataType { get; set; }
    }
    public class DesignCheckResponse
    {
        public bool? changed { get; set; }
        public bool? isFirstVersion { get; set; }
        public string? message { get; set; }
        public List<string>? changes { get; set; }
    }

}

