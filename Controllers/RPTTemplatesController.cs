using ERPToolsAPI.Data;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Options;
using System;
using System.Reflection;
using System.Net.Http.Headers;
using System.Text.Json;
using System.Security.Cryptography;
using System.IdentityModel.Tokens.Jwt;
using System.Security.Claims;
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
        public async Task<ActionResult> GetMappingOptions(int groupId, int typeId)
        {
            var excludeColumns = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "id",
                "projectid",
                "envelopebreakingresultid",
                "createdat",
                "uploadedbatch",
                "uploadbatch",
                "nrdataid",
                "extraid",
                "envelopid",
                "envelopeid",
                "status",
                "lotno"
            };

            List<string> FilterColumns(List<string> columns) =>
                columns.Where(c => !excludeColumns.Contains(c)).ToList();

            var nrColumns = FilterColumns(GetModelColumns<NRData>(exclude: new[] { "NRDatas" }));
            var envColumns = FilterColumns(GetModelColumns<EnvelopeBreakingResult>());
            var envBreakageColumns = FilterColumns(GetModelColumns<EnvelopeBreakage>());
            var boxColumns = FilterColumns(GetModelColumns<BoxBreakingResult>());

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

            var envBreakageQuery = _context.EnvelopeBreakages.AsQueryable();
            if (projectIds.Count > 0)
            {
                envBreakageQuery = envBreakageQuery.Where(e => projectIds.Contains(e.ProjectId));
            }

            var envBreakageInnerRows = await envBreakageQuery
                .Where(e => !string.IsNullOrWhiteSpace(e.InnerEnvelope))
                .Select(e => e.InnerEnvelope)
                .ToListAsync();

            foreach (var json in envBreakageInnerRows)
            {
                try
                {
                    using var doc = JsonDocument.Parse(json);
                    if (doc.RootElement.ValueKind != JsonValueKind.Object) continue;
                    foreach (var prop in doc.RootElement.EnumerateObject())
                    {
                        if (!string.IsNullOrWhiteSpace(prop.Name))
                            envBreakageJsonKeys.Add(prop.Name);
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
                envBreakageColumns = envBreakageColumns.OrderBy(x => x).ToList(),
                boxColumns = boxColumns.OrderBy(x => x).ToList(),
                nrJsonKeys = nrJsonKeys
                    .Where(k => !excludeColumns.Contains(k))
                    .OrderBy(x => x)
                    .ToList(),
                envBreakageJsonKeys = envBreakageJsonKeys
                    .Where(k => !excludeColumns.Contains(k))
                    .OrderBy(x => x)
                    .ToList()
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
            return Ok(new { templateId = template.TemplateId, message = "Template updated." });
        }

        // POST: api/RPTTemplates/upload
        // multipart/form-data: file (.rpt), typeId, templateName, optional groupId, optional projectId
        [HttpPost("upload")]
        public async Task<ActionResult> Upload([FromForm] int typeId, [FromForm] string templateName, IFormFile file,
            [FromForm] int? groupId, [FromForm] int? projectId, [FromForm] List<int>? moduleIds, [FromForm] bool forceUpload = false)
        {
            if (file == null || file.Length == 0)
                return BadRequest("No file uploaded.");

            var ext = Path.GetExtension(file.FileName).ToLowerInvariant();
            if (ext != ".rpt")
                return BadRequest("Only .rpt files are accepted.");

            if (string.IsNullOrWhiteSpace(templateName))
                return BadRequest("templateName is required.");

            groupId = NormalizeNullableId(groupId);
            projectId = NormalizeNullableId(projectId);
            if (projectId.HasValue && !groupId.HasValue)
                return BadRequest("groupId is required when projectId is provided.");

            // Determine next version for this scope+type+name combo
            var scopeQuery = _context.RPTTemplates
                .Where(t => t.TypeId == typeId && t.TemplateName == templateName);
            if (projectId.HasValue)
                scopeQuery = scopeQuery.Where(t => t.ProjectId == projectId && t.GroupId == groupId);
            else if (groupId.HasValue)
                scopeQuery = scopeQuery.Where(t => t.GroupId == groupId && t.ProjectId == null);
            else
                scopeQuery = scopeQuery.Where(t => t.GroupId == null && t.ProjectId == null);

            var lastVersion = await scopeQuery.MaxAsync(t => (int?)t.Version) ?? 0;

            // Save file to wwwroot/rpt-templates/{scope}/
            var webRoot = _env.WebRootPath ?? Path.Combine(Directory.GetCurrentDirectory(), "wwwroot");
            var folderParts = new List<string> { webRoot, "rpt-templates" };
            string scopeSlug;
            if (projectId.HasValue)
            {
                folderParts.Add(groupId.Value.ToString());
                folderParts.Add(typeId.ToString());
                folderParts.Add("projects");
                folderParts.Add(projectId.Value.ToString());
                scopeSlug = $"g{groupId}_t{typeId}_p{projectId}";
            }
            else if (groupId.HasValue)
            {
                folderParts.Add(groupId.Value.ToString());
                folderParts.Add(typeId.ToString());
                scopeSlug = $"g{groupId}_t{typeId}";
            }
            else
            {
                folderParts.Add("standard");
                folderParts.Add(typeId.ToString());
                scopeSlug = $"std_t{typeId}";
            }

            var folder = Path.Combine(folderParts.ToArray());
            Directory.CreateDirectory(folder);
            var fileName     = $"{templateName}_{scopeSlug}_v{lastVersion + 1}{ext}";
            var absolutePath = Path.Combine(folder, fileName);
            var relativePath = Path.Combine(folderParts.Skip(1).ToArray());
            relativePath = Path.Combine(relativePath, fileName);

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

            var previousActive = await scopeQuery
                .Where(t => t.IsActive)
                .OrderByDescending(t => t.Version)
                .FirstOrDefaultAsync();

            var newHash = ComputeFileHash(absolutePath);
            var previousHash = TryGetPreviousHash(previousActive, webRoot);

            var fileIdentical = previousHash != null
                && string.Equals(previousHash, newHash, StringComparison.OrdinalIgnoreCase);

            var fieldsSame = previousActive != null
                && AreFieldsEqual(previousActive.ParsedFieldsJson, parsedFieldsJson);

            var likelyLayoutOnly = !fileIdentical && fieldsSame && previousActive != null;

            if (fileIdentical && !forceUpload)
            {
                try
                {
                    if (System.IO.File.Exists(absolutePath))
                        System.IO.File.Delete(absolutePath);
                }
                catch
                {
                    // ignore cleanup errors
                }
                return Conflict(new
                {
                    message = "No changes detected between this file and the latest version.",
                    allowForceUpload = true,
                    changeSummary = new
                    {
                        fileIdentical,
                        fieldsSame,
                        likelyLayoutOnly
                    }
                });
            }

            // Deactivate previous versions (never delete)
            var previous = await scopeQuery.Where(t => t.IsActive).ToListAsync();
            previous.ForEach(t => t.IsActive = false);

            var uploadedByUserId = GetUserIdFromToken();

            var template = new RPTTemplate
            {
                GroupId      = groupId,
                TypeId       = typeId,
                ProjectId    = projectId,
                UploadedByUserId = uploadedByUserId,
                ModuleIds    = moduleIds != null && moduleIds.Count > 0 ? moduleIds : null,
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

            // If user skips mapping for new version, reuse mapping from previous version if available
            if (previous.Count > 0)
            {
                var previousTemplateId = previous
                    .OrderByDescending(t => t.Version)
                    .Select(t => (int?)t.TemplateId)
                    .FirstOrDefault();

                if (previousTemplateId.HasValue)
                {
                    var prevMapping = await _context.RPTMappings
                        .FirstOrDefaultAsync(m => m.TemplateId == previousTemplateId.Value);
                    if (prevMapping != null)
                    {
                        var exists = await _context.RPTMappings
                            .AnyAsync(m => m.TemplateId == template.TemplateId);
                        if (!exists)
                        {
                            _context.RPTMappings.Add(new RPTMapping
                            {
                                TemplateId = template.TemplateId,
                                MappingJson = prevMapping.MappingJson
                            });
                            await _context.SaveChangesAsync();
                        }
                    }
                }
            }
            else if (projectId.HasValue && groupId.HasValue)
            {
                // For first-time project overrides, reuse mapping from group/standard template if available
                var fallbackTemplate = await _context.RPTTemplates
                    .Where(t => t.TypeId == typeId
                                && t.TemplateName == templateName
                                && t.GroupId == groupId
                                && t.ProjectId == null
                                && t.IsActive)
                    .OrderByDescending(t => t.Version)
                    .FirstOrDefaultAsync();

                if (fallbackTemplate == null)
                {
                    fallbackTemplate = await _context.RPTTemplates
                        .Where(t => t.TypeId == typeId
                                    && t.TemplateName == templateName
                                    && t.GroupId == null
                                    && t.ProjectId == null
                                    && t.IsActive)
                        .OrderByDescending(t => t.Version)
                        .FirstOrDefaultAsync();
                }

                if (fallbackTemplate != null)
                {
                    var fallbackMapping = await _context.RPTMappings
                        .FirstOrDefaultAsync(m => m.TemplateId == fallbackTemplate.TemplateId);
                    if (fallbackMapping != null)
                    {
                        var exists = await _context.RPTMappings
                            .AnyAsync(m => m.TemplateId == template.TemplateId);
                        if (!exists)
                        {
                            _context.RPTMappings.Add(new RPTMapping
                            {
                                TemplateId = template.TemplateId,
                                MappingJson = fallbackMapping.MappingJson
                            });
                            await _context.SaveChangesAsync();
                        }
                    }
                }
            }

            return Ok(new
            {
                template.TemplateId,
                template.TemplateName,
                template.Version,
                template.RPTFilePath,
                parsedFields = parsedFields,
                parseError = parseError,
                changeSummary = new
                {
                    fileIdentical,
                    fieldsSame,
                    likelyLayoutOnly
                }
            });
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
            var uploadedByUserId = GetUserIdFromToken();

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
                    GroupId      = targetGroupId,
                    TypeId       = targetTypeId.Value,
                    ProjectId    = targetProjectId,
                    UploadedByUserId = uploadedByUserId,
                    ModuleIds    = src.ModuleIds,
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
                return JsonSerializer.Deserialize<List<string>>(json) ?? new List<string>();
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
                    c.Type == ClaimTypes.Name ||
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
    }

    public class ImportGroupRequest
    {
        public int SourceGroupId  { get; set; }
        public int? SourceTypeId   { get; set; }
        public int TargetGroupId  { get; set; }
        public int TargetTypeId   { get; set; }
        public int? TargetProjectId { get; set; }
        public int? SourceProjectId { get; set; }
        public string? SourceScope { get; set; }
        public bool CopyMappings  { get; set; } = true;
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
}
