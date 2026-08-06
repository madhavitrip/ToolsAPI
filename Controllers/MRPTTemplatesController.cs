using ERPToolsAPI.Data;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Options;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Tools.Models;
using Tools.Services;
using ERPToolsAPI.Models;
using System.Text.Json;

namespace Tools.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class MRPTTemplatesController : ControllerBase
    {
        private readonly ERPToolsDbContext _context;
        private readonly IWebHostEnvironment _env;
        private readonly ILoggerService _loggerService;

        public MRPTTemplatesController(
            ERPToolsDbContext context,
            IWebHostEnvironment env,
            ILoggerService loggerService)
        {
            _context = context;
            _env = env;
            _loggerService = loggerService;
        }

        [AllowAnonymous]
        [HttpGet]
        public async Task<ActionResult<IEnumerable<object>>> GetMRPTTemplates()
        {
            var templates = await _context.MRPTTemplates.ToListAsync();
            return Ok(templates);
        }

        [AllowAnonymous]
        [HttpGet("by-group")]
        public async Task<ActionResult<IEnumerable<object>>> GetByGroup(
            [FromQuery] int typeId,
            [FromQuery] int? groupId)
        {
            var query = _context.MRPTTemplates
                .Where(t => t.TypeId == typeId && t.IsDeleted == false);

            if (groupId.HasValue)
            {
                query = query.Where(t => t.GroupId == groupId);
            }
            else
            {
                query = query.Where(t => t.GroupId == null);
            }

            var templates = await query.ToListAsync();

            var grouped = templates.GroupBy(t => t.TemplateName)
                .Select(g =>
                {
                    var active = g.FirstOrDefault(x => x.IsActive == true);
                    return active ?? g.OrderByDescending(x => x.Version).First();
                })
                .ToList();

            var resultIds = grouped.Select(t => t.TemplateId).ToList();
            var mappings = await _context.RPTMappings.Where(m => resultIds.Contains(m.TemplateId)).ToListAsync();
            var mappedIds = mappings.Select(m => m.TemplateId).ToHashSet();

            var result = grouped.Select(t => new
            {
                t.TemplateId,
                t.GroupId,
                t.TypeId,
                t.TemplateName,
                t.SubName,
                t.Version,
                t.CreatedDate,
                t.UpdatedDate,
                t.IsActive,
                t.IsDeleted,
                t.ReportStatus,
                t.UploadedByUserId,
                t.ModuleIds,
                t.ParsedFieldsJson,
                t.RequiredFieldsJson,
                HasMapping = mappedIds.Contains(t.TemplateId)
            });

            return Ok(result);
        }

        [Authorize]
        [HttpPost("upload")]
        public async Task<ActionResult> Upload([FromForm] IFormFile file, [FromForm] int typeId, [FromForm] int? groupId, [FromForm] string templateName, [FromForm] string? subName)
        {
            if (file == null || file.Length == 0) return BadRequest("No file uploaded.");
            var ext = Path.GetExtension(file.FileName).ToLowerInvariant();
            if (ext != ".rpt") return BadRequest("Only .rpt files are accepted.");

            var uploadedByUserId = LogHelper.GetTriggeredBy(User);

            var scopeQuery = _context.MRPTTemplates.Where(t => t.TypeId == typeId && t.TemplateName == templateName);
            if (groupId.HasValue) scopeQuery = scopeQuery.Where(t => t.GroupId == groupId);
            else scopeQuery = scopeQuery.Where(t => t.GroupId == null);

            var lastVersion = await scopeQuery.MaxAsync(t => (int?)t.Version) ?? 0;

            var existingActive = await scopeQuery.Where(t => t.IsActive == true).ToListAsync();
            existingActive.ForEach(t => t.IsActive = false);

            var webRoot = _env.WebRootPath ?? Path.Combine(Directory.GetCurrentDirectory(), "wwwroot");
            var folderParts = new List<string> { webRoot, "rpt-templates" };
            string scopeSlug;
            
            if (groupId.HasValue)
            {
                folderParts.Add(groupId.Value.ToString());
                folderParts.Add(typeId.ToString());
                scopeSlug = $"master_g{groupId}_t{typeId}";
            }
            else
            {
                folderParts.Add("standard");
                folderParts.Add(typeId.ToString());
                scopeSlug = $"master_std_t{typeId}";
            }

            var folder = Path.Combine(folderParts.ToArray());
            Directory.CreateDirectory(folder);
            var fileName = $"{templateName}_{scopeSlug}_v{lastVersion + 1}{ext}";
            var absolutePath = Path.Combine(folder, fileName);
            var relativePath = Path.Combine(folderParts.Skip(1).ToArray());
            relativePath = Path.Combine(relativePath, fileName);

            using (var stream = new FileStream(absolutePath, FileMode.Create))
                await file.CopyToAsync(stream);

            // Parsing placeholder - assumes ParseFieldsAsync logic exists or can be skipped for now
            string parsedFieldsJson = "[]";
            string requiredFieldsJson = "[]";

            var template = new MRPTTemplate
            {
                GroupId = groupId,
                TypeId = typeId,
                TemplateName = templateName,
                SubName = subName,
                RPTFilePath = relativePath.Replace('\\', '/'),
                Version = lastVersion + 1,
                IsActive = true,
                UploadedByUserId = uploadedByUserId,
                CreatedDate = DateTime.Now,
                UpdatedDate = DateTime.Now,
                ParsedFieldsJson = parsedFieldsJson,
                RequiredFieldsJson = requiredFieldsJson
            };

            _context.MRPTTemplates.Add(template);
            await _context.SaveChangesAsync();

            // Carry over mapping from previous version if exists
            var previousTemplateId = await scopeQuery.OrderByDescending(t => t.Version).Select(t => (int?)t.TemplateId).FirstOrDefaultAsync();
            if (previousTemplateId.HasValue)
            {
                var prevMapping = await _context.RPTMappings.FirstOrDefaultAsync(m => m.TemplateId == previousTemplateId.Value);
                if (prevMapping != null && !string.IsNullOrEmpty(prevMapping.MappingJson))
                {
                    _context.RPTMappings.Add(new RPTMapping
                    {
                        TemplateId = template.TemplateId,
                        MappingJson = prevMapping.MappingJson
                    });
                    await _context.SaveChangesAsync();
                }
            }

            return Ok(new { template.TemplateId, template.TemplateName, template.Version });
        }

        [AllowAnonymous]
        [HttpGet("{id}/download")]
        public async Task<ActionResult> Download(int id)
        {
            var template = await _context.MRPTTemplates.FindAsync(id);
            if (template == null || string.IsNullOrEmpty(template.RPTFilePath))
                return NotFound("Template or file not found.");

            var webRoot = _env.WebRootPath ?? Path.Combine(Directory.GetCurrentDirectory(), "wwwroot");
            var filePath = Path.Combine(webRoot, template.RPTFilePath.Replace('/', Path.DirectorySeparatorChar).Replace('\\', Path.DirectorySeparatorChar).TrimStart(Path.DirectorySeparatorChar));

            if (!System.IO.File.Exists(filePath))
                return NotFound("File not found on disk.");

            var bytes = await System.IO.File.ReadAllBytesAsync(filePath);
            return File(bytes, "application/octet-stream", Path.GetFileName(filePath));
        }

        [Authorize]
        [HttpDelete("{id}/soft-delete")]
        public async Task<ActionResult> SoftDelete(int id)
        {
            var template = await _context.MRPTTemplates.FindAsync(id);
            if (template == null) return NotFound();

            var scopeQuery = _context.MRPTTemplates.Where(t => t.TypeId == template.TypeId && t.TemplateName == template.TemplateName && t.GroupId == template.GroupId);
            var items = await scopeQuery.ToListAsync();
            
            foreach (var item in items)
            {
                item.IsDeleted = true;
                item.IsActive = false;
                item.UpdatedDate = DateTime.Now;
            }

            await _context.SaveChangesAsync();
            return Ok();
        }

        [Authorize]
        [HttpPost("{id}/restore")]
        public async Task<ActionResult> Restore(int id)
        {
            var template = await _context.MRPTTemplates.FindAsync(id);
            if (template == null) return NotFound();

            var scopeQuery = _context.MRPTTemplates.Where(t => t.TypeId == template.TypeId && t.TemplateName == template.TemplateName && t.GroupId == template.GroupId);
            var items = await scopeQuery.ToListAsync();
            
            foreach (var item in items)
            {
                item.IsDeleted = false;
            }
            
            var latest = items.OrderByDescending(t => t.Version).FirstOrDefault();
            if (latest != null) latest.IsActive = true;

            await _context.SaveChangesAsync();
            return Ok();
        }

        [AllowAnonymous]
        [HttpGet("{id}/versions")]
        public async Task<ActionResult> GetVersions(int id)
        {
            var current = await _context.MRPTTemplates.FindAsync(id);
            if (current == null) return NotFound();

            var versions = await _context.MRPTTemplates
                .Where(t => t.TypeId == current.TypeId && t.TemplateName == current.TemplateName && t.GroupId == current.GroupId && t.IsDeleted == false)
                .OrderByDescending(t => t.Version)
                .Select(t => new { t.TemplateId, t.Version, t.CreatedDate, t.IsActive })
                .ToListAsync();

            return Ok(versions);
        }

        [Authorize]
        [HttpPost("{id}/activate")]
        public async Task<ActionResult> ActivateVersion(int id)
        {
            var target = await _context.MRPTTemplates.FindAsync(id);
            if (target == null) return NotFound();

            var all = await _context.MRPTTemplates
                .Where(t => t.TypeId == target.TypeId && t.TemplateName == target.TemplateName && t.GroupId == target.GroupId)
                .ToListAsync();

            foreach (var item in all) item.IsActive = false;
            target.IsActive = true;
            target.UpdatedDate = DateTime.Now;

            await _context.SaveChangesAsync();
            return Ok();
        }

        [AllowAnonymous]
        [HttpGet("{id}/mapping")]
        public async Task<ActionResult> GetMapping(int id)
        {
            var template = await _context.MRPTTemplates.FindAsync(id);
            if (template == null) return NotFound();

            var mapping = await _context.RPTMappings.FirstOrDefaultAsync(m => m.TemplateId == id);
            return Ok(new { templateId = id, mappingJson = mapping?.MappingJson });
        }

        public class MSaveMappingRequest
        {
            public string MappingJson { get; set; }
        }

        [Authorize]
        [HttpPost("{id}/mapping")]
        public async Task<ActionResult> SaveMapping(int id, [FromBody] MSaveMappingRequest req)
        {
            var template = await _context.MRPTTemplates.FindAsync(id);
            if (template == null) return NotFound();

            var mapping = await _context.RPTMappings.FirstOrDefaultAsync(m => m.TemplateId == id);
            if (mapping == null)
            {
                mapping = new RPTMapping { TemplateId = id };
                _context.RPTMappings.Add(mapping);
            }
            
            mapping.MappingJson = req.MappingJson;
            template.UpdatedDate = DateTime.Now;
            
            await _context.SaveChangesAsync();
            return Ok();
        }
    }
}
