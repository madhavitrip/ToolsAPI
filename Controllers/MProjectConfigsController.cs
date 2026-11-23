using ERPToolsAPI.Data;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Tools.Models;
using Tools.Services;
using Tools.Middleware;
using System.Text.Json;

namespace Tools.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class MProjectConfigsController : ControllerBase
    {
        private readonly ERPToolsDbContext _context;
        private readonly ILoggerService _loggerService;
        public MProjectConfigsController(ERPToolsDbContext context, ILoggerService loggerService)
        {
            _context = context;
            _loggerService = loggerService;
        }

        // POST: api/MProjectConfigs
        [HttpPost]
        [RequireMasterAuth(Module = "MProjectConfigs", Operation = "SAVE MASTER")]
        public async Task<ActionResult<MProjectConfigs>> PostMProjectConfig(MProjectConfigs projectConfig)
        {
            try
            {
                var config = await _context.MProjectConfigs
                    .Where(p => p.TypeId == projectConfig.TypeId && p.GroupId == projectConfig.GroupId)
                    .FirstOrDefaultAsync();

                string oldValue = null;
                if (config != null)
                {
                    // Serialize old values for logging
                    oldValue = JsonSerializer.Serialize(new
                    {
                        config.Id,
                        config.GroupId,
                        config.TypeId,
                        config.Modules,
                        config.Envelope,
                        config.BoxBreakingCriteria,
                        config.DuplicateRemoveFields,
                        config.SortingBoxReport,
                        config.EnvelopeMakingCriteria,
                        config.BoxCapacity,
                        config.BoxNumber,
                        config.OmrSerialNumber,
                        config.ResetOnSymbolChange,
                        config.IsInnerBundlingDone,
                        config.ResetOmrSerialOnCatchChange,
                        config.BookletSerialNumber,
                        config.ResetBookletSerialOnCatchChange,
                        config.DuplicateCriteria,
                        config.Enhancement,
                        config.InnerBundlingCriteria
                    });

                    await _loggerService.LogEventAsync(
                        $"MProjectConfig for TypeId {config.TypeId} and GroupId {config.GroupId} already exists",
                        "MProjectConfigs",
                        LogHelper.GetTriggeredBy(User),
                        config.GroupId,
                        oldValue,
                        null
                    );

                    var folderName = $"{config.TypeId}_{config.GroupId}";
                    var reportPath = FileStorageHelper.GetGroupFolder(folderName);
                    if (Directory.Exists(reportPath))
                    {
                        try { Directory.Delete(reportPath, true); } catch { }
                    }

                    var legacyReportPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", folderName);
                    if (Directory.Exists(legacyReportPath))
                    {
                        try { Directory.Delete(legacyReportPath, true); } catch { }
                    }

                    _context.MProjectConfigs.Remove(config);
                    await _context.SaveChangesAsync();
                }

                _context.MProjectConfigs.Add(projectConfig);
                await _context.SaveChangesAsync();

                // Serialize new values for logging
                var newValue = JsonSerializer.Serialize(new
                {
                    projectConfig.Id,
                    projectConfig.GroupId,
                    projectConfig.TypeId,
                    projectConfig.Modules,
                    projectConfig.Envelope,
                    projectConfig.BoxBreakingCriteria,
                    projectConfig.DuplicateRemoveFields,
                    projectConfig.SortingBoxReport,
                    projectConfig.EnvelopeMakingCriteria,
                    projectConfig.BoxCapacity,
                    projectConfig.BoxNumber,
                    projectConfig.OmrSerialNumber,
                    projectConfig.ResetOnSymbolChange,
                    projectConfig.IsInnerBundlingDone,
                    projectConfig.ResetOmrSerialOnCatchChange,
                    projectConfig.BookletSerialNumber,
                    projectConfig.ResetBookletSerialOnCatchChange,
                    projectConfig.DuplicateCriteria,
                    projectConfig.Enhancement,
                    projectConfig.InnerBundlingCriteria
                });

                await _loggerService.LogEventAsync(
                    $"Created a new MProjectConfig with TypeId {projectConfig.TypeId} and GroupId {projectConfig.GroupId}",
                    "MProjectConfigs",
                    LogHelper.GetTriggeredBy(User),
                    projectConfig.GroupId,
                    oldValue,
                    newValue
                );

                return Ok(projectConfig);
            }
            catch (Exception ex)
            {
                await _loggerService.LogErrorAsync("Error creating MProjectConfigs", ex.Message, nameof(MProjectConfigsController));
                return StatusCode(500, "Internal server error");
            }
        }

        [HttpGet]
        public async Task<ActionResult<IEnumerable<int>>> GetAllGroups()
        {
            try
            {
                var groupIds = await _context.MProjectConfigs
                    .Select(x => x.GroupId)
                    .Distinct()
                    .ToListAsync();

                return Ok(groupIds);
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message);
            }
        }

        [HttpGet("SelectedGroup")]
        public async Task<ActionResult<IEnumerable<int>>> GetAllTypeOfGroup(int GroupId)
        {
            try
            {
                var groupIds = await _context.MProjectConfigs.Where(s=>s.GroupId == GroupId)
                    .Select(x => x.TypeId)
                    .Distinct()
                    .ToListAsync();

                return Ok(groupIds);
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message);
            }
        }

        [HttpGet("ByTypeGroup/{typeId}/{groupId}")]
        public async Task<ActionResult<MProjectConfigs>> GetMProjectConfigByTypeGroup(int typeId, int groupId)
        {
            var config = await _context.MProjectConfigs
                .FirstOrDefaultAsync(p => p.TypeId == typeId && p.GroupId == groupId);

            if (config == null)
            {
                return NotFound(new { message = $"No configuration found for TypeId: {typeId} and GroupId: {groupId}" });
            }

            return Ok(config);
        }

        // GET: api/MProjectConfigs/Report
        [HttpGet("Report")]
        public async Task<ActionResult<object>> GetMasterConfigReport(
            [FromQuery] int pageNumber = 1,
            [FromQuery] int pageSize = 15,
            [FromQuery] string? search = null,
            [FromQuery] int? groupId = null,
            [FromQuery] int? typeId = null,
            [FromQuery] string? sortBy = null,
            [FromQuery] string? sortOrder = null)
        {
            try
            {
                var configs = await _context.MProjectConfigs.ToListAsync();
                var modules = await _context.Modules.ToListAsync();
                var fields = await _context.Fields.ToListAsync();

                var fieldMap = fields.ToDictionary(f => f.FieldId, f => f.Name);
                var moduleMap = modules.ToDictionary(m => m.Id, m => m.Name);

                string ResolveFieldNames(List<int>? ids) =>
                    ids == null || ids.Count == 0
                        ? "-"
                        : string.Join(", ", ids.Select(id => fieldMap.TryGetValue(id, out var n) ? n : id.ToString()));

                string ResolveModuleNames(List<int>? ids) =>
                    ids == null || ids.Count == 0
                        ? "-"
                        : string.Join(", ", ids.Select(id => moduleMap.TryGetValue(id, out var n) ? n : id.ToString()));

                // Apply GroupId and TypeId filters on the raw configs (before projection)
                if (groupId.HasValue)
                    configs = configs.Where(c => c.GroupId == groupId.Value).ToList();

                if (typeId.HasValue)
                    configs = configs.Where(c => c.TypeId == typeId.Value).ToList();

                // Project to resolved DTOs
                var projected = configs.Select(c =>
                {
                    string innerEnvelope = "-";
                    string outerEnvelope = "-";
                    try
                    {
                        if (!string.IsNullOrWhiteSpace(c.Envelope))
                        {
                            var env = JsonSerializer.Deserialize<JsonElement>(c.Envelope);
                            innerEnvelope = env.TryGetProperty("Inner", out var inner) ? (inner.GetString() ?? "-") : "-";
                            outerEnvelope = env.TryGetProperty("Outer", out var outer) ? (outer.GetString() ?? "-") : "-";
                            if (string.IsNullOrEmpty(innerEnvelope)) innerEnvelope = "-";
                            if (string.IsNullOrEmpty(outerEnvelope)) outerEnvelope = "-";
                        }
                    }
                    catch { }

                    return new
                    {
                        c.Id,
                        c.GroupId,
                        c.TypeId,
                        GroupName = c.GroupId.ToString(),
                        TypeName = c.TypeId.ToString(),
                        Modules = ResolveModuleNames(c.Modules),
                        InnerEnvelope = innerEnvelope,
                        OuterEnvelope = outerEnvelope,
                        BoxBreakingCriteria = ResolveFieldNames(c.BoxBreakingCriteria),
                        DuplicateRemoveFields = ResolveFieldNames(c.DuplicateRemoveFields),
                        SortingBoxReport = ResolveFieldNames(c.SortingBoxReport),
                        EnvelopeMakingCriteria = ResolveFieldNames(c.EnvelopeMakingCriteria),
                        DuplicateCriteria = ResolveFieldNames(c.DuplicateCriteria),
                        InnerBundlingCriteria = ResolveFieldNames(c.InnerBundlingCriteria),
                        c.BoxCapacity,
                        c.Enhancement,
                        c.BoxNumber,
                        c.OmrSerialNumber,
                        c.ResetOnSymbolChange,
                        c.IsInnerBundlingDone,
                        c.ResetOmrSerialOnCatchChange,
                        c.BookletSerialNumber,
                        c.ResetBookletSerialOnCatchChange,
                    };
                }).ToList();

                // Apply global search across key resolved string fields
                if (!string.IsNullOrWhiteSpace(search))
                {
                    var s = search.ToLower();
                    projected = projected.Where(r =>
                        r.GroupId.ToString().Contains(s) ||
                        r.TypeId.ToString().Contains(s) ||
                        r.Modules.ToLower().Contains(s) ||
                        r.InnerEnvelope.ToLower().Contains(s) ||
                        r.OuterEnvelope.ToLower().Contains(s) ||
                        r.BoxBreakingCriteria.ToLower().Contains(s) ||
                        r.DuplicateRemoveFields.ToLower().Contains(s) ||
                        r.SortingBoxReport.ToLower().Contains(s) ||
                        r.EnvelopeMakingCriteria.ToLower().Contains(s) ||
                        r.DuplicateCriteria.ToLower().Contains(s) ||
                        r.InnerBundlingCriteria.ToLower().Contains(s)
                    ).ToList();
                }

                // Apply sorting
                bool isAsc = string.IsNullOrEmpty(sortOrder) || sortOrder.ToLower() == "asc" || sortOrder.ToLower() == "ascend";
                if (!string.IsNullOrEmpty(sortBy))
                {
                    projected = sortBy.ToLower() switch
                    {
                        "groupid" or "groupname" => isAsc ? projected.OrderBy(r => r.GroupId).ToList() : projected.OrderByDescending(r => r.GroupId).ToList(),
                        "typeid" or "typename" => isAsc ? projected.OrderBy(r => r.TypeId).ToList() : projected.OrderByDescending(r => r.TypeId).ToList(),
                        "modules" => isAsc ? projected.OrderBy(r => r.Modules).ToList() : projected.OrderByDescending(r => r.Modules).ToList(),
                        "innerenvelope" => isAsc ? projected.OrderBy(r => r.InnerEnvelope).ToList() : projected.OrderByDescending(r => r.InnerEnvelope).ToList(),
                        "outerenvelope" => isAsc ? projected.OrderBy(r => r.OuterEnvelope).ToList() : projected.OrderByDescending(r => r.OuterEnvelope).ToList(),
                        "boxbreakingcriteria" => isAsc ? projected.OrderBy(r => r.BoxBreakingCriteria).ToList() : projected.OrderByDescending(r => r.BoxBreakingCriteria).ToList(),
                        "duplicateremovefields" => isAsc ? projected.OrderBy(r => r.DuplicateRemoveFields).ToList() : projected.OrderByDescending(r => r.DuplicateRemoveFields).ToList(),
                        "sortingboxreport" => isAsc ? projected.OrderBy(r => r.SortingBoxReport).ToList() : projected.OrderByDescending(r => r.SortingBoxReport).ToList(),
                        "envelopemakingcriteria" => isAsc ? projected.OrderBy(r => r.EnvelopeMakingCriteria).ToList() : projected.OrderByDescending(r => r.EnvelopeMakingCriteria).ToList(),
                        "duplicatecriteria" => isAsc ? projected.OrderBy(r => r.DuplicateCriteria).ToList() : projected.OrderByDescending(r => r.DuplicateCriteria).ToList(),
                        "innerbundlingcriteria" => isAsc ? projected.OrderBy(r => r.InnerBundlingCriteria).ToList() : projected.OrderByDescending(r => r.InnerBundlingCriteria).ToList(),
                        "boxcapacity" => isAsc ? projected.OrderBy(r => r.BoxCapacity).ToList() : projected.OrderByDescending(r => r.BoxCapacity).ToList(),
                        "enhancement" => isAsc ? projected.OrderBy(r => r.Enhancement).ToList() : projected.OrderByDescending(r => r.Enhancement).ToList(),
                        "boxnumber" => isAsc ? projected.OrderBy(r => r.BoxNumber).ToList() : projected.OrderByDescending(r => r.BoxNumber).ToList(),
                        "omrserialnumber" => isAsc ? projected.OrderBy(r => r.OmrSerialNumber).ToList() : projected.OrderByDescending(r => r.OmrSerialNumber).ToList(),
                        "bookletserialnumber" or "bsn" => isAsc ? projected.OrderBy(r => r.BookletSerialNumber).ToList() : projected.OrderByDescending(r => r.BookletSerialNumber).ToList(),
                        "resetonsymbolchange" or "rosc" => isAsc ? projected.OrderBy(r => r.ResetOnSymbolChange).ToList() : projected.OrderByDescending(r => r.ResetOnSymbolChange).ToList(),
                        "isinnerbundlingdone" or "iibd" => isAsc ? projected.OrderBy(r => r.IsInnerBundlingDone).ToList() : projected.OrderByDescending(r => r.IsInnerBundlingDone).ToList(),
                        "resetomrserialoncatchchange" or "rocc" => isAsc ? projected.OrderBy(r => r.ResetOmrSerialOnCatchChange).ToList() : projected.OrderByDescending(r => r.ResetOmrSerialOnCatchChange).ToList(),
                        "resetbookletserialoncatchchange" or "rbcc" => isAsc ? projected.OrderBy(r => r.ResetBookletSerialOnCatchChange).ToList() : projected.OrderByDescending(r => r.ResetBookletSerialOnCatchChange).ToList(),
                        _ => projected.OrderBy(r => r.Id).ToList()
                    };
                }
                else
                {
                    projected = projected.OrderBy(r => r.Id).ToList();
                }

                // Pagination
                int total = projected.Count;
                int safePageNumber = Math.Max(1, pageNumber);
                int safePageSize = Math.Max(1, pageSize);
                var pagedData = projected.Skip((safePageNumber - 1) * safePageSize).Take(safePageSize).ToList();

                return Ok(new
                {
                    data = pagedData,
                    total,
                    currentPage = safePageNumber,
                    pageSize = safePageSize,
                });
            }
            catch (Exception ex)
            {
                await _loggerService.LogErrorAsync("Error fetching MProjectConfigs report", ex.Message, nameof(MProjectConfigsController));
                return StatusCode(500, "Internal server error");
            }
        }

    }

}
