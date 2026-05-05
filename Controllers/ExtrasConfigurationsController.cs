using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using ERPToolsAPI.Data;
using Tools.Models;
using static Tools.Controllers.ExtraEnvelopesController;
using Tools.Services;
using Microsoft.VisualStudio.Web.CodeGenerators.Mvc.Templates.Blazor;

namespace Tools.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExtrasConfigurationsController : ControllerBase
    {
        private readonly ERPToolsDbContext _context;
        private readonly ILoggerService _loggerService;
        public ExtrasConfigurationsController(ERPToolsDbContext context, ILoggerService loggerService)
        {
            _context = context;
            _loggerService = loggerService;
        }

        // GET: api/ExtrasConfigurations
        [HttpGet]
        public async Task<ActionResult<IEnumerable<ExtrasConfiguration>>> GetExtraConfigurations()
        {
            return await _context.ExtraConfigurations.ToListAsync();
        }


        // GET: api/ExtrasConfigurations/ByProjectId
        [HttpGet("ByProject/{projectId}")]
        public async Task<ActionResult<IEnumerable<ExtrasConfiguration>>> GetExtrasByProjectId(int projectId)
        {
            var extras = await _context.ExtraConfigurations
                .Where(e => e.ProjectId == projectId)
                .ToListAsync();

            if (extras == null || extras.Count == 0)
            {
                return NotFound(new { message = $"No configurations found for ProjectId: {projectId}" });
            }

            // Log what's being returned
            Console.WriteLine($" Returning {extras.Count} ExtrasConfiguration(s) for ProjectId {projectId}:");
            foreach (var extra in extras)
            {
                Console.WriteLine($"   ID: {extra.Id}, ExtraType: {extra.ExtraType}, Mode: {extra.Mode}, nodalValue: {(string.IsNullOrEmpty(extra.nodalValue) ? "NULL/EMPTY" : "HAS DATA")}");
            }

            return Ok(extras);
        }


        // GET: api/ExtrasConfigurations/5
        [HttpGet("{id}")]
        public async Task<ActionResult<ExtrasConfiguration>> GetExtrasConfiguration(int id)
        {
            var extrasConfiguration = await _context.ExtraConfigurations.FindAsync(id);

            if (extrasConfiguration == null)
            {
                return NotFound();
            }

            return extrasConfiguration;
        }

        [HttpGet("DistinctNodals/{projectId}")]
        public async Task<IActionResult> GetDistinctNodalsByProject(int projectId)
        {
            // Only filter by project and return distinct NodalCode from NRData where Status == true
            var nodalCodes = await _context.NRDatas
                .Where(n => n.ProjectId == projectId 
                         && n.Status == true 
                         && !string.IsNullOrEmpty(n.NodalCode))
                .Select(n => n.NodalCode)
                .Distinct()
                .OrderBy(n => n)
                .ToListAsync();
            return Ok(nodalCodes);
        }

        // PUT: api/ExtrasConfigurations/5
        // To protect from overposting attacks, see https://go.microsoft.com/fwlink/?linkid=2123754
        [HttpPut("{id}")]
        public async Task<IActionResult> PutExtrasConfiguration(int id, ExtrasConfiguration extrasConfiguration)
        {
            if (id != extrasConfiguration.Id)
            {
                return BadRequest();
            }

            _context.Entry(extrasConfiguration).State = EntityState.Modified;

            try
            {
                await _context.SaveChangesAsync();
                _loggerService.LogEvent($"Updated ExtrasConfiguration with ID {id}", "ExtrasConfiguration", LogHelper.GetTriggeredBy(User), extrasConfiguration.ProjectId);

            }
            catch (Exception ex)
            {
                if (!ExtrasConfigurationExists(id))
                {
                    _loggerService.LogEvent($"ExtrasConfiguration with ID {id} not found during updating", "ExtrasConfiguration", LogHelper.GetTriggeredBy(User), extrasConfiguration.ProjectId);

                    return NotFound();
                }
                else
                {
                    _loggerService.LogError("Error updating ExtrasConfiguration", ex.Message, nameof(ExtrasConfigurationsController));
                    return StatusCode(500, "Internal server error");
                }
            }

            return NoContent();
        }

        // POST: api/ExtrasConfigurations
        // To protect from overposting attacks, see https://go.microsoft.com/fwlink/?linkid=2123754
        [HttpPost]
        public async Task<ActionResult<ExtrasConfiguration>> PostExtrasConfiguration(ExtrasConfiguration extrasConfiguration)
        {
            try
            {
                // Log incoming request for debugging
                Console.WriteLine($" Received ExtrasConfiguration:");
                Console.WriteLine($"   ProjectId: {extrasConfiguration.ProjectId}");
                Console.WriteLine($"   ExtraType: {extrasConfiguration.ExtraType}");
                Console.WriteLine($"   Mode: {extrasConfiguration.Mode}");
                Console.WriteLine($"   Value: {extrasConfiguration.Value}");
                Console.WriteLine($"   nodalValue: {extrasConfiguration.nodalValue ?? "NULL"}");
                Console.WriteLine($"   EnvelopeType: {extrasConfiguration.EnvelopeType}");

                var extra = await _context.ExtraConfigurations
               .Where(x => x.ProjectId == extrasConfiguration.ProjectId && extrasConfiguration.ExtraType == x.ExtraType).ToListAsync();

                if (extra.Any())
                {

                    var projectId = extrasConfiguration.ProjectId;
                   
                    _context.ExtraConfigurations.RemoveRange(extra);
                    await _context.SaveChangesAsync();
                    Console.WriteLine($" Deleted {extra.Count} old ExtrasConfiguration record(s) for ProjectId {projectId}");
                    _loggerService.LogEvent($"Deleted {projectId} old ExtrasConfiguration record(s)",
                  "ExtrasConfiguration", LogHelper.GetTriggeredBy(User), extrasConfiguration.ProjectId);
                }


                _context.ExtraConfigurations.Add(extrasConfiguration);
                await _context.SaveChangesAsync();
                
                Console.WriteLine($" Saved ExtrasConfiguration with ID {extrasConfiguration.Id}");
                Console.WriteLine($"   nodalValue in DB: {extrasConfiguration.nodalValue ?? "NULL"}");
                
                _loggerService.LogEvent($"Created new ExtrasConfiguration with ProjectID {extrasConfiguration.ProjectId}", "ExtrasConfiguration", LogHelper.GetTriggeredBy(User), extrasConfiguration.ProjectId);

                return CreatedAtAction("GetExtrasConfiguration", new { id = extrasConfiguration.Id }, extrasConfiguration);
            }
            catch (DbUpdateConcurrencyException ex)
            {
                _loggerService.LogError("Concurrency error when saving ExtrasConfiguration", ex.Message, nameof(ExtrasConfigurationsController));
                return Conflict("Concurrency conflict occurred. The data may have been modified or deleted by another process.");
            }
            catch (Exception ex)
            {
                _loggerService.LogError("Error creating ExtrasConfiguration", ex.Message, nameof(ExtrasConfigurationsController));
                return StatusCode(500, "Internal server error");
            }
        }

        // DELETE: api/ExtrasConfigurations/5
        [HttpDelete("{id}")]
        public async Task<IActionResult> DeleteExtrasConfiguration(int id)
        {
            try
            {
                var extrasConfiguration = await _context.ExtraConfigurations.FirstOrDefaultAsync(s=>s.ProjectId==id);
                if (extrasConfiguration == null)
                {
                    return NotFound();
                }

                _context.ExtraConfigurations.Remove(extrasConfiguration);
                await _context.SaveChangesAsync();
                _loggerService.LogEvent($"Deleted a ExtrasConfiguration with ID {id}", "ExtrasConfiguration", LogHelper.GetTriggeredBy(User), extrasConfiguration.ProjectId);

                return NoContent();
            }
            catch (Exception ex)
            {
                _loggerService.LogError("Error deleting ExtrasConfiguration", ex.Message, nameof(ExtrasConfigurationsController));
                return StatusCode(500, "Internal server error");
            }

        }

        private bool ExtrasConfigurationExists(int id)
        {
            return _context.ExtraConfigurations.Any(e => e.Id == id);
        }
    }


}

