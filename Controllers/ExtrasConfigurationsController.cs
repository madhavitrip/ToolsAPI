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
                _loggerService.LogEvent($"Updated ExtrasConfiguration with ID {id}", "ExtrasConfiguration", User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0);

            }
            catch (Exception ex)
            {
                if (!ExtrasConfigurationExists(id))
                {
                    _loggerService.LogEvent($"ExtrasConfiguration with ID {id} not found during updating", "ExtrasConfiguration", User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0);

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
                var extra = await _context.ExtraConfigurations.FindAsync(extrasConfiguration.ProjectId);
                if (extra != null)
                {
                    _loggerService.LogEvent($"ProjectConfig for {extra.ProjectId} already exists", "ProjectConfig", User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0);
                    _context.ExtraConfigurations.Remove(extra);
                    await _context.SaveChangesAsync();
                }
                _context.ExtraConfigurations.Add(extrasConfiguration);
                await _context.SaveChangesAsync();
                //_loggerService.LogEvent($"Created new ExtrasConfiguration with ProjectID {extrasConfiguration.ProjectId}", "ExtrasConfiguration", User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0);

                return CreatedAtAction("GetExtrasConfiguration", new { id = extrasConfiguration.Id }, extrasConfiguration);
            }
            catch (Exception ex)
            {
              //  _loggerService.LogError("Error creating ExtrasConfiguration", ex.Message, nameof(ExtrasConfigurationsController));
                return StatusCode(500, "Internal server error");
            }
        }

        // DELETE: api/ExtrasConfigurations/5
        [HttpDelete("{id}")]
        public async Task<IActionResult> DeleteExtrasConfiguration(int id)
        {
            try
            {
                var extrasConfiguration = await _context.ExtraConfigurations.FindAsync(id);
                if (extrasConfiguration == null)
                {
                    return NotFound();
                }

                _context.ExtraConfigurations.Remove(extrasConfiguration);
                await _context.SaveChangesAsync();
                _loggerService.LogEvent($"Deleted a ExtrasConfiguration with ID {id}", "ExtrasConfiguration", User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0);

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
