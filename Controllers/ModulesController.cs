using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using ERPToolsAPI.Data;
using Tools.Models;
using Microsoft.CodeAnalysis;
using Tools.Services;
using static Tools.Controllers.ExtraEnvelopesController;

namespace Tools.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ModulesController : ControllerBase
    {
        private readonly ERPToolsDbContext _context;
        private readonly ILoggerService _loggerService;
        public ModulesController(ERPToolsDbContext context, ILoggerService loggerService)
        {
            _context = context;
            _loggerService = loggerService;
        }

        // GET: api/Modules
        [HttpGet]
        public async Task<ActionResult<IEnumerable<Module>>> GetModules()
        {
            return await _context.Modules.ToListAsync();
        }

        // GET: api/Modules/5
        [HttpGet("{id}")]
        public async Task<ActionResult<Module>> GetModule(int id)
        {
            var @module = await _context.Modules.FindAsync(id);

            if (@module == null)
            {
                return NotFound();
            }

            return @module;
        }

        // PUT: api/Modules/5
        // To protect from overposting attacks, see https://go.microsoft.com/fwlink/?linkid=2123754
        [HttpPut("{id}")]
        public async Task<IActionResult> PutModule(int id, Module @module)
        {
            if (id != @module.Id)
            {
                return BadRequest();
            }

            var existingModule = await _context.Modules
                .AsNoTracking()
                .FirstOrDefaultAsync(m => m.Id == id);

            if (existingModule == null)
            {
                var triggeredBy = LogHelper.GetTriggeredBy(User);
                _loggerService.LogEvent(
                    $"Module with ID {id} not found during updating",
                    "Module",
                    triggeredBy,
                    0,
                    LogHelper.ToJson(existingModule),
                    LogHelper.ToJson(@module)
                );
                return NotFound();
            }

            _context.Entry(@module).State = EntityState.Modified;

            try
            {
                await _context.SaveChangesAsync();
                var triggeredBy = LogHelper.GetTriggeredBy(User);
                _loggerService.LogEvent(
                    $"Updated Module for id {id}",
                    "Module",
                    triggeredBy,
                    0,
                    LogHelper.ToJson(existingModule),
                    LogHelper.ToJson(@module)
                );

            }
            catch (Exception ex)
            {
                _loggerService.LogError("Error updating Module", ex.Message, nameof(ModulesController));
                return StatusCode(500, "Internal server error");
            }

            return NoContent();
        }

        // POST: api/Modules
        // To protect from overposting attacks, see https://go.microsoft.com/fwlink/?linkid=2123754
        [HttpPost]
        public async Task<ActionResult<Module>> PostModule(Module @module)
        {
            try
            {
                _context.Modules.Add(@module);
                await _context.SaveChangesAsync();
                var triggeredBy = LogHelper.GetTriggeredBy(User);
                _loggerService.LogEvent(
                    $"Created a new Module with ID {@module.Id}",
                    "Module",
                    triggeredBy,
                    0,
                    string.Empty,
                    LogHelper.ToJson(@module)
                );

                return CreatedAtAction("GetModule", new { id = @module.Id }, @module);
            }
            catch (Exception ex)
            {
                _loggerService.LogError("Error creating Module", ex.Message, nameof(ModulesController));
                return StatusCode(500, "Internal server error");
            }
        }

        // DELETE: api/Modules/5
        [HttpDelete("{id}")]
        public async Task<IActionResult> DeleteModule(int id)
        {
            try
            {
                var @module = await _context.Modules.FindAsync(id);
                if (@module == null)
                {
                    return NotFound();
                }

                _context.Modules.Remove(@module);
                await _context.SaveChangesAsync();
                _loggerService.LogEvent($"Deleted a Module with ID {id}", "Module", User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0,0);

                return NoContent();
            }
            catch (Exception ex)
            {
                _loggerService.LogError("Error deleting Module", ex.Message, nameof(ModulesController));
                return StatusCode(500, "Internal server error");
            }
        }

    }
}
