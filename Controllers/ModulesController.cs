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
        public async Task<ActionResult<IEnumerable<object>>> GetModules()
        {
            // get all modules from DB
            var modules = await _context.Modules.ToListAsync();

            var result = modules.Select(m => new
            {
                m.Id,
                m.Name,
                ParentModuleIds = m.ParentModuleIds,
                ParentModuleNames = m.ParentModuleIds.Any()
                    ? modules.Where(p => m.ParentModuleIds.Contains(p.Id))
                             .Select(p => p.Name)
                             .ToList()
                    : new List<string>()
            }).ToList();

            return result;
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

            _context.Entry(@module).State = EntityState.Modified;

            try
            {
                await _context.SaveChangesAsync();
                await _loggerService.LogEventAsync($"Updated Module for id {id}", "Module", LogHelper.GetTriggeredBy(User), 0);

            }
            catch (Exception ex)
            {
                if (!ModuleExists(id))
                {
                    await _loggerService.LogEventAsync($"Module with ID {id} not found during updating", "Module", LogHelper.GetTriggeredBy(User), 0);

                    return NotFound();
                }
                else
                {

                    await _loggerService.LogErrorAsync("Error updating Module", ex.Message, nameof(ModulesController));
                    return StatusCode(500, "Internal server error");

                }
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
                await _loggerService.LogEventAsync($"Created a new Module with ID {module.Id}", "Module", LogHelper.GetTriggeredBy(User), 0);

                return CreatedAtAction("GetModule", new { id = @module.Id }, @module);
            }
            catch (Exception ex)
            {
                await _loggerService.LogErrorAsync("Error creating Module", ex.Message, nameof(ModulesController));
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
                await _loggerService.LogEventAsync($"Deleted a Module with ID {id}", "Module", LogHelper.GetTriggeredBy(User), 0);

                return NoContent();
            }
            catch (Exception ex)
            {
                await _loggerService.LogErrorAsync("Error deleting Module", ex.Message, nameof(ModulesController));
                return StatusCode(500, "Internal server error");
            }
        }

        private bool ModuleExists(int id)
        {
            return _context.Modules.Any(e => e.Id == id);
        }
    }
}

