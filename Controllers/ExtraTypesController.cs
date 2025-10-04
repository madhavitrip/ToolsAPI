using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using ERPToolsAPI.Data;
using Tools.Models;
using Tools.Services;
using Microsoft.CodeAnalysis;

namespace Tools.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExtraTypesController : ControllerBase
    {
        private readonly ERPToolsDbContext _context;
        private readonly ILoggerService _loggerService;
        public ExtraTypesController(ERPToolsDbContext context, ILoggerService loggerService)
        {
            _context = context;
            _loggerService = loggerService;
        }

        // GET: api/ExtraTypes
        [HttpGet]
        public async Task<ActionResult<IEnumerable<ExtraType>>> GetExtraType()
        {
            return await _context.ExtraType.ToListAsync();
        }

        // GET: api/ExtraTypes/5
        [HttpGet("{id}")]
        public async Task<ActionResult<ExtraType>> GetExtraType(int id)
        {
            var extraType = await _context.ExtraType.FindAsync(id);

            if (extraType == null)
            {
                return NotFound();
            }

            return extraType;
        }

        // PUT: api/ExtraTypes/5
        // To protect from overposting attacks, see https://go.microsoft.com/fwlink/?linkid=2123754
        [HttpPut("{id}")]
        public async Task<IActionResult> PutExtraType(int id, ExtraType extraType)
        {
            if (id != extraType.ExtraTypeId)
            {
                return BadRequest();
            }

            _context.Entry(extraType).State = EntityState.Modified;

            try
            {
                await _context.SaveChangesAsync();
                _loggerService.LogEvent($"Updated ExtraType for {id} ", "NRData", User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0);

            }
            catch (Exception ex)
            {
                if (!ExtraTypeExists(id))
                {
                    _loggerService.LogEvent($"ExtraType with ID {id} not found during updating", "ExtraType", User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0);
                    return NotFound();
                }
                else
                {
                    _loggerService.LogError("Error updating ExtraType", ex.Message, nameof(ExtraTypesController));
                    return StatusCode(500, "Internal server error");
                }
            }

            return NoContent();
        }

        // POST: api/ExtraTypes
        // To protect from overposting attacks, see https://go.microsoft.com/fwlink/?linkid=2123754
        [HttpPost]
        public async Task<ActionResult<ExtraType>> PostExtraType(ExtraType extraType)
        {
            try
            {
                _context.ExtraType.Add(extraType);
                await _context.SaveChangesAsync();
                _loggerService.LogEvent($"Created a new ExtraType with ID {extraType.ExtraTypeId}", "ExtraType", User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0);
                return CreatedAtAction("GetExtraType", new { id = extraType.ExtraTypeId }, extraType);
            }
            catch (Exception ex)
            {
                _loggerService.LogError("Error creating ExtraTypes", ex.Message, nameof(ExtraTypesController));
                return StatusCode(500, "Internal server error");
            }
        }

        // DELETE: api/ExtraTypes/5
        [HttpDelete("{id}")]
        public async Task<IActionResult> DeleteExtraType(int id)
        {
            try
            {
                var extraType = await _context.ExtraType.FindAsync(id);
                if (extraType == null)
                {
                    return NotFound();
                }

                _context.ExtraType.Remove(extraType);
                await _context.SaveChangesAsync();
                _loggerService.LogEvent($"Deleted a ExtraTypes with ID {id}", "ExtraTypes", User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0);

                return NoContent();
            }
            catch (Exception ex)
            {
                _loggerService.LogError("Error creating ExtraTypes", ex.Message, nameof(ExtraTypesController));
                return StatusCode(500, "Internal server error");
            }
        }

        private bool ExtraTypeExists(int id)
        {
            return _context.ExtraType.Any(e => e.ExtraTypeId == id);
        }
    }
}
