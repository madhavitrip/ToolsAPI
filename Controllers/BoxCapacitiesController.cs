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

namespace Tools.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class BoxCapacitiesController : ControllerBase
    {
        private readonly ERPToolsDbContext _context;
        private readonly ILoggerService _loggerService;
        public BoxCapacitiesController(ERPToolsDbContext context, ILoggerService loggerService)
        {
            _context = context;
            _loggerService = loggerService;
        }

        // GET: api/BoxCapacities
        [HttpGet]
        public async Task<ActionResult<IEnumerable<BoxCapacity>>> GetBoxCapacity()
        {
            return await _context.BoxCapacity.ToListAsync();
        }

        // GET: api/BoxCapacities/5
        [HttpGet("{id}")]
        public async Task<ActionResult<BoxCapacity>> GetBoxCapacity(int id)
        {
            var boxCapacity = await _context.BoxCapacity.FindAsync(id);

            if (boxCapacity == null)
            {
                return NotFound();
            }

            return boxCapacity;
        }

        // PUT: api/BoxCapacities/5
        // To protect from overposting attacks, see https://go.microsoft.com/fwlink/?linkid=2123754
        [HttpPut("{id}")]
        public async Task<IActionResult> PutBoxCapacity(int id, BoxCapacity boxCapacity)
        {
            if (id != boxCapacity.BoxCapacityId)
            {
                return BadRequest();
            }

            _context.Entry(boxCapacity).State = EntityState.Modified;

            try
            {
                _loggerService.LogEvent($"Updated BoxCapacity with ID {id}", "BoxCapacity", User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0, 0);
                await _context.SaveChangesAsync();
            }
            catch (Exception ex)
            {
                if (!BoxCapacityExists(id))
                {
                    _loggerService.LogEvent($"BoxCapacity with ID {id} not found during updating", "BoxCapacity", User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0, 0);

                    return NotFound();
                }
                else
                {
                    _loggerService.LogError("Error updating BoxCapacity", ex.Message, nameof(BoxCapacitiesController));
                    return StatusCode(500, "Internal server error");
                }
            }

            return NoContent();
        }

        // POST: api/BoxCapacities
        // To protect from overposting attacks, see https://go.microsoft.com/fwlink/?linkid=2123754
        [HttpPost]
        public async Task<ActionResult<BoxCapacity>> PostBoxCapacity(BoxCapacity boxCapacity)
        {
            try
            {
                _context.BoxCapacity.Add(boxCapacity);
                await _context.SaveChangesAsync();
                _loggerService.LogEvent($"BoxCapacity with {boxCapacity} has been created", "BoxCapacity", User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0, 0);
                return CreatedAtAction("GetBoxCapacity", new { id = boxCapacity.BoxCapacityId }, boxCapacity);
            }
            catch (Exception ex)
            {
                _loggerService.LogError("Error creating BoxCapacity", ex.Message, nameof(BoxCapacitiesController));
                return StatusCode(500, "Internal server error");

            }
        }

        // DELETE: api/BoxCapacities/5
        [HttpDelete("{id}")]
        public async Task<IActionResult> DeleteBoxCapacity(int id)
        {
            try
            {
                var boxCapacity = await _context.BoxCapacity.FindAsync(id);
                if (boxCapacity == null)
                {
                    _loggerService.LogError($"BoxCapacity with ID {id} not found", "BoxCapacity",nameof(BoxCapacitiesController));
                    return NotFound();
                }
                _loggerService.LogEvent($"BoxCapacity with ID {id} has been deleted", "BoxCapacity", User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0, 0);

                _context.BoxCapacity.Remove(boxCapacity);
                await _context.SaveChangesAsync();

                return NoContent();
            }
            catch (Exception ex)
            {

                    _loggerService.LogError($"Error deleting BoxCapacity with Id {id}", ex.Message, nameof(BoxCapacitiesController));
                    return StatusCode(500, "Internal server error");
                
            }
        }

        private bool BoxCapacityExists(int id)
        {
            return _context.BoxCapacity.Any(e => e.BoxCapacityId == id);
        }
    }
}
