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
using static Tools.Controllers.ExtraEnvelopesController;

namespace Tools.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class FieldsController : ControllerBase
    {
        private readonly ERPToolsDbContext _context;
        private readonly ILoggerService _loggerService;
        public FieldsController(ERPToolsDbContext context, ILoggerService loggerService)
        {
            _context = context;
            _loggerService = loggerService;
        }

        // GET: api/Fields
        [HttpGet]
        public async Task<ActionResult<IEnumerable<Field>>> GetFields()
        {
            return await _context.Fields.ToListAsync();
        }

        // GET: api/Fields/5
        [HttpGet("{id}")]
        public async Task<ActionResult<Field>> GetField(int id)
        {
            var @field = await _context.Fields.FindAsync(id);

            if (@field == null)
            {
                return NotFound();
            }

            return @field;
        }

        // PUT: api/Fields/5
        // To protect from overposting attacks, see https://go.microsoft.com/fwlink/?linkid=2123754
        [HttpPut("{id}")]
        public async Task<IActionResult> PutField(int id, Field @field)
        {
            if (id != @field.FieldId)
            {
                return BadRequest();
            }

            _context.Entry(@field).State = EntityState.Modified;

            try
            {
                await _context.SaveChangesAsync();
                _loggerService.LogEvent($"Updated Field for Id {id}", "Field", User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0,0);

            }
            catch (Exception ex)
            {
                if (!FieldExists(id))
                {
                    _loggerService.LogEvent($"Field with ID {id} not found during updating", "Field", User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0,0);

                    return NotFound();
                }
                else
                {

                    _loggerService.LogError("Error updating Field", ex.Message, nameof(FieldsController));
                    return StatusCode(500, "Internal server error");
                }
            }

            return NoContent();
        }

        // POST: api/Fields
        // To protect from overposting attacks, see https://go.microsoft.com/fwlink/?linkid=2123754
        [HttpPost]
        public async Task<ActionResult<Field>> PostField(Field @field)
        {
            try
            {
                _context.Fields.Add(@field);
                await _context.SaveChangesAsync();
                _loggerService.LogEvent($"Created a new Field with ID {field.FieldId}", "Field", User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0,0);

                return CreatedAtAction("GetField", new { id = @field.FieldId }, @field);
            }
            catch (Exception ex)
            {
                _loggerService.LogError("Error creating Field", ex.Message, nameof(FieldsController));
                return StatusCode(500, "Internal server error");
            }
        }

        // DELETE: api/Fields/5
        [HttpDelete("{id}")]
        public async Task<IActionResult> DeleteField(int id)
        {
            try
            {
                var @field = await _context.Fields.FindAsync(id);
                if (@field == null)
                {
                    return NotFound();
                }

                _context.Fields.Remove(@field);
                await _context.SaveChangesAsync();
                _loggerService.LogEvent($"Deleted a Field with ID {id}", "Field", User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0,0);

                return NoContent();
            }
            catch (Exception ex)
            {
                _loggerService.LogError("Error deleting Field", ex.Message, nameof(FieldsController));
                return StatusCode(500, "Internal server error");
            }
        }

        private bool FieldExists(int id)
        {
            return _context.Fields.Any(e => e.FieldId == id);
        }
    }
}
