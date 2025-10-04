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
    public class EnvelopeTypesController : ControllerBase
    {
        private readonly ERPToolsDbContext _context;
        private readonly ILoggerService _loggerService;

        public EnvelopeTypesController(ERPToolsDbContext context, ILoggerService loggerService)
        {
            _context = context;
            _loggerService = loggerService;
        }

        // GET: api/EnvelopeTypes
        [HttpGet]
        public async Task<ActionResult<IEnumerable<EnvelopeType>>> GetEnvelopesTypes()
        {
            return await _context.EnvelopesTypes.ToListAsync();
        }

        // GET: api/EnvelopeTypes/5
        [HttpGet("{id}")]
        public async Task<ActionResult<EnvelopeType>> GetEnvelopeType(int id)
        {
            var envelopeType = await _context.EnvelopesTypes.FindAsync(id);

            if (envelopeType == null)
            {
                return NotFound();
            }

            return envelopeType;
        }

        // PUT: api/EnvelopeTypes/5
        // To protect from overposting attacks, see https://go.microsoft.com/fwlink/?linkid=2123754
        [HttpPut("{id}")]
        public async Task<IActionResult> PutEnvelopeType(int id, EnvelopeType envelopeType)
        {
            if (id != envelopeType.EnvelopeId)
            {
                return BadRequest();
            }

            _context.Entry(envelopeType).State = EntityState.Modified;

            try
            {

                await _context.SaveChangesAsync();
                _loggerService.LogEvent($"Updated EnvelopeType with {id}", "EnvelopeType", User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0);

            }
            catch (Exception ex)
            {
                if (!EnvelopeTypeExists(id))
                {
                    _loggerService.LogEvent($"EnvelopeType with ID {id} not found during updating", "EnvelopeType", User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0);
                    return NotFound();
                }
                else
                {
                    _loggerService.LogError("Error updating EnvelopeType", ex.Message, nameof(EnvelopeTypesController));
                    return StatusCode(500, "Internal Server Error");
                }
            }

            return NoContent();
        }

        // POST: api/EnvelopeTypes
        // To protect from overposting attacks, see https://go.microsoft.com/fwlink/?linkid=2123754
        [HttpPost]
        public async Task<ActionResult<EnvelopeType>> PostEnvelopeType(EnvelopeType envelopeType)
        {
            try
            {
                _context.EnvelopesTypes.Add(envelopeType);
                await _context.SaveChangesAsync();
                _loggerService.LogEvent($"Created a new EnvelopeType with ID {envelopeType.EnvelopeId}", "EnvelopeType", User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0);
                return CreatedAtAction("GetEnvelopeType", new { id = envelopeType.EnvelopeId }, envelopeType);
            }
            catch (Exception ex)
            {
                _loggerService.LogError("Error creating EnvelopeType", ex.Message, nameof(EnvelopeTypesController));
                return StatusCode(500, "Internal Server Error");
            }
        }

        // DELETE: api/EnvelopeTypes/5
        [HttpDelete("{id}")]
        public async Task<IActionResult> DeleteEnvelopeType(int id)
        {
            try
            {
                var envelopeType = await _context.EnvelopesTypes.FindAsync(id);
                if (envelopeType == null)
                {
                    return NotFound();
                }

                _context.EnvelopesTypes.Remove(envelopeType);
                await _context.SaveChangesAsync();
                _loggerService.LogEvent($"Deleted a EnvelopeType with ID {id}", "EnvelopeType", User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0);
                return NoContent();
            }
            catch (Exception ex)
            {
                _loggerService.LogError($"Error deleting EnvelopeType with Id{id}", ex.Message, nameof(EnvelopeTypesController));
                return StatusCode(500, "Internal Server Error");
            }
        }

        private bool EnvelopeTypeExists(int id)
        {
            return _context.EnvelopesTypes.Any(e => e.EnvelopeId == id);
        }
    }
}
