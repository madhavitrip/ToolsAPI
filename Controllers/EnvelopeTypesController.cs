using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using ERPToolsAPI.Data;
using Tools.Models;

namespace Tools.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class EnvelopeTypesController : ControllerBase
    {
        private readonly ERPToolsDbContext _context;

        public EnvelopeTypesController(ERPToolsDbContext context)
        {
            _context = context;
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
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!EnvelopeTypeExists(id))
                {
                    return NotFound();
                }
                else
                {
                    throw;
                }
            }

            return NoContent();
        }

        // POST: api/EnvelopeTypes
        // To protect from overposting attacks, see https://go.microsoft.com/fwlink/?linkid=2123754
        [HttpPost]
        public async Task<ActionResult<EnvelopeType>> PostEnvelopeType(EnvelopeType envelopeType)
        {
            _context.EnvelopesTypes.Add(envelopeType);
            await _context.SaveChangesAsync();

            return CreatedAtAction("GetEnvelopeType", new { id = envelopeType.EnvelopeId }, envelopeType);
        }

        // DELETE: api/EnvelopeTypes/5
        [HttpDelete("{id}")]
        public async Task<IActionResult> DeleteEnvelopeType(int id)
        {
            var envelopeType = await _context.EnvelopesTypes.FindAsync(id);
            if (envelopeType == null)
            {
                return NotFound();
            }

            _context.EnvelopesTypes.Remove(envelopeType);
            await _context.SaveChangesAsync();

            return NoContent();
        }

        private bool EnvelopeTypeExists(int id)
        {
            return _context.EnvelopesTypes.Any(e => e.EnvelopeId == id);
        }
    }
}
