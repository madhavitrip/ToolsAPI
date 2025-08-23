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
    public class ExtraTypesController : ControllerBase
    {
        private readonly ERPToolsDbContext _context;

        public ExtraTypesController(ERPToolsDbContext context)
        {
            _context = context;
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
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!ExtraTypeExists(id))
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

        // POST: api/ExtraTypes
        // To protect from overposting attacks, see https://go.microsoft.com/fwlink/?linkid=2123754
        [HttpPost]
        public async Task<ActionResult<ExtraType>> PostExtraType(ExtraType extraType)
        {
            _context.ExtraType.Add(extraType);
            await _context.SaveChangesAsync();

            return CreatedAtAction("GetExtraType", new { id = extraType.ExtraTypeId }, extraType);
        }

        // DELETE: api/ExtraTypes/5
        [HttpDelete("{id}")]
        public async Task<IActionResult> DeleteExtraType(int id)
        {
            var extraType = await _context.ExtraType.FindAsync(id);
            if (extraType == null)
            {
                return NotFound();
            }

            _context.ExtraType.Remove(extraType);
            await _context.SaveChangesAsync();

            return NoContent();
        }

        private bool ExtraTypeExists(int id)
        {
            return _context.ExtraType.Any(e => e.ExtraTypeId == id);
        }
    }
}
