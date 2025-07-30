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
    public class BoxCapacitiesController : ControllerBase
    {
        private readonly ERPToolsDbContext _context;

        public BoxCapacitiesController(ERPToolsDbContext context)
        {
            _context = context;
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
                await _context.SaveChangesAsync();
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!BoxCapacityExists(id))
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

        // POST: api/BoxCapacities
        // To protect from overposting attacks, see https://go.microsoft.com/fwlink/?linkid=2123754
        [HttpPost]
        public async Task<ActionResult<BoxCapacity>> PostBoxCapacity(BoxCapacity boxCapacity)
        {
            _context.BoxCapacity.Add(boxCapacity);
            await _context.SaveChangesAsync();

            return CreatedAtAction("GetBoxCapacity", new { id = boxCapacity.BoxCapacityId }, boxCapacity);
        }

        // DELETE: api/BoxCapacities/5
        [HttpDelete("{id}")]
        public async Task<IActionResult> DeleteBoxCapacity(int id)
        {
            var boxCapacity = await _context.BoxCapacity.FindAsync(id);
            if (boxCapacity == null)
            {
                return NotFound();
            }

            _context.BoxCapacity.Remove(boxCapacity);
            await _context.SaveChangesAsync();

            return NoContent();
        }

        private bool BoxCapacityExists(int id)
        {
            return _context.BoxCapacity.Any(e => e.BoxCapacityId == id);
        }
    }
}
