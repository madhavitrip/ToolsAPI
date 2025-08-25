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
    public class ExtrasConfigurationsController : ControllerBase
    {
        private readonly ERPToolsDbContext _context;

        public ExtrasConfigurationsController(ERPToolsDbContext context)
        {
            _context = context;
        }

        // GET: api/ExtrasConfigurations
        [HttpGet]
        public async Task<ActionResult<IEnumerable<ExtrasConfiguration>>> GetExtraConfigurations()
        {
            return await _context.ExtraConfigurations.ToListAsync();
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
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!ExtrasConfigurationExists(id))
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

        // POST: api/ExtrasConfigurations
        // To protect from overposting attacks, see https://go.microsoft.com/fwlink/?linkid=2123754
        [HttpPost]
        public async Task<ActionResult<ExtrasConfiguration>> PostExtrasConfiguration(ExtrasConfiguration extrasConfiguration)
        {
            var extra = await _context.ExtraConfigurations.FindAsync(extrasConfiguration.ProjectId);
            if (extra!= null)
            {
                return Conflict(new { message = "A configuration already exists for this project." });
            }
            _context.ExtraConfigurations.Add(extrasConfiguration);
            await _context.SaveChangesAsync();

            return CreatedAtAction("GetExtrasConfiguration", new { id = extrasConfiguration.Id }, extrasConfiguration);
        }

        // DELETE: api/ExtrasConfigurations/5
        [HttpDelete("{id}")]
        public async Task<IActionResult> DeleteExtrasConfiguration(int id)
        {
            var extrasConfiguration = await _context.ExtraConfigurations.FindAsync(id);
            if (extrasConfiguration == null)
            {
                return NotFound();
            }

            _context.ExtraConfigurations.Remove(extrasConfiguration);
            await _context.SaveChangesAsync();

            return NoContent();
        }

        private bool ExtrasConfigurationExists(int id)
        {
            return _context.ExtraConfigurations.Any(e => e.Id == id);
        }
    }
}
