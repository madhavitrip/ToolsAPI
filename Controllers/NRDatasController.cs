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
    public class NRDatasController : ControllerBase
    {
        private readonly ERPToolsDbContext _context;

        public NRDatasController(ERPToolsDbContext context)
        {
            _context = context;
        }

        // GET: api/NRDatas
        [HttpGet]
        public async Task<ActionResult<IEnumerable<NRData>>> GetNRDatas()
        {
            return await _context.NRDatas.ToListAsync();
        }

        // GET: api/NRDatas/5
        [HttpGet("{id}")]
        public async Task<ActionResult<NRData>> GetNRData(int id)
        {
            var nRData = await _context.NRDatas.FindAsync(id);

            if (nRData == null)
            {
                return NotFound();
            }

            return nRData;
        }

        // PUT: api/NRDatas/5
        // To protect from overposting attacks, see https://go.microsoft.com/fwlink/?linkid=2123754
        [HttpPut("{id}")]
        public async Task<IActionResult> PutNRData(int id, NRData nRData)
        {
            if (id != nRData.Id)
            {
                return BadRequest();
            }

            _context.Entry(nRData).State = EntityState.Modified;

            try
            {
                await _context.SaveChangesAsync();
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!NRDataExists(id))
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

        // POST: api/NRDatas
        // To protect from overposting attacks, see https://go.microsoft.com/fwlink/?linkid=2123754
        [HttpPost]
        public async Task<ActionResult<NRData>> PostNRData(NRData nRData)
        {
            _context.NRDatas.Add(nRData);
            await _context.SaveChangesAsync();

            return CreatedAtAction("GetNRData", new { id = nRData.Id }, nRData);
        }

        // DELETE: api/NRDatas/5
        [HttpDelete("{id}")]
        public async Task<IActionResult> DeleteNRData(int id)
        {
            var nRData = await _context.NRDatas.FindAsync(id);
            if (nRData == null)
            {
                return NotFound();
            }

            _context.NRDatas.Remove(nRData);
            await _context.SaveChangesAsync();

            return NoContent();
        }

        private bool NRDataExists(int id)
        {
            return _context.NRDatas.Any(e => e.Id == id);
        }
    }
}
