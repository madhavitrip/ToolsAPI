using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using ERPToolsAPI.Data;
using Tools.Models;
using System.Text.Json;

namespace Tools.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class EnvelopeBreakagesController : ControllerBase
    {
        private readonly ERPToolsDbContext _context;

        public EnvelopeBreakagesController(ERPToolsDbContext context)
        {
            _context = context;
        }

        // GET: api/EnvelopeBreakages
        [HttpGet]
        public async Task<ActionResult<IEnumerable<EnvelopeBreakage>>> GetEnvelopeBreakages(int ProjectId)
        {
            var NRData = await _context.NRDatas.Where(p=>p.ProjectId == ProjectId).ToListAsync();
            var Envelope = await _context.EnvelopeBreakages.Where(p => p.ProjectId == ProjectId).ToListAsync();

            var Consolidated = from nr in NRData
                               join env in Envelope
                               on nr.Id equals env.NrDataId
                               select new
                               {
                                NRDataId = nr.Id,
                                InnerEnv = env.InnerEnvelope,
                                OuterEnv = env.OuterEnvelope,
                                CourseName = nr.CourseName,
                                SubjectName = nr.SubjectName,
                                NrData = nr.NRDatas,
                                CatchNo = nr.CatchNo,
                                CenterCode = nr.CenterCode,
                                ExamTime = nr.ExamTime,
                                ExamDate = nr.ExamDate,
                                ProjectId = nr.ProjectId,
                                Quantity = nr.Quantity,

                               };
              return Ok(Consolidated);
        }

        // GET: api/EnvelopeBreakages/5
        [HttpGet("{id}")]
        public async Task<ActionResult<EnvelopeBreakage>> GetEnvelopeBreakage(int id)
        {
            var envelopeBreakage = await _context.EnvelopeBreakages.FindAsync(id);

            if (envelopeBreakage == null)
            {
                return NotFound();
            }

            return envelopeBreakage;
        }

        // PUT: api/EnvelopeBreakages/5
        // To protect from overposting attacks, see https://go.microsoft.com/fwlink/?linkid=2123754
        [HttpPut("{id}")]
        public async Task<IActionResult> PutEnvelopeBreakage(int id, EnvelopeBreakage envelopeBreakage)
        {
            if (id != envelopeBreakage.EnvelopeId)
            {
                return BadRequest();
            }

            _context.Entry(envelopeBreakage).State = EntityState.Modified;

            try
            {
                await _context.SaveChangesAsync();
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!EnvelopeBreakageExists(id))
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

        // POST: api/EnvelopeBreakages
        // To protect from overposting attacks, see https://go.microsoft.com/fwlink/?linkid=2123754
        [HttpPost("EnvelopeConfiguration")]
        public async Task<IActionResult> EnvelopeConfiguration(int ProjectId)
        {
            var envelopesJson = await _context.ProjectConfigs
                .Where(s => s.ProjectId == ProjectId)
                .Select(s => s.Envelope)
                .FirstOrDefaultAsync();

            if (envelopesJson == null)
                return NotFound("No envelope config found.");

            var nrDataList = await _context.NRDatas
                .Where(s => s.ProjectId == ProjectId)
                .ToListAsync();

            if (!nrDataList.Any())
                return NotFound("No NRData found.");

            var envelopeDict = JsonSerializer.Deserialize<Dictionary<string, string>>(envelopesJson);
            if (envelopeDict == null)
                return BadRequest("Invalid envelope JSON format.");

            var innerSizes = envelopeDict["Inner"]
                .Split(',', StringSplitOptions.RemoveEmptyEntries)
                .Select(e => e.Trim().ToUpper().Replace("E", ""))
                .Where(x => int.TryParse(x, out _))
                .Select(int.Parse)
                .OrderByDescending(x => x)
                .ToList();

            var outerSizes = envelopeDict["Outer"]
                .Split(',', StringSplitOptions.RemoveEmptyEntries)
                .Select(e => e.Trim().ToUpper().Replace("E", ""))
                .Where(x => int.TryParse(x, out _))
                .Select(int.Parse)
                .OrderByDescending(x => x)
                .ToList();

            foreach (var row in nrDataList)
            {
                int quantity = row.Quantity ?? 0;
                int remaining = quantity;

                Dictionary<string, string> innerBreakdown = new();
                foreach (var size in innerSizes)
                {
                    int count = remaining / size;
                    if (count > 0)
                    {
                        innerBreakdown[$"E{size}"] = count.ToString();
                        remaining -= count * size;
                    }
                }

                remaining = quantity; // reset for outer
                Dictionary<string, string> outerBreakdown = new();
                foreach (var size in outerSizes)
                {
                    int count = remaining / size;
                    if (count > 0)
                    {
                        outerBreakdown[$"E{size}"] = count.ToString();
                        remaining -= count * size;
                    }
                }

                var envelope = new EnvelopeBreakage
                {
                    ProjectId = ProjectId,
                    NrDataId = row.Id,
                    InnerEnvelope = JsonSerializer.Serialize(innerBreakdown),
                    OuterEnvelope = JsonSerializer.Serialize(outerBreakdown)
                };

                // ✅ Add to database
                _context.EnvelopeBreakages.Add(envelope);
            }

            await _context.SaveChangesAsync();

            return Ok("Envelope breakdown successfully saved to EnvelopeBreakage table.");
        }

        // DELETE: api/EnvelopeBreakages/5
        [HttpDelete("{id}")]
        public async Task<IActionResult> DeleteEnvelopeBreakage(int id)
        {
            var envelopeBreakage = await _context.EnvelopeBreakages.FindAsync(id);
            if (envelopeBreakage == null)
            {
                return NotFound();
            }

            _context.EnvelopeBreakages.Remove(envelopeBreakage);
            await _context.SaveChangesAsync();

            return NoContent();
        }

        private bool EnvelopeBreakageExists(int id)
        {
            return _context.EnvelopeBreakages.Any(e => e.EnvelopeId == id);
        }
    }
}
