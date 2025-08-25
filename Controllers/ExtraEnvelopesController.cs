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
    public class ExtraEnvelopesController : ControllerBase
    {
        private readonly ERPToolsDbContext _context;

        public ExtraEnvelopesController(ERPToolsDbContext context)
        {
            _context = context;
        }

        // GET: api/ExtraEnvelopes
        [HttpGet]
        public async Task<ActionResult<IEnumerable<ExtraEnvelopes>>> GetExtrasEnvelope(int ProjectId)
        {
            var NrData = await _context.NRDatas.Where(p=>p.ProjectId == ProjectId).ToListAsync();

            return await _context.ExtrasEnvelope.ToListAsync();
        }

        // GET: api/ExtraEnvelopes/5
        [HttpGet("{id}")]
        public async Task<ActionResult<ExtraEnvelopes>> GetExtraEnvelopes(int id)
        {
            var extraEnvelopes = await _context.ExtrasEnvelope.FindAsync(id);

            if (extraEnvelopes == null)
            {
                return NotFound();
            }

            return extraEnvelopes;
        }

        // PUT: api/ExtraEnvelopes/5
        // To protect from overposting attacks, see https://go.microsoft.com/fwlink/?linkid=2123754
        [HttpPut("{id}")]
        public async Task<IActionResult> PutExtraEnvelopes(int id, ExtraEnvelopes extraEnvelopes)
        {
            if (id != extraEnvelopes.Id)
            {
                return BadRequest();
            }

            _context.Entry(extraEnvelopes).State = EntityState.Modified;

            try
            {
                await _context.SaveChangesAsync();
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!ExtraEnvelopesExists(id))
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

        public class EnvelopeType
        {
            public string Inner { get; set; }
            public string Outer { get; set; }
        }


        // POST: api/ExtraEnvelopes
        // To protect from overposting attacks, see https://go.microsoft.com/fwlink/?linkid=2123754
        [HttpPost]
        public async Task<ActionResult> PostExtraEnvelopes(int ProjectId)
        {
            var nrDataList = await _context.NRDatas
                .Where(d => d.ProjectId == ProjectId)
                .ToListAsync();

            var extraConfig = await _context.ExtraConfigurations
                .Where(c => c.ProjectId == ProjectId).ToListAsync();

            if (extraConfig == null || !extraConfig.Any())
                return BadRequest("No ExtraConfiguration found for the project.");

            var envelopesToAdd = new List<ExtraEnvelopes>();
            foreach (var config in extraConfig)
            {
                EnvelopeType envelopeType;
                try
                {
                    envelopeType = JsonSerializer.Deserialize<EnvelopeType>(config.EnvelopeType);
                }
                catch
                {
                    return BadRequest($"Invalid EnvelopeType JSON for ExtraType {config.ExtraType}");
                }


                int innerCapacity = GetEnvelopeCapacity(envelopeType.Inner);
                int outerCapacity = GetEnvelopeCapacity(envelopeType.Outer);


                foreach (var data in nrDataList)
                {
                    int calculatedQuantity = 0;

                    switch (config.Mode)
                    {
                        case "Fixed":
                            calculatedQuantity = int.Parse(config.Value);
                            break;

                        case "Percentage":
                            if (decimal.TryParse(config.Value, out var percentValue))
                            {
                                calculatedQuantity = (int)Math.Ceiling((double)(data.Quantity * percentValue) / 100);
                            }
                            break;

                        case "Range":
                            // calculatedQuantity = HandleRangeLogic(data, extraConfig);
                            break;
                    }

                    int innerCount = (int)Math.Ceiling((double)calculatedQuantity / innerCapacity);
                    int outerCount = (int)Math.Ceiling((double)calculatedQuantity / outerCapacity);

                    var envelope = new ExtraEnvelopes
                    {
                        ProjectId =ProjectId,
                        NRDataId = data.Id,
                        ExtraId = config.ExtraType,
                        Quantity = calculatedQuantity,
                        InnerEnvelope = innerCount.ToString(),
                        OuterEnvelope = outerCount.ToString(),

                    };

                    envelopesToAdd.Add(envelope);

                    // Log
                    Console.WriteLine($"NRDataId {data.Id} → Quantity: {calculatedQuantity}, Inner: {innerCount} packets, Outer: {outerCount} packets");
                }
            }
            await _context.ExtrasEnvelope.AddRangeAsync(envelopesToAdd);
            await _context.SaveChangesAsync();

            return Ok(envelopesToAdd);
        }

        private int GetEnvelopeCapacity(string envelopeCode)
        {
            if (string.IsNullOrWhiteSpace(envelopeCode))
                return 1; // default to 1 if null or invalid

            // Expecting format like "E10", "E25", etc.
            var numberPart = new string(envelopeCode.Where(char.IsDigit).ToArray());

            return int.TryParse(numberPart, out var capacity) ? capacity : 1;
        }


        // DELETE: api/ExtraEnvelopes/5
        [HttpDelete("{id}")]
        public async Task<IActionResult> DeleteExtraEnvelopes(int id)
        {
            var extraEnvelopes = await _context.ExtrasEnvelope.FindAsync(id);
            if (extraEnvelopes == null)
            {
                return NotFound();
            }

            _context.ExtrasEnvelope.Remove(extraEnvelopes);
            await _context.SaveChangesAsync();

            return NoContent();
        }

        private bool ExtraEnvelopesExists(int id)
        {
            return _context.ExtrasEnvelope.Any(e => e.Id == id);
        }
    }
}
