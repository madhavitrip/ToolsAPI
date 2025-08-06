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
        public async Task<ActionResult> PostNRData([FromBody] JsonElement inputData)
        {
            int projectId = inputData.GetProperty("projectId").GetInt32();
            var dataArray = inputData.GetProperty("data").EnumerateArray();

            foreach (var item in dataArray)
            {
                var nRData = new NRData();
                nRData.ProjectId = projectId;

                var extraData = new Dictionary<string, string>();
                var nRDataType = typeof(NRData);

                foreach (var prop in item.EnumerateObject())
                {
                    string key = prop.Name;
                    string value = prop.Value.GetString();

                    var propInfo = nRDataType.GetProperty(key.Replace(" ", ""),
                                        System.Reflection.BindingFlags.IgnoreCase |
                                        System.Reflection.BindingFlags.Public |
                                        System.Reflection.BindingFlags.Instance);

                    if (propInfo != null)
                    {
                        object convertedValue = Convert.ChangeType(value, propInfo.PropertyType);
                        propInfo.SetValue(nRData, convertedValue);
                    }
                    else
                    {
                        extraData[key] = value;
                    }
                }

                if (extraData.Count > 0)
                    nRData.NRDatas = System.Text.Json.JsonSerializer.Serialize(extraData);

                _context.NRDatas.Add(nRData);
            }

            await _context.SaveChangesAsync();
            return Ok("Data inserted successfully");
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
