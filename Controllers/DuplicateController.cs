using ERPToolsAPI.Data;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Reflection;

namespace Tools.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class DuplicateController : ControllerBase
    {
        private readonly ERPToolsDbContext _context;

        public DuplicateController(ERPToolsDbContext context)
        {
            _context = context;
        }

        [HttpPost]
        public async Task<IActionResult> MergeFields(int ProjectId, string mergefields, bool consolidate)
        {
            // Split the merge fields (e.g., "center_code,catch_number")
            var mergeFields = mergefields.Split(',', StringSplitOptions.RemoveEmptyEntries)
                                          .Select(f => f.Trim().ToLower()).ToList();

            if (mergeFields.Count == 0)
                return BadRequest("No merge fields provided.");

            // Get the data for the project
            var data = await _context.NRDatas
                .Where(p => p.ProjectId == ProjectId)
                .ToListAsync();

            if (!data.Any())
                return NotFound("No data found for this project.");

            // Group the data based on the merge fields
            var grouped = data.GroupBy(d =>
            {
                var key = new List<string>();
                foreach (var field in mergeFields)
                {
                    var value = d.GetType().GetProperty(field, BindingFlags.IgnoreCase | BindingFlags.Public | BindingFlags.Instance)
                                 ?.GetValue(d)?.ToString()?.Trim() ?? "";
                    key.Add(value);
                }
                return string.Join("|", key); // Composite key
            });

            int mergedCount = 0;

            foreach (var group in grouped)
            {
                if (group.Count() <= 1)
                    continue;

                // Consolidate: Sum quantity, keep first row
                var keep = group.First();
                if (consolidate)
                {
                    keep.Quantity = group.Sum(x => x.Quantity);
                }

                // Remove duplicates (excluding the one we keep)
                var duplicates = group.Skip(1).ToList();
                _context.NRDatas.RemoveRange(duplicates);

                mergedCount += duplicates.Count;
            }

            await _context.SaveChangesAsync();

            return Ok(new
            {
                MergedRows = mergedCount,
                Consolidated = consolidate
            });
        }

    }
}
