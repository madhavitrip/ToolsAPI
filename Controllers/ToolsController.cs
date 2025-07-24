/*using ERPToolsAPI.Data;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using System.Security.Claims;
using ERPToolsAPI.Models;

namespace ERPToolsAPI.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    [Authorize]
    public class ToolsController : ControllerBase
    {
        private readonly ERPToolsDbContext _context;

        public ToolsController(ERPToolsDbContext context)
        {
            _context = context;
        }

        [HttpGet]
        public IActionResult GetUserTools()
        {
            int userId = int.Parse(User.FindFirst(ClaimTypes.Name)?.Value ?? "0");
            var tools = _context.ToolRecords.Where(t => t.UserId == userId).ToList();
            return Ok(tools);
        }

        [HttpPost]
        public IActionResult AddTool([FromBody] ToolRecord tool)
        {
            int userId = int.Parse(User.FindFirst(ClaimTypes.Name)?.Value ?? "0");
            tool.UserId = userId;
            tool.CreatedAt = DateTime.UtcNow;

            _context.ToolRecords.Add(tool);
            _context.SaveChanges();

            return Ok(new { message = "Tool record added", tool });
        }
    }
}
*/