using System;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;
using ERPToolsAPI.Data;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Tools.Services;

namespace ToolsAPI.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class MasterAuthController : ControllerBase
    {
        private readonly IMasterAuthService _masterAuthService;
        private readonly ERPToolsDbContext _context;

        public MasterAuthController(IMasterAuthService masterAuthService, ERPToolsDbContext context)
        {
            _masterAuthService = masterAuthService;
            _context = context;
        }

        public class VerifyPasscodeRequest
        {
            public int GroupId { get; set; }
            public string Passcode { get; set; } = string.Empty;
            public string ModuleName { get; set; } = "Master";
            public string OperationType { get; set; } = "VERIFY";
        }

        public class ChangePasscodeRequest
        {
            public int GroupId { get; set; }
            public string CurrentPasscode { get; set; } = string.Empty;
            public string NewPasscode { get; set; } = string.Empty;
        }

        /// <summary>
        /// Check if a Master Authorization Passcode has been initialized/set for a specific Group
        /// </summary>
        [HttpGet("status")]
        public IActionResult GetStatus([FromQuery] int groupId = 0, [FromQuery] int projectId = 0)
        {
            if (groupId <= 0 && projectId > 0)
            {
                groupId = _masterAuthService.ResolveGroupIdFromProjectId(projectId);
            }

            bool isSet = _masterAuthService.IsPasscodeSetForGroup(groupId);
            return Ok(new { isPasscodeSet = isSet, groupId });
        }

        /// <summary>
        /// Reset active brute force lockout for a Group
        /// </summary>
        [HttpPost("reset-lockout")]
        public IActionResult ResetLockout([FromQuery] int groupId = 0)
        {
            int userId = GetCurrentUserId();
            string ipAddress = GetClientIpAddress();
            _masterAuthService.ResetLockout(groupId, userId, ipAddress);
            _masterAuthService.ClearAllLockouts();
            return Ok(new { success = true, message = "Lockout cleared successfully. You may now set or enter your passcode." });
        }

        /// <summary>
        /// Explicitly verify Master Authorization Passcode for a Group
        /// </summary>
        [HttpPost("verify")]
        public async Task<IActionResult> VerifyPasscode([FromBody] VerifyPasscodeRequest request)
        {
            int userId = GetCurrentUserId();
            string ipAddress = GetClientIpAddress();

            var (isValid, errorMessage) = await _masterAuthService.VerifyPasscodeAsync(
                request.Passcode,
                request.GroupId,
                userId,
                request.OperationType,
                request.ModuleName,
                ipAddress
            );

            if (!isValid)
            {
                return StatusCode(403, new { success = false, message = errorMessage, groupId = request.GroupId });
            }

            return Ok(new { success = true, message = "Master authorization passcode verified successfully.", groupId = request.GroupId });
        }

        /// <summary>
        /// Fetch audit logs for master authorization attempts
        /// </summary>
        [HttpGet("audit-logs")]
        public async Task<IActionResult> GetAuditLogs([FromQuery] int limit = 50, [FromQuery] string module = null, [FromQuery] int groupId = 0)
        {
            var query = _context.MasterAuthAuditLogs.AsNoTracking().AsQueryable();

            if (groupId > 0)
            {
                query = query.Where(l => l.GroupId == groupId);
            }

            if (!string.IsNullOrWhiteSpace(module))
            {
                query = query.Where(l => l.ModuleName == module);
            }

            var logs = await query
                .OrderByDescending(l => l.Timestamp)
                .Take(limit > 0 ? limit : 50)
                .ToListAsync();

            return Ok(logs);
        }

        /// <summary>
        /// Create or Change Master Authorization Passcode for a Group
        /// </summary>
        [HttpPost("change-passcode")]
        public async Task<IActionResult> ChangePasscode([FromBody] ChangePasscodeRequest request)
        {
            int userId = GetCurrentUserId();
            string ipAddress = GetClientIpAddress();

            var (success, error) = await _masterAuthService.UpdatePasscodeAsync(
                request.CurrentPasscode,
                request.NewPasscode,
                request.GroupId,
                userId,
                ipAddress
            );

            if (success)
            {
                return Ok(new { success = true, message = $"Master authorization passcode updated successfully for Group {request.GroupId}.", groupId = request.GroupId });
            }

            return BadRequest(new { success = false, message = error, groupId = request.GroupId });
        }

        private int GetCurrentUserId()
        {
            var userIdClaim = User?.Claims.FirstOrDefault(c => c.Type == "userid" || c.Type == ClaimTypes.NameIdentifier);
            if (userIdClaim != null && int.TryParse(userIdClaim.Value, out var userId))
            {
                return userId;
            }
            if (Request.Headers.TryGetValue("X-User-Id", out var headerVal) && int.TryParse(headerVal.ToString(), out var headerUserId))
            {
                return headerUserId;
            }
            return 0;
        }

        private string GetClientIpAddress()
        {
            string ip = HttpContext.Connection.RemoteIpAddress?.ToString() ?? "0.0.0.0";
            if (Request.Headers.TryGetValue("X-Forwarded-For", out var forwardedFor))
            {
                ip = forwardedFor.ToString().Split(',')[0].Trim();
            }
            return ip;
        }
    }
}
