using ERPToolsAPI.Data;
using ERPToolsAPI.Models;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.IdentityModel.Tokens.Jwt;
using System.Security.Claims;
using System.Text;

namespace ERPToolsAPI.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class UserLogsController : ControllerBase
    {
        private readonly ERPToolsDbContext _context;
        private readonly string _baseUrl;

        public UserLogsController(ERPToolsDbContext context,IConfiguration configuration )
        {
            _context = context;
            _baseUrl = configuration["ApiSettings:BaseUrl"];
        }


        [HttpPost("login")]
        public async Task<IActionResult> LoginFromTools([FromBody] LoginRequest request)
        {
            try
            {
                using var httpClient = new HttpClient();

                // Call ERP API for authentication
                var erpLoginPayload = new
                {
                    userName = request.UserName,
                    password = request.Password
                };

                var content = new StringContent(JsonConvert.SerializeObject(erpLoginPayload), Encoding.UTF8, "application/json");

                var erpAuthResponse = await httpClient.PostAsync($"{_baseUrl}/Login/login", content);

                if (!erpAuthResponse.IsSuccessStatusCode)
                {
                    return Unauthorized(new { message = "Invalid credentials" });
                }

                var responseContent = await erpAuthResponse.Content.ReadAsStringAsync();
                var jwtToken = JObject.Parse(responseContent)["token"]?.ToString(); // adjust based on ERP's actual response

                if (string.IsNullOrEmpty(jwtToken))
                    return StatusCode(500, "ERP didn't return token");

                // ✅ Decode token to get UserId
                var handler = new JwtSecurityTokenHandler();
                var token = handler.ReadJwtToken(jwtToken);
                var userIdClaim = token.Claims.FirstOrDefault(c => c.Type == ClaimTypes.Name); // or "sub" or custom
                if (userIdClaim == null)
                    return StatusCode(500, "UserId not found in token");

                int userId = int.Parse(userIdClaim.Value);

                // 📝 Log to ERPTools DB
                var log = new UserLoginLog
                {
                    UserId = userId,
                    LoginTime = DateTime.UtcNow,
                   
                };

                _context.UserLoginLogs.Add(log);
                await _context.SaveChangesAsync();

                // Return token to frontend
                return Ok(new { token = jwtToken });

            }
            catch (Exception ex)
            {
                return StatusCode(500, new { message = "Login failed", error = ex.Message });
            }
        }


        [Authorize]
        [HttpPost("log-login")]
        public IActionResult LogLogin(UserLoginLog  userLoginLog)
        {
            var userIdClaim = User.FindFirst(ClaimTypes.Name);
            if (userIdClaim == null) return Unauthorized();

            var userId = int.Parse(userIdClaim.Value);
            var ipAddress = HttpContext.Connection.RemoteIpAddress?.ToString();
            var userAgent = Request.Headers["User-Agent"].ToString();

            var log = new UserLoginLog
            {
                UserId = userId,
                LoginTime = DateTime.UtcNow
            };

            _context.UserLoginLogs.Add(log);
            _context.SaveChanges();

            return Ok(new { message = "Login logged successfully" });
        }
    }
}
