using System.IdentityModel.Tokens.Jwt;
using System.Security.Claims;
using System.Text.Json;
using Microsoft.AspNetCore.Http;

namespace Tools.Services
{
    public static class LogHelper
    {
        public static int GetTriggeredBy(ClaimsPrincipal user, HttpRequest request = null)
        {
            // Try claims first (populated when [Authorize] is present)
            if (user != null)
            {
                var value = user.FindFirst("userid")?.Value;
                if (int.TryParse(value, out var id) && id > 0)
                    return id;
            }

            // Fallback: parse JWT directly from Authorization header
            // (for controllers without [Authorize] where claims aren't populated)
            if (request != null)
            {
                var token = request.Headers["Authorization"].ToString()?.Replace("Bearer ", "").Trim();
                if (!string.IsNullOrWhiteSpace(token))
                {
                    try
                    {
                        var handler = new JwtSecurityTokenHandler();
                        if (handler.CanReadToken(token))
                        {
                            var jwt = handler.ReadToken(token) as JwtSecurityToken;
                            var claim = jwt?.Claims.FirstOrDefault(c => c.Type == "userid")?.Value;
                            if (int.TryParse(claim, out var jwtId) && jwtId > 0)
                                return jwtId;
                        }
                    }
                    catch { }
                }
            }

            return 0;
        }

        public static string ToJson(object value)
        {
            return value == null ? string.Empty : JsonSerializer.Serialize(value);
        }
    }
}
