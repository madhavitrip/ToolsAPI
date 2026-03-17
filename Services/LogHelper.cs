using System.Security.Claims;
using System.Text.Json;

namespace Tools.Services
{
    public static class LogHelper
    {
        public static int GetTriggeredBy(ClaimsPrincipal user)
        {
            if (user == null)
            {
                return 0;
            }

            var value = user.FindFirst(ClaimTypes.Name)?.Value;
            return int.TryParse(value, out var id) ? id : 0;
        }

        public static string ToJson(object value)
        {
            return value == null ? string.Empty : JsonSerializer.Serialize(value);
        }
    }
}
