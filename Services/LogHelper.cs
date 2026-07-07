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

        /// <summary>
        /// Extracts the roleId from JWT token claims.
        /// Returns 0 if roleId is not found or invalid.
        /// ALWAYS parses the JWT directly from Authorization header for reliability.
        /// </summary>
        public static int GetUserRoleId(ClaimsPrincipal user, HttpRequest request = null)
        {
            // Primary method: Parse JWT directly from Authorization header (most reliable)
            if (request != null)
            {
                var authHeader = request.Headers["Authorization"].ToString();
                if (!string.IsNullOrWhiteSpace(authHeader))
                {
                    var token = authHeader.Replace("Bearer ", "").Trim();
                    if (!string.IsNullOrWhiteSpace(token))
                    {
                        try
                        {
                            var handler = new JwtSecurityTokenHandler();
                            if (handler.CanReadToken(token))
                            {
                                var jwt = handler.ReadToken(token) as JwtSecurityToken;
                                
                                if (jwt != null)
                                {
                                    // Debug: Log all claims
                                    System.Diagnostics.Debug.WriteLine("[JWT CLAIMS DEBUG]");
                                    foreach (var claim in jwt.Claims)
                                    {
                                        System.Diagnostics.Debug.WriteLine($"  Claim: Type='{claim.Type}', Value='{claim.Value}'");
                                    }
                                    
                                    // Try roleId claim (custom claim from external auth system) - exact match
                                    var roleIdClaim = jwt.Claims.FirstOrDefault(c => c.Type == "roleId")?.Value;
                                    if (!string.IsNullOrWhiteSpace(roleIdClaim) && int.TryParse(roleIdClaim, out var roleId) && roleId > 0)
                                    {
                                        System.Diagnostics.Debug.WriteLine($"[JWT] Found roleId: {roleId}");
                                        return roleId;
                                    }
                                    
                                    // Try "roleid" (lowercase)
                                    var roleIdClaimLower = jwt.Claims.FirstOrDefault(c => c.Type == "roleid")?.Value;
                                    if (!string.IsNullOrWhiteSpace(roleIdClaimLower) && int.TryParse(roleIdClaimLower, out var roleIdFromLower) && roleIdFromLower > 0)
                                    {
                                        System.Diagnostics.Debug.WriteLine($"[JWT] Found roleid (lowercase): {roleIdFromLower}");
                                        return roleIdFromLower;
                                    }
                                    
                                    // Try "role" claim as fallback (standard ASP.NET role)
                                    var roleClaim = jwt.Claims.FirstOrDefault(c => c.Type == System.Security.Claims.ClaimTypes.Role)?.Value;
                                    if (!string.IsNullOrWhiteSpace(roleClaim) && int.TryParse(roleClaim, out var roleIdFromRole) && roleIdFromRole > 0)
                                    {
                                        System.Diagnostics.Debug.WriteLine($"[JWT] Found role claim: {roleIdFromRole}");
                                        return roleIdFromRole;
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"[JWT ERROR] {ex.Message}");
                        }
                    }
                }
            }

            // Fallback: Try User claims (in case [Authorize] populated them)
            if (user != null)
            {
                System.Diagnostics.Debug.WriteLine("[USER CLAIMS DEBUG]");
                foreach (var claim in user.Claims)
                {
                    System.Diagnostics.Debug.WriteLine($"  Claim: Type='{claim.Type}', Value='{claim.Value}'");
                }
                
                // Try "roleId" claim first (custom claim)
                var roleIdClaim = user.FindFirst("roleId")?.Value;
                if (!string.IsNullOrWhiteSpace(roleIdClaim) && int.TryParse(roleIdClaim, out var roleId) && roleId > 0)
                {
                    System.Diagnostics.Debug.WriteLine($"[USER] Found roleId: {roleId}");
                    return roleId;
                }
                
                // Try ClaimTypes.Role as fallback (standard ASP.NET role)
                var roleClaim = user.FindFirst(System.Security.Claims.ClaimTypes.Role)?.Value;
                if (!string.IsNullOrWhiteSpace(roleClaim) && int.TryParse(roleClaim, out var roleIdFromRole) && roleIdFromRole > 0)
                {
                    System.Diagnostics.Debug.WriteLine($"[USER] Found role: {roleIdFromRole}");
                    return roleIdFromRole;
                }
            }

            System.Diagnostics.Debug.WriteLine("[JWT] No roleId found - returning 0");
            return 0;
        }

        public static string ToJson(object value)
        {
            return value == null ? string.Empty : JsonSerializer.Serialize(value);
        }
    }
}

