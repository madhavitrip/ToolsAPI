using Microsoft.EntityFrameworkCore;
using Tools.Models;
using ERPToolsAPI.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;

namespace Tools.Services
{
    /// <summary>
    /// Centralized authorization service for role-based and project-level access control.
    /// Reusable across all modules (Master Configuration, RPT Templates, Lots, Catches, etc.)
    /// Roles are extracted from JWT token claims.
    /// </summary>
    public interface IAuthorizationService
    {
        /// <summary>
        /// Check if user can edit master configuration for a specific project.
        /// Roles must be extracted from JWT token claims.
        /// </summary>
        Task<bool> CanEditMasterConfigurationAsync(int userId, int projectId, ClaimsPrincipal user);

        /// <summary>
        /// Check if user can edit RPT master for a specific project.
        /// Roles must be extracted from JWT token claims.
        /// </summary>
        Task<bool> CanEditRPTMasterAsync(int userId, int projectId, ClaimsPrincipal user);

        /// <summary>
        /// Generic authorization check for any module/feature.
        /// Roles must be extracted from JWT token claims.
        /// </summary>
        Task<bool> IsAuthorizedForProjectAsync(int userId, int projectId, ClaimsPrincipal user);

        /// <summary>
        /// Get user's assigned groups for a project
        /// </summary>
        Task<List<int>> GetUserGroupsForProjectAsync(int userId, int projectId);

        /// <summary>
        /// Extract role IDs from JWT token claims
        /// </summary>
        List<int> ExtractRolesFromToken(ClaimsPrincipal user);

        /// <summary>
        /// Check if user has a full-access role (Developer/Admin/Head)
        /// </summary>
        bool HasFullAccessRole(List<int> userRoleIds);

        /// <summary>
        /// Check if role is in the authorized roles list
        /// </summary>
        bool IsAuthorizedRole(List<int> userRoleIds);
    }

    public class AuthorizationService : IAuthorizationService
    {
        private readonly ERPToolsDbContext _context;

        // Authorized roles: Developer (1), Admin (2), Head (3), Manager (4)
        private readonly List<int> _authorizedRoleIds = new List<int> { 1, 2, 3, 4 };

        // Full access roles: Developer (1), Admin (2), Head (3)
        private readonly List<int> _fullAccessRoleIds = new List<int> { 1, 2, 3 };

        public AuthorizationService(ERPToolsDbContext context)
        {
            _context = context;
        }

        /// <summary>
        /// Extract role IDs from JWT token claims.
        /// Looks for 'role' claims in the token.
        /// </summary>
        public List<int> ExtractRolesFromToken(ClaimsPrincipal user)
        {
            var roleIds = new List<int>();

            if (user == null)
                return roleIds;

            var roleClaims = user.FindAll(ClaimTypes.Role);

            foreach (var roleClaim in roleClaims)
            {
                if (int.TryParse(roleClaim.Value, out var roleId))
                {
                    roleIds.Add(roleId);
                }
            }

            return roleIds;
        }

        /// <summary>
        /// Check if role is an authorized role (Developer, Admin, Head, Manager)
        /// </summary>
        public bool IsAuthorizedRole(List<int> userRoleIds)
        {
            if (userRoleIds == null || userRoleIds.Count == 0)
                return false;

            return userRoleIds.Any(roleId => _authorizedRoleIds.Contains(roleId));
        }

        /// <summary>
        /// Check if user has a full-access role (Developer, Admin, Head)
        /// </summary>
        public bool HasFullAccessRole(List<int> userRoleIds)
        {
            if (userRoleIds == null || userRoleIds.Count == 0)
                return false;

            return userRoleIds.Any(roleId => _fullAccessRoleIds.Contains(roleId));
        }

        /// <summary>
        /// Check if user can edit master configuration for a specific project
        /// Business Rules:
        /// 1. Developer (1), Admin (2), Head (3) have global access
        /// 2. Manager (4) has global access
        /// 3. All other roles are denied
        /// </summary>
        public async Task<bool> CanEditMasterConfigurationAsync(int userId, int projectId, ClaimsPrincipal user)
        {
            var userRoleIds = ExtractRolesFromToken(user);

            // Check if user has authorized role
            return IsAuthorizedRole(userRoleIds);
        }

        /// <summary>
        /// Check if user can edit RPT master for a specific project
        /// Same rules as master configuration
        /// </summary>
        public async Task<bool> CanEditRPTMasterAsync(int userId, int projectId, ClaimsPrincipal user)
        {
            var userRoleIds = ExtractRolesFromToken(user);

            // Check if user has authorized role
            return IsAuthorizedRole(userRoleIds);
        }

        /// <summary>
        /// Generic authorization check for any module/feature
        /// Can be reused for Lots, Catches, Pipelines, Reports, etc.
        /// </summary>
        public async Task<bool> IsAuthorizedForProjectAsync(int userId, int projectId, ClaimsPrincipal user)
        {
            var userRoleIds = ExtractRolesFromToken(user);

            // Check if user has authorized role
            return IsAuthorizedRole(userRoleIds);
        }

        /// <summary>
        /// Get all groups the user is assigned to across all projects, or for a specific project
        /// </summary>
        public async Task<List<int>> GetUserGroupsForProjectAsync(int userId, int projectId)
        {
            // Placeholder - returns empty for now
            return new List<int>();
        }
    }
}
