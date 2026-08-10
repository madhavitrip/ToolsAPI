using System;
using System.Collections.Concurrent;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using ERPToolsAPI.Data;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using ToolsAPI.Models;

namespace Tools.Services
{
    public interface IMasterAuthService
    {
        bool IsPasscodeSetForGroup(int groupId);
        Task<(bool IsValid, string ErrorMessage)> VerifyPasscodeAsync(string passcode, int groupId, int userId, string operation, string module, string ipAddress);
        Task LogAuditAsync(int groupId, int userId, string operation, string module, string ipAddress, bool success, string message);
        (bool IsLockedOut, int RemainingMinutes) CheckLockoutStatus(int groupId, int userId, string ipAddress);
        void ResetLockout(int groupId, int userId, string ipAddress);
        void ClearAllLockouts();
        Task<(bool Success, string Error)> UpdatePasscodeAsync(string currentPasscode, string newPasscode, int groupId, int userId, string ipAddress);
        string HashPasscode(string passcode);
        int ResolveGroupIdFromProjectId(int projectId);
    }

    public class MasterAuthService : IMasterAuthService
    {
        private readonly ERPToolsDbContext _context;
        private readonly IOptions<MasterAuthSettings> _settings;
        private readonly ILogger<MasterAuthService> _logger;

        // In-memory rate-limiting and brute force lockout tracker
        private static readonly ConcurrentDictionary<string, FailedAttemptInfo> _failedAttempts = new ConcurrentDictionary<string, FailedAttemptInfo>();

        private class FailedAttemptInfo
        {
            public int AttemptCount { get; set; }
            public DateTime? LockoutEndTime { get; set; }
            public DateTime LastAttemptTime { get; set; }
        }

        public MasterAuthService(ERPToolsDbContext context, IOptions<MasterAuthSettings> settings, ILogger<MasterAuthService> logger)
        {
            _context = context;
            _settings = settings;
            _logger = logger;
        }

        public bool IsPasscodeSetForGroup(int groupId)
        {
            if (groupId <= 0)
            {
                // Fallback to global setting if no groupId specified
                return !string.IsNullOrWhiteSpace(_settings.Value?.PasscodeHash);
            }

            var record = _context.GroupMasterAuthPasscodes.AsNoTracking().FirstOrDefault(g => g.GroupId == groupId);
            return record != null && !string.IsNullOrWhiteSpace(record.PasscodeHash);
        }

        public int ResolveGroupIdFromProjectId(int projectId)
        {
            if (projectId <= 0) return 0;
            try
            {
                var project = _context.Projects.AsNoTracking().FirstOrDefault(p => p.ProjectId == projectId);
                return project?.GroupId ?? 0;
            }
            catch
            {
                return 0;
            }
        }

        public (bool IsLockedOut, int RemainingMinutes) CheckLockoutStatus(int groupId, int userId, string ipAddress)
        {
            var trackerKey = GetTrackerKey(groupId, userId, ipAddress);
            if (_failedAttempts.TryGetValue(trackerKey, out var info))
            {
                if (info.LockoutEndTime.HasValue && info.LockoutEndTime.Value > DateTime.UtcNow)
                {
                    var remaining = (int)Math.Ceiling((info.LockoutEndTime.Value - DateTime.UtcNow).TotalMinutes);
                    return (true, remaining <= 0 ? 1 : remaining);
                }

                if (info.LockoutEndTime.HasValue && info.LockoutEndTime.Value <= DateTime.UtcNow)
                {
                    _failedAttempts.TryRemove(trackerKey, out _);
                }
            }

            return (false, 0);
        }

        public void ResetLockout(int groupId, int userId, string ipAddress)
        {
            var trackerKey = GetTrackerKey(groupId, userId, ipAddress);
            _failedAttempts.TryRemove(trackerKey, out _);
        }

        public void ClearAllLockouts()
        {
            _failedAttempts.Clear();
        }

        public async Task<(bool IsValid, string ErrorMessage)> VerifyPasscodeAsync(string passcode, int groupId, int userId, string operation, string module, string ipAddress)
        {
            var trackerKey = GetTrackerKey(groupId, userId, ipAddress);

            // 1. Check lockout status
            var lockout = CheckLockoutStatus(groupId, userId, ipAddress);
            if (lockout.IsLockedOut)
            {
                var lockMessage = $"Too many failed authorization attempts for Group {groupId}. Locked out for {lockout.RemainingMinutes} minute(s).";
                await LogAuditAsync(groupId, userId, operation, module, ipAddress, false, lockMessage);
                return (false, lockMessage);
            }

            // 2. Check if passcode is configured for this Group
            if (!IsPasscodeSetForGroup(groupId))
            {
                string groupMsg = groupId > 0
                    ? $"No Master Passcode has been set for Group {groupId}. Please create the Group Master PIN first."
                    : "No Master Passcode has been set yet. Please create your Master PIN first.";
                await LogAuditAsync(groupId, userId, operation, module, ipAddress, false, groupMsg);
                return (false, groupMsg);
            }

            // 3. Validate passcode input
            if (string.IsNullOrWhiteSpace(passcode))
            {
                const string emptyMsg = "Master authorization passcode is required.";
                await LogAuditAsync(groupId, userId, operation, module, ipAddress, false, emptyMsg);
                return (false, emptyMsg);
            }

            // 4. Retrieve stored hash for this Group (or fallback to global settings)
            string targetHash = string.Empty;
            if (groupId > 0)
            {
                var record = await _context.GroupMasterAuthPasscodes.AsNoTracking().FirstOrDefaultAsync(g => g.GroupId == groupId);
                targetHash = record?.PasscodeHash ?? string.Empty;
            }
            if (string.IsNullOrEmpty(targetHash))
            {
                targetHash = _settings.Value.PasscodeHash ?? string.Empty;
            }

            bool isValid = VerifyHash(passcode, targetHash);

            if (!isValid)
            {
                var maxAttempts = _settings.Value.MaxFailedAttempts > 0 ? _settings.Value.MaxFailedAttempts : 5;
                var lockoutMins = _settings.Value.LockoutMinutes > 0 ? _settings.Value.LockoutMinutes : 15;

                var info = _failedAttempts.AddOrUpdate(trackerKey,
                    k => new FailedAttemptInfo { AttemptCount = 1, LastAttemptTime = DateTime.UtcNow },
                    (k, existing) =>
                    {
                        existing.AttemptCount += 1;
                        existing.LastAttemptTime = DateTime.UtcNow;
                        if (existing.AttemptCount >= maxAttempts)
                        {
                            existing.LockoutEndTime = DateTime.UtcNow.AddMinutes(lockoutMins);
                        }
                        return existing;
                    });

                string errorMsg;
                if (info.LockoutEndTime.HasValue && info.LockoutEndTime.Value > DateTime.UtcNow)
                {
                    errorMsg = $"Invalid authorization passcode. Account is now locked out for {lockoutMins} minute(s).";
                }
                else
                {
                    int remainingAttempts = maxAttempts - info.AttemptCount;
                    errorMsg = $"Invalid authorization passcode. ({remainingAttempts} attempt(s) remaining before lockout)";
                }

                await LogAuditAsync(groupId, userId, operation, module, ipAddress, false, errorMsg);
                return (false, errorMsg);
            }

            // Authentication succeeded -> Reset lockout tracker
            _failedAttempts.TryRemove(trackerKey, out _);
            await LogAuditAsync(groupId, userId, operation, module, ipAddress, true, $"Master authorization successful for Group {groupId}");
            return (true, "Authorized");
        }

        public async Task LogAuditAsync(int groupId, int userId, string operation, string module, string ipAddress, bool success, string message)
        {
            MasterAuthAuditLog auditLog = null;
            try
            {
                auditLog = new MasterAuthAuditLog
                {
                    UserId = userId,
                    GroupId = groupId,
                    OperationType = operation ?? "MASTER_ACTION",
                    ModuleName = module ?? "GENERAL",
                    IpAddress = ipAddress ?? "UNKNOWN",
                    Success = success,
                    Message = message ?? (success ? "Success" : "Failed"),
                    Timestamp = DateTime.UtcNow
                };

                _context.MasterAuthAuditLogs.Add(auditLog);
                await _context.SaveChangesAsync();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to record MasterAuth Audit Log for GroupId {GroupId}, UserId {UserId}", groupId, userId);
                if (auditLog != null)
                {
                    _context.Entry(auditLog).State = EntityState.Detached;
                }
            }
        }

        public async Task<(bool Success, string Error)> UpdatePasscodeAsync(string currentPasscode, string newPasscode, int groupId, int userId, string ipAddress)
        {
            bool isSet = IsPasscodeSetForGroup(groupId);

            // If passcode is already set for this group, verify current passcode
            if (isSet)
            {
                string existingHash = string.Empty;
                if (groupId > 0)
                {
                    var rec = await _context.GroupMasterAuthPasscodes.AsNoTracking().FirstOrDefaultAsync(g => g.GroupId == groupId);
                    existingHash = rec?.PasscodeHash ?? string.Empty;
                }
                if (string.IsNullOrEmpty(existingHash))
                {
                    existingHash = _settings.Value.PasscodeHash ?? string.Empty;
                }

                if (!VerifyHash(currentPasscode, existingHash))
                {
                    return (false, "Current master passcode is incorrect.");
                }
            }

            if (string.IsNullOrWhiteSpace(newPasscode) || newPasscode.Length < 4)
            {
                return (false, "New master passcode must be at least 4 characters long.");
            }

            string newHash = HashPasscode(newPasscode);

            // Save groupwise in database
            if (groupId > 0)
            {
                var record = await _context.GroupMasterAuthPasscodes.FirstOrDefaultAsync(g => g.GroupId == groupId);
                if (record == null)
                {
                    record = new GroupMasterAuthPasscode
                    {
                        GroupId = groupId,
                        PasscodeHash = newHash,
                        UpdatedByUserId = userId,
                        UpdatedAt = DateTime.UtcNow
                    };
                    _context.GroupMasterAuthPasscodes.Add(record);
                }
                else
                {
                    record.PasscodeHash = newHash;
                    record.UpdatedByUserId = userId;
                    record.UpdatedAt = DateTime.UtcNow;
                    _context.GroupMasterAuthPasscodes.Update(record);
                }
                await _context.SaveChangesAsync();
            }
            else
            {
                // Global fallback
                _settings.Value.PasscodeHash = newHash;
            }

            ResetLockout(groupId, userId, ipAddress);
            ClearAllLockouts();
            await LogAuditAsync(groupId, userId, "UPDATE_PASSCODE", "MasterAuth", ipAddress, true, $"Updated Master Authorization Passcode for Group {groupId}");

            return (true, string.Empty);
        }

        private string GetTrackerKey(int groupId, int userId, string ipAddress)
        {
            return $"group_{groupId}_user_{userId}_ip_{ipAddress}";
        }

        public string HashPasscode(string passcode)
        {
            using (var sha256 = SHA256.Create())
            {
                var salt = Encoding.UTF8.GetBytes("ERPToolsMasterAuthSalt2026");
                var passBytes = Encoding.UTF8.GetBytes(passcode ?? string.Empty);
                var combined = new byte[salt.Length + passBytes.Length];
                Buffer.BlockCopy(salt, 0, combined, 0, salt.Length);
                Buffer.BlockCopy(passBytes, 0, combined, salt.Length, passBytes.Length);

                var hashBytes = sha256.ComputeHash(combined);
                return Convert.ToBase64String(hashBytes);
            }
        }

        private bool VerifyHash(string passcode, string storedHash)
        {
            if (string.IsNullOrEmpty(storedHash)) return false;
            var computedHash = HashPasscode(passcode);
            return string.Equals(computedHash, storedHash, StringComparison.Ordinal);
        }
    }
}
