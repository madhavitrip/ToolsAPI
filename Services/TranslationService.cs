using ERPToolsAPI.Data;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Caching.Memory;
using System.Text.Json;
using Tools.Services;
using ToolsAPI.Models;

namespace ToolsAPI.Services
{
    // ==========================================
    // 1. TRANSLATION SERVICE CONTRACT
    // ==========================================
    public interface ITranslationService
    {
        Task<Dictionary<string, string>> GetBundleAsync(string lang, CancellationToken cancellationToken = default);
        Task<PagedResult<MessageDto>> GetMessagesAsync(string? module, string? search, bool? missingHindiOnly, int pageNumber, int pageSize, CancellationToken cancellationToken = default);
        Task<MessageDto?> GetMessageByIdAsync(int id, CancellationToken cancellationToken = default);
        Task<MessageDto> CreateMessageAsync(CreateMessageRequest request, int userId, CancellationToken cancellationToken = default);
        Task<MessageDto> UpdateMessageAsync(int id, UpdateMessageRequest request, int userId, CancellationToken cancellationToken = default);
        Task<bool> UpdateStatusAsync(int id, bool isActive, int userId, CancellationToken cancellationToken = default);
        void InvalidateBundleCache();
    }

    // ==========================================
    // 2. TRANSLATION SERVICE IMPLEMENTATION
    // ==========================================
    public class TranslationService : ITranslationService
    {
        private readonly ERPToolsDbContext _context;
        private readonly IMemoryCache _cache;
        private readonly ILoggerService _loggerService;
        private const string CacheKeyPrefix = "app_translation_bundle_";
        private static readonly TimeSpan CacheDuration = TimeSpan.FromSeconds(30);

        public TranslationService(ERPToolsDbContext context, IMemoryCache cache, ILoggerService loggerService)
        {
            _context = context;
            _cache = cache;
            _loggerService = loggerService;
        }

        public async Task<Dictionary<string, string>> GetBundleAsync(string lang, CancellationToken cancellationToken = default)
        {
            var normalizedLang = (lang ?? "en").Trim().ToLowerInvariant();
            var cacheKey = $"{CacheKeyPrefix}{normalizedLang}";

            if (_cache.TryGetValue(cacheKey, out Dictionary<string, string>? cachedBundle) && cachedBundle != null)
            {
                return cachedBundle;
            }

            var query = _context.ApplicationMessages.AsNoTracking().Where(m => m.IsActive);
            Dictionary<string, string> bundle;

            if (normalizedLang == "hi")
            {
                var list = await query
                    .Where(m => m.HindiText != null && m.HindiText.Trim() != "")
                    .Select(m => new { m.MessageKey, Text = m.HindiText! })
                    .ToListAsync(cancellationToken);

                bundle = list.ToDictionary(x => x.MessageKey, x => x.Text, StringComparer.OrdinalIgnoreCase);
            }
            else
            {
                var list = await query
                    .Where(m => m.EnglishText != null && m.EnglishText.Trim() != "")
                    .Select(m => new { m.MessageKey, Text = m.EnglishText })
                    .ToListAsync(cancellationToken);

                bundle = list.ToDictionary(x => x.MessageKey, x => x.Text, StringComparer.OrdinalIgnoreCase);
            }

            _cache.Set(cacheKey, bundle, CacheDuration);
            return bundle;
        }

        public async Task<PagedResult<MessageDto>> GetMessagesAsync(string? module, string? search, bool? missingHindiOnly, int pageNumber, int pageSize, CancellationToken cancellationToken = default)
        {
            var query = _context.ApplicationMessages.AsNoTracking().AsQueryable();

            if (!string.IsNullOrWhiteSpace(module))
            {
                query = query.Where(m => m.Module == module.Trim());
            }

            if (!string.IsNullOrWhiteSpace(search))
            {
                var s = search.Trim();
                query = query.Where(m => m.MessageKey.Contains(s) || m.EnglishText.Contains(s) || (m.HindiText != null && m.HindiText.Contains(s)));
            }

            if (missingHindiOnly == true)
            {
                query = query.Where(m => m.HindiText == null || m.HindiText.Trim() == "");
            }

            var totalCount = await query.CountAsync(cancellationToken);

            var items = await query
                .OrderBy(m => m.Module)
                .ThenBy(m => m.MessageKey)
                .Skip((pageNumber - 1) * pageSize)
                .Take(pageSize)
                .Select(m => new MessageDto
                {
                    Id = m.Id,
                    MessageKey = m.MessageKey,
                    Module = m.Module,
                    Description = m.Description,
                    EnglishText = m.EnglishText,
                    HindiText = m.HindiText,
                    IsActive = m.IsActive,
                    CreatedAt = m.CreatedAt,
                    UpdatedAt = m.UpdatedAt
                })
                .ToListAsync(cancellationToken);

            return new PagedResult<MessageDto>
            {
                Items = items,
                TotalCount = totalCount,
                PageNumber = pageNumber,
                PageSize = pageSize
            };
        }

        public async Task<MessageDto?> GetMessageByIdAsync(int id, CancellationToken cancellationToken = default)
        {
            var m = await _context.ApplicationMessages.AsNoTracking().FirstOrDefaultAsync(x => x.Id == id, cancellationToken);
            if (m == null) return null;

            return new MessageDto
            {
                Id = m.Id,
                MessageKey = m.MessageKey,
                Module = m.Module,
                Description = m.Description,
                EnglishText = m.EnglishText,
                HindiText = m.HindiText,
                IsActive = m.IsActive,
                CreatedAt = m.CreatedAt,
                UpdatedAt = m.UpdatedAt
            };
        }

        public async Task<MessageDto> CreateMessageAsync(CreateMessageRequest request, int userId, CancellationToken cancellationToken = default)
        {
            var normalizedKey = request.MessageKey.Trim();
            var exists = await _context.ApplicationMessages.AnyAsync(m => m.MessageKey == normalizedKey, cancellationToken);
            if (exists)
            {
                throw new InvalidOperationException($"MessageKey '{normalizedKey}' already exists.");
            }

            var msg = new ApplicationMessage
            {
                MessageKey = normalizedKey,
                Module = request.Module.Trim(),
                Description = request.Description?.Trim(),
                EnglishText = request.EnglishText.Trim(),
                HindiText = string.IsNullOrWhiteSpace(request.HindiText) ? null : request.HindiText.Trim(),
                IsActive = true,
                CreatedAt = DateTime.UtcNow,
                UpdatedAt = DateTime.UtcNow
            };

            _context.ApplicationMessages.Add(msg);
            await _context.SaveChangesAsync(cancellationToken);

            InvalidateBundleCache();

            await _loggerService.LogEventAsync(
                message: $"Created translation message: {msg.MessageKey}",
                category: "MessageManagement",
                triggeredBy: userId,
                ProjectId: 0,
                oldValue: null,
                newValue: JsonSerializer.Serialize(new { msg.MessageKey, msg.Module, msg.EnglishText, msg.HindiText })
            );

            return new MessageDto
            {
                Id = msg.Id,
                MessageKey = msg.MessageKey,
                Module = msg.Module,
                Description = msg.Description,
                EnglishText = msg.EnglishText,
                HindiText = msg.HindiText,
                IsActive = msg.IsActive,
                CreatedAt = msg.CreatedAt,
                UpdatedAt = msg.UpdatedAt
            };
        }

        public async Task<MessageDto> UpdateMessageAsync(int id, UpdateMessageRequest request, int userId, CancellationToken cancellationToken = default)
        {
            var msg = await _context.ApplicationMessages.FirstOrDefaultAsync(m => m.Id == id, cancellationToken);
            if (msg == null)
            {
                throw new KeyNotFoundException($"Message with ID {id} not found.");
            }

            var oldSnapshot = JsonSerializer.Serialize(new { msg.Description, msg.EnglishText, msg.HindiText, msg.IsActive });

            msg.Description = request.Description?.Trim();
            msg.EnglishText = request.EnglishText.Trim();
            msg.HindiText = string.IsNullOrWhiteSpace(request.HindiText) ? null : request.HindiText.Trim();
            msg.IsActive = request.IsActive;
            msg.UpdatedAt = DateTime.UtcNow;

            await _context.SaveChangesAsync(cancellationToken);

            InvalidateBundleCache();

            var newSnapshot = JsonSerializer.Serialize(new { msg.Description, msg.EnglishText, msg.HindiText, msg.IsActive });

            await _loggerService.LogEventAsync(
                message: $"Updated translation message: {msg.MessageKey}",
                category: "MessageManagement",
                triggeredBy: userId,
                ProjectId: 0,
                oldValue: oldSnapshot,
                newValue: newSnapshot
            );

            return new MessageDto
            {
                Id = msg.Id,
                MessageKey = msg.MessageKey,
                Module = msg.Module,
                Description = msg.Description,
                EnglishText = msg.EnglishText,
                HindiText = msg.HindiText,
                IsActive = msg.IsActive,
                CreatedAt = msg.CreatedAt,
                UpdatedAt = msg.UpdatedAt
            };
        }

        public async Task<bool> UpdateStatusAsync(int id, bool isActive, int userId, CancellationToken cancellationToken = default)
        {
            var msg = await _context.ApplicationMessages.FirstOrDefaultAsync(m => m.Id == id, cancellationToken);
            if (msg == null)
            {
                throw new KeyNotFoundException($"Message with ID {id} not found.");
            }

            var oldStatus = msg.IsActive;
            msg.IsActive = isActive;
            msg.UpdatedAt = DateTime.UtcNow;

            await _context.SaveChangesAsync(cancellationToken);

            InvalidateBundleCache();

            await _loggerService.LogEventAsync(
                message: $"Toggled status of translation message: {msg.MessageKey} to {(isActive ? "Active" : "Inactive")}",
                category: "MessageManagement",
                triggeredBy: userId,
                ProjectId: 0,
                oldValue: oldStatus.ToString(),
                newValue: isActive.ToString()
            );

            return true;
        }

        public void InvalidateBundleCache()
        {
            _cache.Remove($"{CacheKeyPrefix}en");
            _cache.Remove($"{CacheKeyPrefix}hi");
        }
    }
}
