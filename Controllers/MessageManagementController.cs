using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Tools.Middleware;
using Tools.Models;
using Tools.Services;
using ToolsAPI.Models;
using ToolsAPI.Services;

namespace ToolsAPI.Controllers
{
    // ==========================================
    // 1. PUBLIC TRANSLATIONS CONTROLLER
    // ==========================================
    [ApiController]
    [Route("api/[controller]")]
    public class TranslationsController : ControllerBase
    {
        private readonly ITranslationService _translationService;

        public TranslationsController(ITranslationService translationService)
        {
            _translationService = translationService;
        }

        [AllowAnonymous]
        [HttpGet("bundle")]
        public async Task<IActionResult> GetBundle([FromQuery] string lang = "en")
        {
            try
            {
                var bundle = await _translationService.GetBundleAsync(lang, HttpContext.RequestAborted);
                return Ok(bundle);
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { message = "Failed to load translation bundle", error = ex.Message });
            }
        }
    }

    // ==========================================
    // 2. AUTHORIZED MESSAGE MANAGEMENT CONTROLLER
    // ==========================================
    [Authorize]
    [ApiController]
    [Route("api/[controller]")]
    public class MessageManagementController : ControllerBase
    {
        private readonly ITranslationService _translationService;

        public MessageManagementController(ITranslationService translationService)
        {
            _translationService = translationService;
        }

        [HttpGet]
        [RequirePermission(ModuleCodes.MessageManagement, PermissionAction.View)]
        public async Task<IActionResult> GetMessages(
            [FromQuery] string? module,
            [FromQuery] string? search,
            [FromQuery] bool? missingHindiOnly,
            [FromQuery] string? keySearch,
            [FromQuery] string? englishSearch,
            [FromQuery] string? hindiSearch,
            [FromQuery] string? sortBy,
            [FromQuery] string? sortOrder,
            [FromQuery] int pageNumber = 1,
            [FromQuery] int pageSize = 20)
        {
            var result = await _translationService.GetMessagesAsync(
                module,
                search,
                missingHindiOnly,
                keySearch,
                englishSearch,
                hindiSearch,
                sortBy,
                sortOrder,
                pageNumber,
                pageSize,
                HttpContext.RequestAborted);
            return Ok(result);
        }

        [HttpGet("{id:int}")]
        [RequirePermission(ModuleCodes.MessageManagement, PermissionAction.View)]
        public async Task<IActionResult> GetMessageById(int id)
        {
            var message = await _translationService.GetMessageByIdAsync(id, HttpContext.RequestAborted);
            if (message == null)
            {
                return NotFound(new { message = $"Message with ID {id} not found." });
            }

            return Ok(message);
        }

        [HttpPost]
        [RequirePermission(ModuleCodes.MessageManagement, PermissionAction.Create)]
        public async Task<IActionResult> CreateMessage([FromBody] CreateMessageRequest request)
        {
            if (string.IsNullOrWhiteSpace(request.MessageKey) || string.IsNullOrWhiteSpace(request.Module) || string.IsNullOrWhiteSpace(request.EnglishText))
            {
                return BadRequest(new { message = "MessageKey, Module, and EnglishText are required." });
            }

            int userId = LogHelper.GetTriggeredBy(User, Request);

            try
            {
                var created = await _translationService.CreateMessageAsync(request, userId, HttpContext.RequestAborted);
                return CreatedAtAction(nameof(GetMessageById), new { id = created.Id }, created);
            }
            catch (InvalidOperationException ex)
            {
                return Conflict(new { message = ex.Message });
            }
        }

        [HttpPut("{id:int}")]
        [RequirePermission(ModuleCodes.MessageManagement, PermissionAction.Edit)]
        public async Task<IActionResult> UpdateMessage(int id, [FromBody] UpdateMessageRequest request)
        {
            if (string.IsNullOrWhiteSpace(request.EnglishText))
            {
                return BadRequest(new { message = "EnglishText is required." });
            }

            int userId = LogHelper.GetTriggeredBy(User, Request);

            try
            {
                var updated = await _translationService.UpdateMessageAsync(id, request, userId, HttpContext.RequestAborted);
                return Ok(updated);
            }
            catch (KeyNotFoundException ex)
            {
                return NotFound(new { message = ex.Message });
            }
        }

        [HttpPatch("{id:int}/status")]
        [RequirePermission(ModuleCodes.MessageManagement, PermissionAction.Edit)]
        public async Task<IActionResult> UpdateStatus(int id, [FromBody] UpdateStatusRequest request)
        {
            int userId = LogHelper.GetTriggeredBy(User, Request);

            try
            {
                await _translationService.UpdateStatusAsync(id, request.IsActive, userId, HttpContext.RequestAborted);
                return Ok(new { success = true, isActive = request.IsActive });
            }
            catch (KeyNotFoundException ex)
            {
                return NotFound(new { message = ex.Message });
            }
        }
    }
}
