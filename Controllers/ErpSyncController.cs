using System;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Tools.Services;

namespace Tools.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ErpSyncController : ControllerBase
    {
        private readonly IErpSyncService _erpSyncService;
        private readonly ILoggerService _loggerService;

        public ErpSyncController(
            IErpSyncService erpSyncService,
            ILoggerService loggerService)
        {
            _erpSyncService = erpSyncService;
            _loggerService = loggerService;
        }

        /// <summary>
        /// Retrieves all catches with Quantity, Pages, InnerEnv, OuterEnv for a specific project and lot.
        /// Also returns the finalization status of the lot.
        /// </summary>
        [HttpGet("GetLotCatches/{projectId}/{lotNo}")]
        public async Task<IActionResult> GetLotCatches(int projectId, int lotNo)
        {
            if (projectId <= 0 || lotNo <= 0)
                return BadRequest("Valid ProjectId and LotNo are required.");

            try
            {
                var payload = await _erpSyncService.CompileLotCatchesAsync(projectId, lotNo);
                return Ok(payload);
            }
            catch (Exception ex)
            {
                await _loggerService.LogErrorAsync(
                    $"Error fetching lot catches for ERP sync (ProjectId: {projectId}, LotNo: {lotNo})",
                    ex.ToString(),
                    nameof(ErpSyncController)
                );
                return StatusCode(500, new { message = "Error compiling catch details", error = ex.Message });
            }
        }

        /// <summary>
        /// Gets finalization and sync status for all lots in a project.
        /// </summary>
        [HttpGet("GetLotsSyncStatus/{projectId}")]
        public async Task<IActionResult> GetLotsSyncStatus(int projectId)
        {
            if (projectId <= 0)
                return BadRequest("Valid ProjectId is required.");

            try
            {
                var result = await _erpSyncService.GetLotsSyncStatusAsync(projectId);
                return Ok(result);
            }
            catch (Exception ex)
            {
                await _loggerService.LogErrorAsync(
                    $"Error fetching lots sync status for Project {projectId}",
                    ex.ToString(),
                    nameof(ErpSyncController)
                );
                return StatusCode(500, new { message = "Error fetching sync status", error = ex.Message });
            }
        }

        /// <summary>
        /// Finalizes a lot so that its catch details are officially approved and locked for ERP transmission.
        /// </summary>
        [HttpPost("FinalizeLot")]
        public async Task<IActionResult> FinalizeLot([FromBody] FinalizeLotRequest request)
        {
            if (request == null || request.ProjectId <= 0 || request.LotNo <= 0)
                return BadRequest("ProjectId and LotNo are required.");

            try
            {
                var result = await _erpSyncService.FinalizeLotAsync(request.ProjectId, request.LotNo, request.UserId);
                if (!result.Success)
                {
                    return BadRequest(new { message = result.Message });
                }

                return Ok(new
                {
                    success = result.Success,
                    message = result.Message,
                    finalizedAt = result.FinalizedAt,
                    isFinalized = result.IsFinalized
                });
            }
            catch (Exception ex)
            {
                await _loggerService.LogErrorAsync(
                    $"Error finalizing Lot {request.LotNo} for Project {request.ProjectId}",
                    ex.ToString(),
                    nameof(ErpSyncController)
                );
                return StatusCode(500, new { message = "Error finalizing lot", error = ex.Message });
            }
        }

        /// <summary>
        /// Reverts finalization for a lot if adjustments are required before ERP transmission.
        /// </summary>
        [HttpPost("UnfinalizeLot")]
        public async Task<IActionResult> UnfinalizeLot([FromBody] FinalizeLotRequest request)
        {
            if (request == null || request.ProjectId <= 0 || request.LotNo <= 0)
                return BadRequest("ProjectId and LotNo are required.");

            try
            {
                var result = await _erpSyncService.UnfinalizeLotAsync(request.ProjectId, request.LotNo, request.UserId);
                return Ok(new
                {
                    success = result.Success,
                    message = result.Message,
                    isFinalized = result.IsFinalized
                });
            }
            catch (Exception ex)
            {
                await _loggerService.LogErrorAsync(
                    $"Error unfinalizing Lot {request.LotNo} for Project {request.ProjectId}",
                    ex.ToString(),
                    nameof(ErpSyncController)
                );
                return StatusCode(500, new { message = "Error unfinalizing lot", error = ex.Message });
            }
        }

        /// <summary>
        /// Compiles and sends lot catch details to ERP.
        /// STRICT REQUIREMENT: Only allowed if the lot has been finalized.
        /// </summary>
        [HttpPost("SendLotCatchesToErp")]
        public async Task<IActionResult> SendLotCatchesToErp([FromBody] SendLotCatchesRequest request)
        {
            if (request == null || request.ProjectId <= 0 || request.LotNo <= 0)
                return BadRequest("ProjectId and LotNo are required.");

            try
            {
                var result = await _erpSyncService.SendLotCatchesToErpAsync(request.ProjectId, request.LotNo, request.UserId);
                if (!result.Success)
                {
                    return StatusCode(result.StatusCode != 0 ? result.StatusCode : 400, new
                    {
                        success = false,
                        message = result.Message
                    });
                }

                return Ok(new
                {
                    success = result.Success,
                    message = result.Message,
                    transmitted = result.Transmitted,
                    transmissionStatus = result.TransmissionStatus,
                    syncedAt = result.SyncedAt,
                    payload = result.Payload
                });
            }
            catch (Exception ex)
            {
                await _loggerService.LogErrorAsync(
                    $"Error executing SendLotCatchesToErp for Project {request.ProjectId}, Lot {request.LotNo}",
                    ex.ToString(),
                    nameof(ErpSyncController)
                );
                return StatusCode(500, new { message = "Error processing ERP sync", error = ex.Message });
            }
        }
    }
}
