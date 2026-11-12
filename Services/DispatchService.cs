using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Json;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Options;
using Tools.Models;
using ERPToolsAPI.Data;

namespace Tools.Services
{
    public interface IDispatchService
    {
        Task<DispatchInfo> GetDispatchDateAsync(int projectId, int lotNo);
        Task<Dictionary<int, DispatchInfo>> GetDispatchDatesAsync(int projectId, List<int> lotNos, int maxConcurrency = 5);
        Task<bool> SyncAndValidateLotDispatchStatusAsync(int projectId, int lotNo);
        Task<Dictionary<int, bool>> SyncAndValidateLotDispatchStatusesAsync(int projectId, List<int> lotNos);
        Task<List<int>> SyncDispatchedLotRangesForProjectAsync(int projectId);
        Task<bool> IsLotDispatchedAsync(int projectId, int lotNo);
    }

    public class DispatchService : IDispatchService
    {
        private readonly HttpClient _httpClient;
        private readonly ApiSettings _apiSettings;
        private readonly ILoggerService _loggerService;
        private readonly IServiceProvider _serviceProvider;

        public DispatchService(
            HttpClient httpClient,
            IOptions<ApiSettings> apiSettings,
            ILoggerService loggerService,
            IServiceProvider serviceProvider)
        {
            _httpClient = httpClient;
            _apiSettings = apiSettings.Value;
            _loggerService = loggerService;
            _serviceProvider = serviceProvider;
        }

        public async Task<DispatchInfo> GetDispatchDateAsync(int projectId, int lotNo)
        {
            try
            {
                var url = $"{_apiSettings.DispatchApiUrl}/GetDispatchDate?projectId={projectId}&lotNo={lotNo}";
                var response = await _httpClient.GetAsync(url);

                if (!response.IsSuccessStatusCode)
                {
                    if (response.StatusCode != System.Net.HttpStatusCode.NotFound)
                    {
                        await _loggerService.LogErrorAsync(
                            $"Dispatch API call failed for Project {projectId}, Lot {lotNo}",
                            $"Status: {response.StatusCode}",
                            nameof(DispatchService)
                        );
                    }
                    return new DispatchInfo { ProjectId = projectId, LotNo = lotNo, DispatchDate = null };
                }

                var result = await response.Content.ReadFromJsonAsync<DispatchApiResponse>();
                return new DispatchInfo
                {
                    ProjectId = projectId,
                    LotNo = lotNo,
                    DispatchDate = result?.DispatchDate
                };
            }
            catch (Exception ex)
            {
                await _loggerService.LogErrorAsync(
                    $"Exception calling dispatch API for Project {projectId}, Lot {lotNo}",
                    ex.Message,
                    nameof(DispatchService)
                );
                return new DispatchInfo { ProjectId = projectId, LotNo = lotNo, DispatchDate = null };
            }
        }

        public async Task<Dictionary<int, DispatchInfo>> GetDispatchDatesAsync(int projectId, List<int> lotNos, int maxConcurrency = 5)
        {
            var semaphore = new SemaphoreSlim(maxConcurrency);
            var tasks = lotNos.Distinct().Select(async lotNo =>
            {
                await semaphore.WaitAsync();
                try
                {
                    return await GetDispatchDateAsync(projectId, lotNo);
                }
                finally
                {
                    semaphore.Release();
                }
            });

            var results = await Task.WhenAll(tasks);
            return results.ToDictionary(r => r.LotNo, r => r);
        }

        /// <summary>
        /// Checks if a lot's dispatch is done (DispatchDate is present or exceeded).
        /// If dispatched, updates ProjectLotRanges.Status = false (0) for (projectId, lotNo).
        /// Returns true if dispatch is done / exceeded (meaning lot is locked / read-only).
        /// </summary>
        public async Task<bool> SyncAndValidateLotDispatchStatusAsync(int projectId, int lotNo)
        {
            if (projectId <= 0 || lotNo <= 0) return false;

            var dispatchInfo = await GetDispatchDateAsync(projectId, lotNo);
            bool isDispatched = dispatchInfo.IsDispatched;
            bool isExceeded = dispatchInfo.IsDispatchExceeded;
            bool isDispatchDone = isDispatched || isExceeded;

            using var scope = _serviceProvider.CreateScope();
            var dbContext = scope.ServiceProvider.GetRequiredService<ERPToolsDbContext>();

            var lotRange = await dbContext.ProjectLotRanges
                .FirstOrDefaultAsync(r => r.ProjectId == projectId && r.LotNo == lotNo);

            if (isDispatchDone)
            {
                if (lotRange != null && lotRange.Status)
                {
                    lotRange.Status = false;
                    await dbContext.SaveChangesAsync();
                    await _loggerService.LogEventAsync(
                        $"ProjectLotRange status set to 0 (dispatch done) for Project {projectId}, Lot {lotNo}",
                        nameof(DispatchService),
                        0,
                        projectId
                    );
                }
                else if (lotRange == null)
                {
                    var newRange = new ProjectLotRange
                    {
                        ProjectId = projectId,
                        LotNo = lotNo,
                        StartDate = DateTime.UtcNow.ToString("dd-MM-yyyy"),
                        EndDate = DateTime.UtcNow.ToString("dd-MM-yyyy"),
                        Status = false
                    };
                    dbContext.ProjectLotRanges.Add(newRange);
                    await dbContext.SaveChangesAsync();
                    await _loggerService.LogEventAsync(
                        $"ProjectLotRange created with status 0 (dispatched) for Project {projectId}, Lot {lotNo}",
                        nameof(DispatchService),
                        0,
                        projectId
                    );
                }
                return true;
            }
            else if (lotRange != null && !lotRange.Status)
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// Checks multiple lots. If any lot's dispatch is done (DispatchDate is present or exceeded),
        /// updates ProjectLotRanges.Status = false (0). If no ProjectLotRanges record exists for a dispatched lot, creates one with Status = false (0).
        /// Returns a dictionary of LotNo -> isDispatchDone (true = locked/read-only).
        /// </summary>
        public async Task<Dictionary<int, bool>> SyncAndValidateLotDispatchStatusesAsync(int projectId, List<int> lotNos)
        {
            var distinctLots = lotNos.Where(l => l > 0).Distinct().ToList();
            if (distinctLots.Count == 0) return new Dictionary<int, bool>();

            var dispatchInfoDict = await GetDispatchDatesAsync(projectId, distinctLots);
            var dispatchDoneMap = new Dictionary<int, bool>();
            var dispatchedExistingLots = new List<int>();
            var dispatchedMissingLots = new List<int>();

            using var scope = _serviceProvider.CreateScope();
            var dbContext = scope.ServiceProvider.GetRequiredService<ERPToolsDbContext>();

            var existingRanges = await dbContext.ProjectLotRanges
                .Where(r => r.ProjectId == projectId && distinctLots.Contains(r.LotNo))
                .ToListAsync();
            var existingRangesByLot = existingRanges.ToDictionary(r => r.LotNo);

            foreach (var lotNo in distinctLots)
            {
                bool isDbInactive = existingRangesByLot.TryGetValue(lotNo, out var range) && !range.Status;
                bool isDone = isDbInactive || (dispatchInfoDict.TryGetValue(lotNo, out var info) &&
                              (info.IsDispatched || info.IsDispatchExceeded));
                dispatchDoneMap[lotNo] = isDone;
                if (isDone)
                {
                    if (range != null && range.Status)
                    {
                        dispatchedExistingLots.Add(lotNo);
                    }
                    else if (range == null)
                    {
                        dispatchedMissingLots.Add(lotNo);
                    }
                }
            }

            if (dispatchedExistingLots.Count > 0 || dispatchedMissingLots.Count > 0)
            {
                try
                {
                    var rangesToUpdate = existingRanges.Where(r => dispatchedExistingLots.Contains(r.LotNo) && r.Status).ToList();
                    foreach (var range in rangesToUpdate)
                    {
                        range.Status = false;
                    }

                    foreach (var missingLot in dispatchedMissingLots)
                    {
                        var newRange = new ProjectLotRange
                        {
                            ProjectId = projectId,
                            LotNo = missingLot,
                            StartDate = DateTime.UtcNow.ToString("dd-MM-yyyy"),
                            EndDate = DateTime.UtcNow.ToString("dd-MM-yyyy"),
                            Status = false
                        };
                        dbContext.ProjectLotRanges.Add(newRange);
                    }

                    await dbContext.SaveChangesAsync();
                    var allUpdated = dispatchedExistingLots.Concat(dispatchedMissingLots).Distinct().ToList();
                    await _loggerService.LogEventAsync(
                        $"ProjectLotRanges status set to 0 for dispatched lots: {string.Join(",", allUpdated)}",
                        nameof(DispatchService),
                        0,
                        projectId
                    );
                }
                catch (Exception ex)
                {
                    await _loggerService.LogErrorAsync(
                        $"Failed to update ProjectLotRanges status for Project {projectId}",
                        ex.Message,
                        nameof(DispatchService)
                    );
                }
            }

            return dispatchDoneMap;
        }

        /// <summary>
        /// Syncs all ProjectLotRanges records for a project against ERP dispatch API.
        /// Any lot where dispatch is done has its ProjectLotRanges.Status updated to false (0),
        /// and any dispatched lot not yet in ProjectLotRanges is inserted with Status = false (0).
        /// Returns the list of LotNos updated.
        /// </summary>
        public async Task<List<int>> SyncDispatchedLotRangesForProjectAsync(int projectId)
        {
            if (projectId <= 0) return new List<int>();

            try
            {
                using var scope = _serviceProvider.CreateScope();
                var dbContext = scope.ServiceProvider.GetRequiredService<ERPToolsDbContext>();

                var isNewModel = await dbContext.NrData1.AnyAsync(p => p.ProjectId == projectId);
                List<int> projectDataLots = isNewModel
                    ? await dbContext.NrData1.Where(p => p.ProjectId == projectId).Select(p => p.LotNo).Distinct().ToListAsync()
                    : await dbContext.NRDatas.Where(p => p.ProjectId == projectId).Select(p => p.LotNo).Distinct().ToListAsync();

                var existingRanges = await dbContext.ProjectLotRanges
                    .Where(r => r.ProjectId == projectId)
                    .ToListAsync();

                var existingByLot = existingRanges.ToDictionary(r => r.LotNo);
                var allLotNos = projectDataLots.Union(existingRanges.Select(r => r.LotNo)).Where(l => l > 0).Distinct().ToList();

                if (allLotNos.Count == 0) return new List<int>();

                var dispatchDict = await GetDispatchDatesAsync(projectId, allLotNos);
                var updatedLots = new List<int>();

                foreach (var lotNo in allLotNos)
                {
                    if (dispatchDict.TryGetValue(lotNo, out var info) && (info.IsDispatched || info.IsDispatchExceeded))
                    {
                        if (existingByLot.TryGetValue(lotNo, out var range))
                        {
                            if (range.Status)
                            {
                                range.Status = false;
                                updatedLots.Add(lotNo);
                            }
                        }
                        else
                        {
                            var newRange = new ProjectLotRange
                            {
                                ProjectId = projectId,
                                LotNo = lotNo,
                                StartDate = DateTime.UtcNow.ToString("dd-MM-yyyy"),
                                EndDate = DateTime.UtcNow.ToString("dd-MM-yyyy"),
                                Status = false
                            };
                            dbContext.ProjectLotRanges.Add(newRange);
                            updatedLots.Add(lotNo);
                        }
                    }
                }

                if (updatedLots.Count > 0)
                {
                    await dbContext.SaveChangesAsync();
                    await _loggerService.LogEventAsync(
                        $"ProjectLotRanges status set to 0 for dispatched lots: {string.Join(",", updatedLots)}",
                        nameof(DispatchService),
                        0,
                        projectId
                    );
                }

                return updatedLots;
            }
            catch (Exception ex)
            {
                await _loggerService.LogErrorAsync(
                    $"Failed to sync dispatched ProjectLotRanges for Project {projectId}",
                    ex.Message,
                    nameof(DispatchService)
                );
                return new List<int>();
            }
        }

        public async Task<bool> IsLotDispatchedAsync(int projectId, int lotNo)
        {
            if (projectId <= 0 || lotNo <= 0) return false;

            try
            {
                using var scope = _serviceProvider.CreateScope();
                var dbContext = scope.ServiceProvider.GetRequiredService<ERPToolsDbContext>();

                var lotRange = await dbContext.ProjectLotRanges
                    .FirstOrDefaultAsync(r => r.ProjectId == projectId && r.LotNo == lotNo);

                if (lotRange != null && !lotRange.Status)
                {
                    return true;
                }

                var dispatchInfo = await GetDispatchDateAsync(projectId, lotNo);
                if (dispatchInfo.IsDispatched || dispatchInfo.IsDispatchExceeded)
                {
                    if (lotRange != null && lotRange.Status)
                    {
                        lotRange.Status = false;
                        await dbContext.SaveChangesAsync();
                    }
                    return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                await _loggerService.LogErrorAsync(
                    $"Error checking if Lot {lotNo} is dispatched for Project {projectId}",
                    ex.Message,
                    nameof(DispatchService)
                );
                return false;
            }
        }
    }

    public class DispatchInfo
    {
        public int ProjectId { get; set; }
        public int LotNo { get; set; }
        public DateTime? DispatchDate { get; set; }
        public bool IsDispatched => DispatchDate.HasValue;
        public bool IsDispatchExceeded => DispatchDate.HasValue && DispatchDate.Value <= DateTime.Now;
    }

    public class DispatchApiResponse
    {
        public DateTime? DispatchDate { get; set; }
    }
}
