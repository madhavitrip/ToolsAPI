using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.Extensions.Options;
using Tools.Models;

namespace Tools.Services
{
    public interface IDispatchService
    {
        Task<DispatchInfo> GetDispatchDateAsync(int projectId, int lotNo);
        Task<Dictionary<int, DispatchInfo>> GetDispatchDatesAsync(int projectId, List<int> lotNos, int maxConcurrency = 5);
    }

    public class DispatchService : IDispatchService
    {
        private readonly HttpClient _httpClient;
        private readonly ApiSettings _apiSettings;
        private readonly ILoggerService _loggerService;

        public DispatchService(HttpClient httpClient, IOptions<ApiSettings> apiSettings, ILoggerService loggerService)
        {
            _httpClient = httpClient;
            _apiSettings = apiSettings.Value;
            _loggerService = loggerService;
        }

        public async Task<DispatchInfo> GetDispatchDateAsync(int projectId, int lotNo)
        {
            try
            {
                var url = $"{_apiSettings.DispatchApiUrl}/GetDispatchDate?projectId={projectId}&lotNo={lotNo}";
                var response = await _httpClient.GetAsync(url);

                if (!response.IsSuccessStatusCode)
                {
                    _loggerService.LogError(
                        $"Dispatch API call failed for Project {projectId}, Lot {lotNo}",
                        $"Status: {response.StatusCode}",
                        nameof(DispatchService)
                    );
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
                _loggerService.LogError(
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
            var tasks = lotNos.Select(async lotNo =>
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
    }

    public class DispatchInfo
    {
        public int ProjectId { get; set; }
        public int LotNo { get; set; }
        public DateTime? DispatchDate { get; set; }
        public bool IsDispatched => DispatchDate.HasValue;
    }

    public class DispatchApiResponse
    {
        public DateTime? DispatchDate { get; set; }
    }
}
