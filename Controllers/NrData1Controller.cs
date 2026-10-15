using ERPToolsAPI.Data;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Options;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.Json;
using System.Threading.Tasks;
using Tools.Models;
using Tools.Services;
using ToolsAPI.Helpers;
using ToolsAPI.Models;

namespace Tools.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class NrData1Controller : ControllerBase
    {
        private readonly ERPToolsDbContext _context;
        private readonly ILoggerService _loggerService;
        private readonly IOptions<ApiSettings> _apiSettings;
        private readonly IDispatchService _dispatchService;

        public NrData1Controller(
            ERPToolsDbContext context,
            ILoggerService loggerService,
            IOptions<ApiSettings> apiSettings,
            IDispatchService dispatchService)
        {
            _context = context;
            _loggerService = loggerService;
            _apiSettings = apiSettings;
            _dispatchService = dispatchService;
        }

        // GET: api/NrData1
        [HttpGet]
        public async Task<ActionResult<IEnumerable<NrData1>>> GetNrData1()
        {
            return await _context.NrData1.Where(x => x.Status).ToListAsync();
        }

        // GET: api/NrData1/5
        [HttpGet("{id}")]
        public async Task<ActionResult> GetNrData1(int id)
        {
            var record = await _context.NrData1.FindAsync(id);
            if (record == null)
            {
                return NotFound();
            }

            var centers = await _context.CenterList
                .Where(c => c.NRDataId == id && c.Status)
                .ToListAsync();

            return Ok(new
            {
                NrData = record,
                Centers = centers
            });
        }

        private static string NormalizeText(string? value)
        {
            return string.IsNullOrWhiteSpace(value) ? string.Empty : value.Trim();
        }

        // GET: api/NrData1/GetByProjectId/5
        [HttpGet("GetByProjectId/{projectId}")]
        public async Task<ActionResult> GetByProjectId(
            int projectId,
            [FromQuery] int pageSize = 50,
            [FromQuery] int pageNo = 1,
            [FromQuery] string? search = null,
            [FromQuery] string? key = null,
            [FromQuery] string? sortField = null,
            [FromQuery] string? sortOrder = null,
            [FromQuery] string? filters = null,
            [FromQuery] int? lotNo = null,
            [FromQuery] int? batchNo = null)
        {
            var query = from nr in _context.NrData1
                        join c in _context.CenterList on nr.Id equals c.NRDataId
                        where nr.ProjectId == projectId && nr.Status && c.Status
                        select new ToolsAPI.Models.NRDataDto
                        {
                            Id = c.Id,
                            NRDataId = nr.Id,
                            ProjectId = nr.ProjectId,
                            CatchNo = nr.CatchNo,
                            CenterCode = c.CenterCode,
                            ExamDate = nr.ExamDate,
                            ExamTime = nr.ExamTime,
                            SubjectName = nr.SubjectName,
                            CourseName = nr.CourseName,
                            NodalCode = c.NodalCode,
                            NRQuantity = c.NRQuantity,
                            Quantity = c.Quantity,
                            Route = c.Route,
                            Pages = nr.Pages,
                            CenterSort = c.CenterSort,
                            NodalSort = c.NodalSort,
                            RouteSort = c.RouteSort,
                            Symbol = null,
                            LotNo = nr.LotNo,
                            EnvLotNo = nr.EnvLotNo,
                            Batch = nr.Batch,
                            NRDatas = nr.NRDatas,
                            Day = nr.Day,
                            District = c.District,
                            DistrictSort = c.DistrictSort,
                            Status = c.Status,
                            Steps = nr.Steps,
                            VerificationStatus = nr.VerificationStatus,
                            Remarksss = nr.Remarksss
                        };

            if (batchNo.HasValue)
            {
                query = query.Where(d => d.Batch == batchNo.Value);
            }

            if (lotNo.HasValue)
            {
                query = query.Where(d => d.LotNo == lotNo.Value);
            }

            // Apply combination filters if provided
            if (!string.IsNullOrWhiteSpace(filters))
            {
                try
                {
                    var filterDict = JsonSerializer.Deserialize<Dictionary<string, string>>(filters);
                    if (filterDict != null)
                    {
                        foreach (var kvp in filterDict)
                        {
                            var filterKey = kvp.Key;
                            var filterVal = kvp.Value?.Trim().ToLowerInvariant();

                            if (string.IsNullOrWhiteSpace(filterVal))
                                continue;

                            var prop = typeof(ToolsAPI.Models.NRDataDto).GetProperties()
                                .FirstOrDefault(p => p.Name.Equals(filterKey, StringComparison.OrdinalIgnoreCase));

                            if (prop != null)
                            {
                                query = prop.Name.ToLowerInvariant() switch
                                {
                                    "catchno" => query.Where(d => d.CatchNo != null && d.CatchNo.ToLower().Contains(filterVal)),
                                    "centercode" => query.Where(d => d.CenterCode != null && d.CenterCode.ToLower().Contains(filterVal)),
                                    "subjectname" => query.Where(d => d.SubjectName != null && d.SubjectName.ToLower().Contains(filterVal)),
                                    "coursename" => query.Where(d => d.CourseName != null && d.CourseName.ToLower().Contains(filterVal)),
                                    "nodalcode" => query.Where(d => d.NodalCode != null && d.NodalCode.ToLower().Contains(filterVal)),
                                    "route" => query.Where(d => d.Route != null && d.Route.ToLower().Contains(filterVal)),
                                    "examdate" => query.Where(d => d.ExamDate != null && d.ExamDate.ToLower().Contains(filterVal)),
                                    "examtime" => query.Where(d => d.ExamTime != null && d.ExamTime.ToLower().Contains(filterVal)),
                                    "symbol" => query.Where(d => d.Symbol != null && d.Symbol.ToLower().Contains(filterVal)),
                                    _ => query
                                };
                            }
                            else
                            {
                                var likePattern = $"%\"{filterKey}\":%";
                                var valLikePattern = $"%\"{filterKey}\":\"%{filterVal}%\"%";
                                query = query.Where(d => d.NRDatas != null && EF.Functions.Like(d.NRDatas, likePattern) && EF.Functions.Like(d.NRDatas, valLikePattern));
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    await _loggerService.LogErrorAsync("Error parsing filters in NrData1.GetByProjectId", ex.Message, nameof(NrData1Controller));
                }
            }

            if (!string.IsNullOrWhiteSpace(search) && !string.IsNullOrWhiteSpace(key))
            {
                search = search.ToLower();
                switch (key)
                {
                    case "CatchNo":
                        query = query.Where(d => d.CatchNo.ToLower().Contains(search));
                        break;
                    case "CenterCode":
                        query = query.Where(d => d.CenterCode.ToLower().Contains(search));
                        break;
                    case "SubjectName":
                        query = query.Where(d => d.SubjectName.ToLower().Contains(search));
                        break;
                    case "CourseName":
                        query = query.Where(d => d.CourseName.ToLower().Contains(search));
                        break;
                    case "NodalCode":
                        query = query.Where(d => d.NodalCode.ToLower().Contains(search));
                        break;
                    case "Route":
                        query = query.Where(d => d.Route.ToLower().Contains(search));
                        break;
                    case "ExamDate":
                        query = query.Where(d => (d.ExamDate != null && d.ExamDate.ToLower().Contains(search)) || Convert.ToString(d.ExamDate).Contains(search));
                        break;
                    case "ExamTime":
                        query = query.Where(d => (d.ExamTime != null && d.ExamTime.ToLower().Contains(search)) || Convert.ToString(d.ExamTime).Contains(search));
                        break;
                    case "NRQuantity":
                        query = query.Where(d => Convert.ToString(d.NRQuantity).Contains(search));
                        break;
                    case "Quantity":
                        query = query.Where(d => Convert.ToString(d.Quantity).Contains(search));
                        break;
                    default:
                        return BadRequest($"Key '{key}' is not searchable.");
                }
            }
            else if (!string.IsNullOrEmpty(search))
            {
                search = search.ToLower();
                query = query.Where(n =>
                    (n.CatchNo ?? "").ToLower().Contains(search) ||
                    (n.CenterCode ?? "").ToLower().Contains(search) ||
                    (n.SubjectName ?? "").ToLower().Contains(search) ||
                    (n.CourseName ?? "").ToLower().Contains(search) ||
                    (n.NodalCode ?? "").ToLower().Contains(search) ||
                    (n.Route ?? "").ToLower().Contains(search) ||
                    (n.NRDatas ?? "").ToLower().Contains(search)
                );
            }

            var normalizedSortField = NormalizeText(sortField);
            var normalizedSortOrder = NormalizeText(sortOrder).ToLowerInvariant();
            var isAscending = normalizedSortOrder != "descend";

            if (!string.IsNullOrWhiteSpace(normalizedSortField))
            {
                query = normalizedSortField switch
                {
                    "CatchNo" => isAscending ? query.OrderBy(d => d.CatchNo) : query.OrderByDescending(d => d.CatchNo),
                    "CenterCode" => isAscending ? query.OrderBy(d => d.CenterCode) : query.OrderByDescending(d => d.CenterCode),
                    "ExamDate" => isAscending ? query.OrderBy(d => d.ExamDate) : query.OrderByDescending(d => d.ExamDate),
                    "ExamTime" => isAscending ? query.OrderBy(d => d.ExamTime) : query.OrderByDescending(d => d.ExamTime),
                    "NRQuantity" => isAscending ? query.OrderBy(d => d.NRQuantity) : query.OrderByDescending(d => d.NRQuantity),
                    "Quantity" => isAscending ? query.OrderBy(d => d.Quantity) : query.OrderByDescending(d => d.Quantity),
                    "NodalCode" => isAscending ? query.OrderBy(d => d.NodalCode) : query.OrderByDescending(d => d.NodalCode),
                    "CourseName" => isAscending ? query.OrderBy(d => d.CourseName) : query.OrderByDescending(d => d.CourseName),
                    "SubjectName" => isAscending ? query.OrderBy(d => d.SubjectName) : query.OrderByDescending(d => d.SubjectName),
                    "Pages" => isAscending ? query.OrderBy(d => d.Pages) : query.OrderByDescending(d => d.Pages),
                    "CenterSort" => isAscending ? query.OrderBy(d => d.CenterSort) : query.OrderByDescending(d => d.CenterSort),
                    "LotNo" => isAscending ? query.OrderBy(d => d.LotNo) : query.OrderByDescending(d => d.LotNo),
                    "Route" => isAscending ? query.OrderBy(d => d.Route) : query.OrderByDescending(d => d.Route),
                    "District" => isAscending ? query.OrderBy(d => d.District) : query.OrderByDescending(d => d.District),
                    "EnvLotNo" => isAscending ? query.OrderBy(d => d.EnvLotNo) : query.OrderByDescending(d => d.EnvLotNo),
                    "Batch" => isAscending ? query.OrderBy(d => d.Batch) : query.OrderByDescending(d => d.Batch),
                    _ => query.OrderBy(d => d.Id),
                };
            }
            else
            {
                query = query.OrderBy(d => d.Id);
            }

            int totalCount = await query.CountAsync();
            int totalPages = (int)Math.Ceiling(totalCount / (double)pageSize);

            var nrDataList = await query
                .Skip((pageNo - 1) * pageSize)
                .Take(pageSize)
                .Select(d => new
                {
                    d.Id,
                    d.CatchNo,
                    d.CenterCode,
                    d.ExamDate,
                    d.ExamTime,
                    d.SubjectName,
                    d.CourseName,
                    d.NodalCode,
                    d.NRQuantity,
                    d.Quantity,
                    d.Route,
                    d.Pages,
                    d.CenterSort,
                    d.NodalSort,
                    d.RouteSort,
                    d.Symbol,
                    d.LotNo,
                    d.EnvLotNo,
                    d.Batch,
                    d.NRDatas,
                    d.Day,
                    d.District
                })
                .ToListAsync();

            if (!nrDataList.Any())
            {
                var allColumns = typeof(ToolsAPI.Models.NRDataDto).GetProperties()
                    .Select(p => p.Name)
                    .Where(name => name != "Id" && name != "ProjectId" && name != "NRDatas" && name != "NRDataId" && name != "Status" && name != "Steps" && name != "VerificationStatus" && name != "Remarksss")
                    .ToList();

                return Ok(new
                {
                    items = new List<object>(),
                    columns = allColumns,
                    totalCount = 0,
                    totalPages = 0
                });
            }

            var properties = nrDataList.First().GetType().GetProperties();
            var result = nrDataList.Select(d =>
            {
                var dict = new Dictionary<string, object>();
                foreach (var prop in properties)
                {
                    dict[prop.Name] = prop.GetValue(d);
                }
                return dict;
            });

            return Ok(new
            {
                items = result,
                columns = properties.Select(p => p.Name).ToList(),
                totalCount,
                totalPages
            });
        }

        // PUT: api/NrData1/UpdateSingle/{id}
        [HttpPut("UpdateSingle/{id}")]
        public async Task<IActionResult> UpdateSingle(int id, [FromBody] JsonElement inputData, [FromQuery] int? lotNo = null)
        {
            using var transaction = await _context.Database.BeginTransactionAsync();
            try
            {
                var center = await _context.CenterList.FirstOrDefaultAsync(c => c.Id == id);
                NrData1? nr1 = null;

                if (center != null)
                {
                    nr1 = await _context.NrData1.FirstOrDefaultAsync(n => n.Id == center.NRDataId);
                }
                else
                {
                    nr1 = await _context.NrData1.FirstOrDefaultAsync(n => n.Id == id);
                }

                if (center == null && nr1 == null)
                {
                    return NotFound($"Record with ID {id} not found.");
                }

                // Check for soft delete
                bool isDelete = false;
                if (inputData.TryGetProperty("Status", out var statusProp) || inputData.TryGetProperty("status", out statusProp))
                {
                    if (statusProp.ValueKind == JsonValueKind.False ||
                        (statusProp.ValueKind == JsonValueKind.String && statusProp.GetString()?.ToLower() == "false"))
                    {
                        isDelete = true;
                    }
                }

                if (isDelete)
                {
                    if (center != null)
                    {
                        center.Status = false;
                        _context.CenterList.Update(center);

                        bool otherActiveCenters = await _context.CenterList
                            .AnyAsync(c => c.NRDataId == center.NRDataId && c.Id != center.Id && c.Status);

                        if (!otherActiveCenters && nr1 != null)
                        {
                            nr1.Status = false;
                            _context.NrData1.Update(nr1);
                        }
                    }
                    else if (nr1 != null)
                    {
                        nr1.Status = false;
                        _context.NrData1.Update(nr1);

                        var centers = await _context.CenterList
                            .Where(c => c.NRDataId == nr1.Id && c.Status)
                            .ToListAsync();
                        foreach (var c in centers)
                        {
                            c.Status = false;
                            _context.CenterList.Update(c);
                        }
                    }

                    await _context.SaveChangesAsync();
                    await transaction.CommitAsync();
                    return Ok(new { message = "Record deleted successfully" });
                }

                // Field update logic
                var centerProps = typeof(CenterList).GetProperties().ToDictionary(p => p.Name.ToLower(), p => p);
                var nr1Props = typeof(NrData1).GetProperties().ToDictionary(p => p.Name.ToLower(), p => p);

                var extraData = new Dictionary<string, string>();
                var existingExtra = new Dictionary<string, string>();
                if (nr1 != null && !string.IsNullOrWhiteSpace(nr1.NRDatas))
                {
                    try
                    {
                        existingExtra = JsonSerializer.Deserialize<Dictionary<string, string>>(nr1.NRDatas) ?? new();
                    }
                    catch { }
                }

                foreach (var prop in inputData.EnumerateObject())
                {
                    string key = NormalizeKey(prop.Name);
                    string val = prop.Value.ToString();

                    if (key == "id" || key == "projectid") continue;

                    bool handled = false;
                    if (center != null && centerProps.TryGetValue(key, out var cp))
                    {
                        var targetType = Nullable.GetUnderlyingType(cp.PropertyType) ?? cp.PropertyType;
                        object? converted = string.IsNullOrWhiteSpace(val) ? null : Convert.ChangeType(val, targetType);
                        cp.SetValue(center, converted);
                        handled = true;
                    }

                    if (nr1 != null && nr1Props.TryGetValue(key, out var np))
                    {
                        var targetType = Nullable.GetUnderlyingType(np.PropertyType) ?? np.PropertyType;
                        object? converted = string.IsNullOrWhiteSpace(val) ? null : Convert.ChangeType(val, targetType);
                        np.SetValue(nr1, converted);
                        handled = true;
                    }

                    if (!handled)
                    {
                        extraData[prop.Name] = val;
                    }
                }

                if (extraData.Any() && nr1 != null)
                {
                    foreach (var kv in extraData)
                    {
                        existingExtra[kv.Key] = kv.Value;
                    }
                    nr1.NRDatas = JsonSerializer.Serialize(existingExtra);
                }

                if (center != null)
                {
                    _context.CenterList.Update(center);
                }
                if (nr1 != null)
                {
                    _context.NrData1.Update(nr1);
                }

                await _context.SaveChangesAsync();
                await transaction.CommitAsync();

                return Ok(new
                {
                    message = "Data updated successfully",
                    data = (object?)center ?? nr1
                });
            }
            catch (Exception ex)
            {
                await transaction.RollbackAsync();
                await _loggerService.LogErrorAsync("Error in NrData1.UpdateSingle", ex.Message, nameof(NrData1Controller));
                return StatusCode(500, ex.Message);
            }
        }

        // GET: api/NrData1/CheckValueExists/{projectId}
        [HttpGet("CheckValueExists/{projectId}")]
        public async Task<IActionResult> CheckValueExists(int projectId, [FromQuery] string fieldName, [FromQuery] string value, [FromQuery] int? lotNo = null)
        {
            if (string.IsNullOrWhiteSpace(fieldName) || string.IsNullOrWhiteSpace(value))
            {
                return BadRequest("fieldName and value are required.");
            }

            var normalizedValue = value.Trim().ToLowerInvariant();
            var normField = fieldName.Trim().ToLowerInvariant();

            // Center-related fields
            if (normField == "centercode" || normField == "nodalcode" || normField == "route" || normField == "district")
            {
                var centerQuery = _context.CenterList.Where(c => c.ProjectId == projectId && c.Status);
                if (lotNo.HasValue)
                {
                    centerQuery = from c in centerQuery
                                  join nr in _context.NrData1 on c.NRDataId equals nr.Id
                                  where nr.LotNo == lotNo.Value && nr.Status
                                  select c;
                }

                bool exists = normField switch
                {
                    "centercode" => await centerQuery.AnyAsync(c => c.CenterCode != null && c.CenterCode.ToLower() == normalizedValue),
                    "nodalcode" => await centerQuery.AnyAsync(c => c.NodalCode != null && c.NodalCode.ToLower() == normalizedValue),
                    "route" => await centerQuery.AnyAsync(c => c.Route != null && c.Route.ToLower() == normalizedValue),
                    "district" => await centerQuery.AnyAsync(c => c.District != null && c.District.ToLower() == normalizedValue),
                    _ => false
                };
                return Ok(new { exists });
            }

            // Catch-related fields
            var nr1Query = _context.NrData1.Where(n => n.ProjectId == projectId && n.Status);
            if (lotNo.HasValue)
            {
                nr1Query = nr1Query.Where(n => n.LotNo == lotNo.Value);
            }

            bool nr1Exists = normField switch
            {
                "catchno" => await nr1Query.AnyAsync(n => n.CatchNo != null && n.CatchNo.ToLower() == normalizedValue),
                "coursename" => await nr1Query.AnyAsync(n => n.CourseName != null && n.CourseName.ToLower() == normalizedValue),
                "subjectname" => await nr1Query.AnyAsync(n => n.SubjectName != null && n.SubjectName.ToLower() == normalizedValue),
                "examdate" => await nr1Query.AnyAsync(n => n.ExamDate != null && n.ExamDate.ToLower() == normalizedValue),
                "examtime" => await nr1Query.AnyAsync(n => n.ExamTime != null && n.ExamTime.ToLower() == normalizedValue),
                _ => false
            };

            if (nr1Exists)
            {
                return Ok(new { exists = true });
            }

            // Check dynamic fields in NrData1.NRDatas JSON
            var likePattern = $"%\"{fieldName}\":%";
            var jsonRecords = await nr1Query
                .Where(n => n.NRDatas != null && EF.Functions.Like(n.NRDatas, likePattern))
                .Select(n => n.NRDatas)
                .ToListAsync();

            bool jsonExists = false;
            foreach (var jsonStr in jsonRecords)
            {
                try
                {
                    using var doc = JsonDocument.Parse(jsonStr);
                    if (doc.RootElement.TryGetProperty(fieldName, out var propVal))
                    {
                        if (propVal.ToString().Trim().ToLowerInvariant() == normalizedValue)
                        {
                            jsonExists = true;
                            break;
                        }
                    }
                }
                catch { }
            }

            return Ok(new { exists = jsonExists });
        }

        // POST: api/NrData1/single
        [HttpPost("single")]
        public async Task<IActionResult> PostSingleCatch([FromBody] JsonElement inputData, [FromQuery] int? lotNo = null)
        {
            string catchNo = "";
            int effectiveLot = lotNo ?? 0;
            try
            {
                if (inputData.TryGetProperty("data", out var dataArr) && dataArr.GetArrayLength() > 0)
                {
                    var firstItem = dataArr[0];
                    foreach (var prop in firstItem.EnumerateObject())
                    {
                        if (prop.Name.Equals("catchno", StringComparison.OrdinalIgnoreCase))
                        {
                            catchNo = prop.Value.ToString();
                        }
                        if (prop.Name.Equals("lotno", StringComparison.OrdinalIgnoreCase) && int.TryParse(prop.Value.ToString(), out int parsedLot))
                        {
                            effectiveLot = parsedLot;
                        }
                    }
                }
            }
            catch { }

            var res = await PostNRData(inputData, lotNo);
            if (res is OkObjectResult okObj)
            {
                return Ok(new
                {
                    message = "Data added successfully",
                    lotNo = effectiveLot,
                    catchNo = catchNo,
                    result = okObj.Value
                });
            }
            return res;
        }

        // POST: api/NrData1/merge-catchnos
        [HttpPost("merge-catchnos")]
        public async Task<ActionResult> MergeCatchNos(int ProjectId, [FromBody] NRDatasController.MergeCatchNoRequest request, [FromQuery] int? lotNo = null)
        {
            if (request == null)
            {
                return BadRequest("Please select at least 2 catch numbers to merge.");
            }

            var separator = NormalizeText(request.Separator);
            if (separator != "/" && separator != "-")
            {
                return BadRequest("Separator must be '/' or '-'.");
            }

            var requestedCatchNos = (request.CatchNos ?? new List<string>())
                .Select(NormalizeText)
                .Where(v => !string.IsNullOrWhiteSpace(v))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (requestedCatchNos.Count < 2)
            {
                return BadRequest("Please select at least 2 catch numbers to merge.");
            }

            var nr1Rows = await _context.NrData1
                .Where(n => n.ProjectId == ProjectId && n.Status && requestedCatchNos.Contains(n.CatchNo))
                .ToListAsync();

            if (lotNo.HasValue)
            {
                nr1Rows = nr1Rows.Where(n => n.LotNo == lotNo.Value).ToList();
            }

            if (nr1Rows.Count < 2)
            {
                var legacy = new NRDatasController(_context, _loggerService, _apiSettings, _dispatchService)
                {
                    ControllerContext = this.ControllerContext
                };
                return await legacy.MergeCatchNos(ProjectId, request, lotNo);
            }

            var primaryNr1 = nr1Rows.First();
            var secondaryNr1List = nr1Rows.Skip(1).ToList();
            var mergedCatchNo = string.Join(separator, requestedCatchNos);

            primaryNr1.CatchNo = mergedCatchNo;
            if (!string.IsNullOrWhiteSpace(request.ExamDate)) primaryNr1.ExamDate = request.ExamDate;
            if (!string.IsNullOrWhiteSpace(request.ExamTime)) primaryNr1.ExamTime = request.ExamTime;

            var primaryCenters = await _context.CenterList
                .Where(c => c.NRDataId == primaryNr1.Id && c.Status)
                .ToListAsync();

            var primaryCenterMap = new Dictionary<string, CenterList>(StringComparer.OrdinalIgnoreCase);
            foreach (var c in primaryCenters)
            {
                var k = (c.CenterCode ?? "").Trim().ToLowerInvariant();
                if (!primaryCenterMap.ContainsKey(k))
                {
                    primaryCenterMap[k] = c;
                }
            }

            foreach (var sec in secondaryNr1List)
            {
                sec.Status = false;
                var secCenters = await _context.CenterList
                    .Where(c => c.NRDataId == sec.Id && c.Status)
                    .ToListAsync();

                foreach (var sc in secCenters)
                {
                    var cKey = (sc.CenterCode ?? "").Trim().ToLowerInvariant();
                    if (primaryCenterMap.TryGetValue(cKey, out var existingCenter))
                    {
                        existingCenter.Quantity += sc.Quantity;
                        existingCenter.NRQuantity += sc.NRQuantity;
                        sc.Status = false;
                    }
                    else
                    {
                        sc.NRDataId = primaryNr1.Id;
                        primaryCenterMap[cKey] = sc;
                    }
                }
            }

            try
            {
                var legacy = new NRDatasController(_context, _loggerService, _apiSettings, _dispatchService)
                {
                    ControllerContext = this.ControllerContext
                };
                await legacy.MergeCatchNos(ProjectId, request, lotNo);
            }
            catch { }

            await _context.SaveChangesAsync();
            return Ok(new { message = $"Catch numbers merged successfully into {mergedCatchNo}." });
        }

        // GET: api/NrData1/ErrorReport
        [HttpGet("ErrorReport")]
        public async Task<ActionResult> ErrorReport([FromQuery] int ProjectId, [FromQuery] int? lotNo = null)
        {
            var legacy = new NRDatasController(_context, _loggerService, _apiSettings, _dispatchService)
            {
                ControllerContext = this.ControllerContext
            };
            return await legacy.GetDuplicateswrtCatch(ProjectId, lotNo);
        }

        // PUT: api/NrData1
        [HttpPut]
        public async Task<ActionResult> ResolveConflicts([FromQuery] int ProjectId, [FromBody] NRDatasController.ConflictResolutionDto payload, [FromQuery] int? lotNo = null)
        {
            var legacy = new NRDatasController(_context, _loggerService, _apiSettings, _dispatchService)
            {
                ControllerContext = this.ControllerContext
            };
            return await legacy.ResolveConflicts(ProjectId, payload, lotNo);
        }

        // PUT: api/NrData1/conflicts/status
        [HttpPut("conflicts/status")]
        public async Task<ActionResult> UpdateConflictStatus([FromQuery] int ProjectId, [FromBody] NRDatasController.ConflictStatusUpdateDto payload, [FromQuery] int? lotNo = null)
        {
            var legacy = new NRDatasController(_context, _loggerService, _apiSettings, _dispatchService)
            {
                ControllerContext = this.ControllerContext
            };
            return await legacy.UpdateConflictStatus(ProjectId, payload, lotNo);
        }

        // GET: api/NrData1/centers/5
        [HttpGet("centers/{nrDataId}")]
        public async Task<ActionResult<IEnumerable<CenterList>>> GetCentersByNrDataId(int nrDataId)
        {
            var centers = await _context.CenterList
                .Where(c => c.NRDataId == nrDataId && c.Status)
                .ToListAsync();

            return Ok(centers);
        }

        // GET: api/NrData1/center-debug/5
        [HttpGet("center-debug/{projectId}")]
        public async Task<ActionResult> GetCenterDebug(int projectId)
        {
            var centers = await _context.CenterList
                .Where(c => c.ProjectId == projectId && c.Status)
                .ToListAsync();
            return Ok(centers);
        }


        // POST: api/NrData1
        [HttpPost]
        [HttpPost("PostNRData")]
        public async Task<IActionResult> PostNRData([FromBody] JsonElement inputData, [FromQuery] int? lotNo = null)
        {
            if (!ModelState.IsValid)
                return BadRequest(ModelState);

            try
            {
                int projectId = inputData.GetProperty("projectId").GetInt32();
                var incomingData = inputData.GetProperty("data");

                bool isChangedNR = false;
                if (inputData.TryGetProperty("isChangedNR", out JsonElement isChangedNRElem) && isChangedNRElem.ValueKind == JsonValueKind.True)
                {
                    isChangedNR = true;
                }

                // 1. Fetch project configs and fields
                var projectConfig = await _context.ProjectConfigs
                    .FirstOrDefaultAsync(x => x.ProjectId == projectId);

                var extraConfigs = await _context.ExtraConfigurations
                    .Where(x => x.ProjectId == projectId)
                    .ToListAsync();

                var allFields = await _context.Fields
                    .AsNoTracking()
                    .ToListAsync();

                // Unique fields from Fields table (where IsUnique is true)
                var uniqueFieldsFromDb = allFields
                    .Where(f => f.IsUnique)
                    .Select(f => f.Name.Trim())
                    .ToList();

                var uniqueFieldSet = new HashSet<string>(uniqueFieldsFromDb, StringComparer.OrdinalIgnoreCase)
                {
                    "CatchNo" // CatchNo is always part of unique criteria
                };

                // Dynamic matching fields if duplicate criteria is configured
                var duplicateFieldIds = projectConfig?.DuplicateCriteria ?? new List<int>();
                var duplicateCriteriaFields = allFields
                    .Where(f => duplicateFieldIds.Contains(f.FieldId))
                    .Select(f => f.Name.Trim())
                    .ToList();

                // 2. Reflection Property Maps
                var nrData1Properties = typeof(NrData1)
                    .GetProperties()
                    .ToDictionary(p => p.Name.ToLower(), p => p);

                var centerListProperties = typeof(CenterList)
                    .GetProperties()
                    .ToDictionary(p => p.Name.ToLower(), p => p);

                // Center-specific property names that should never be mapped to NrData1
                var centerSpecificNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                {
                    "centercode", "nodalcode", "route", "routesort", "centersort",
                    "nodalsort", "district", "districtsort", "centerdata", "status",
                    "centername", "collegename", "collegecode"
                };

                // 3. Batch determination
                var existingNrData1List = await _context.NrData1
                    .Where(x => x.ProjectId == projectId)
                    .ToListAsync();

                int batchId = 1;
                if (existingNrData1List.Any() && isChangedNR)
                {
                    batchId = existingNrData1List.Max(x => x.Batch) + 1;
                }

                // 4. Parse incoming rows and group by Catch / Unique fields
                // We keep a data structure per unique catch:
                // CatchGroup: NrData1 prototype, accumulated Quantity, accumulated NRQuantity, and list of CenterList records
                var catchGroups = new Dictionary<string, CatchGroupData>(StringComparer.OrdinalIgnoreCase);
                var extraEnvelopesToAdd = new List<ExtraEnvelopes>();

                // Fetch all ExtraTypes from database to avoid hardcoding names
                var allExtraTypes = await _context.ExtraType.ToListAsync();
                var extraTypesByName = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                foreach (var et in allExtraTypes)
                {
                    if (!string.IsNullOrWhiteSpace(et.Type))
                    {
                        extraTypesByName[et.Type.Trim()] = et.ExtraTypeId;
                        if (!et.Type.EndsWith("Extra", StringComparison.OrdinalIgnoreCase))
                        {
                            extraTypesByName[$"{et.Type.Trim()} Extra"] = et.ExtraTypeId;
                        }
                    }
                }

                int totalItems = incomingData.GetArrayLength();

                for (int i = 0; i < totalItems; i++)
                {
                    var item = incomingData[i];

                    // Extract catch number and center code
                    string itemCatchNo = "";
                    string itemCenterCode = "";
                    int itemQty = 0;
                    int itemNrQty = 0;

                    var uniqueExtraData = new Dictionary<string, string>();
                    var centerExtraData = new Dictionary<string, string>();

                    // Temporary dictionaries for mapping properties
                    var nrData1FieldValues = new Dictionary<PropertyInfo, object?>();
                    var centerListFieldValues = new Dictionary<PropertyInfo, object?>();

                    foreach (var prop in item.EnumerateObject())
                    {
                        string rawKey = prop.Name;
                        string key = rawKey.Replace(" ", "").ToLower();
                        string value = prop.Value.ToString().Trim();

                        if (key == "catchno") itemCatchNo = value;
                        if (key == "centercode") itemCenterCode = value;
                        if (key == "quantity" && int.TryParse(value, out int qVal)) itemQty = qVal;
                        if (key == "nrquantity" && int.TryParse(value, out int nrqVal)) itemNrQty = nrqVal;

                        // Check if this property is a UNIQUE field:
                        // 1. In uniqueFieldSet (IsUnique == true in Fields table)
                        // 2. OR is a typed property on NrData1 (and NOT center-specific)
                        bool isFieldUnique = uniqueFieldSet.Contains(rawKey) ||
                                             uniqueFieldSet.Contains(key) ||
                                             (nrData1Properties.ContainsKey(key) && !centerSpecificNames.Contains(key));

                        if (isFieldUnique)
                        {
                            // Unique field -> belongs to NrData1
                            if (nrData1Properties.TryGetValue(key, out var propInfo))
                            {
                                if (propInfo.Name.Equals("Id", StringComparison.OrdinalIgnoreCase) ||
                                    propInfo.Name.Equals("ProjectId", StringComparison.OrdinalIgnoreCase))
                                {
                                    continue;
                                }

                                var converted = ConvertPropertyValue(propInfo, value);
                                if (converted != null)
                                {
                                    nrData1FieldValues[propInfo] = converted;
                                }
                            }
                            else if (rawKey != "projectId" && rawKey != "isCorrectedNrdataReport")
                            {
                                uniqueExtraData[rawKey] = value;
                            }
                        }
                        else
                        {
                            // Non-unique field -> belongs to CenterList ("rest details")
                            if (centerListProperties.TryGetValue(key, out var propInfo))
                            {
                                if (propInfo.Name.Equals("Id", StringComparison.OrdinalIgnoreCase) ||
                                    propInfo.Name.Equals("ProjectId", StringComparison.OrdinalIgnoreCase) ||
                                    propInfo.Name.Equals("NRDataId", StringComparison.OrdinalIgnoreCase))
                                {
                                    continue;
                                }

                                var converted = ConvertPropertyValue(propInfo, value);
                                if (converted != null)
                                {
                                    centerListFieldValues[propInfo] = converted;
                                }
                            }
                            else if (rawKey != "projectId" && rawKey != "isCorrectedNrdataReport")
                            {
                                centerExtraData[rawKey] = value;
                            }
                        }
                    }

                    // Ensure Quantity and NRQuantity for this center are preserved in CenterData
                    if (!centerExtraData.ContainsKey("Quantity") && itemQty > 0)
                        centerExtraData["Quantity"] = itemQty.ToString();
                    if (!centerExtraData.ContainsKey("NRQuantity") && itemNrQty > 0)
                        centerExtraData["NRQuantity"] = itemNrQty.ToString();

                    // Determine unique grouping key (CatchNo, or composite matching fields)
                    string catchKey = !string.IsNullOrWhiteSpace(itemCatchNo) ? itemCatchNo.Trim() : $"ROW_{i}";

                    if (!catchGroups.TryGetValue(catchKey, out var catchGroup))
                    {
                        var nrData1 = new NrData1
                        {
                            ProjectId = projectId,
                            CatchNo = itemCatchNo,
                            Batch = batchId,
                            Steps = PipelineNavigator.STEP_UPLOADED,
                          
                        };

                        if (lotNo.HasValue)
                        {
                            nrData1.EnvLotNo = lotNo.Value;
                        }

                        // Apply typed unique properties
                        foreach (var kvp in nrData1FieldValues)
                        {
                            if (kvp.Key.CanWrite && kvp.Key.Name != "Quantity" && kvp.Key.Name != "NRQuantity")
                            {
                                kvp.Key.SetValue(nrData1, kvp.Value);
                            }
                        }

                        // Calculate Day from ExamDate
                        if (!string.IsNullOrWhiteSpace(nrData1.ExamDate))
                        {
                            if (DateTime.TryParseExact(
                                nrData1.ExamDate.Trim(),
                                "dd-MM-yyyy",
                                CultureInfo.InvariantCulture,
                                DateTimeStyles.None,
                                out DateTime examDate))
                            {
                                nrData1.Day = examDate.ToString("dddd", CultureInfo.InvariantCulture);
                            }
                            else if (DateTime.TryParse(nrData1.ExamDate, out DateTime fallbackDate))
                            {
                                nrData1.Day = fallbackDate.DayOfWeek.ToString();
                            }
                        }

                        catchGroup = new CatchGroupData
                        {
                            NrData1 = nrData1,
                            UniqueExtraData = new Dictionary<string, string>(uniqueExtraData, StringComparer.OrdinalIgnoreCase)
                        };
                        catchGroups[catchKey] = catchGroup;
                    }
                    else
                    {
                        // Merge any additional unique dynamic data
                        foreach (var kvp in uniqueExtraData)
                        {
                            if (!catchGroup.UniqueExtraData.ContainsKey(kvp.Key))
                            {
                                catchGroup.UniqueExtraData[kvp.Key] = kvp.Value;
                            }
                        }
                    }
               
                    // Build CenterList record
                    string itemRemark = "";
                    if (centerExtraData.TryGetValue("Remark", out var remVal)) itemRemark = remVal;
                    else if (uniqueExtraData.TryGetValue("Remark", out var uremVal)) itemRemark = uremVal;

                    bool isReserveRemark = !string.IsNullOrWhiteSpace(itemRemark) &&
                        (itemRemark.Contains("Reserve", StringComparison.OrdinalIgnoreCase) ||
                         itemRemark.Contains("Extra", StringComparison.OrdinalIgnoreCase));

                    bool isExtraRow = (!string.IsNullOrWhiteSpace(itemCenterCode) && extraTypesByName.ContainsKey(itemCenterCode.Trim())) || isReserveRemark;
                    int? extraTypeId = isReserveRemark ? 1 : (isExtraRow ? extraTypesByName[itemCenterCode.Trim()] : null);

                    var centerList = new CenterList
                    {
                        ProjectId = projectId,
                        CenterCode = itemCenterCode,
                        Status = true,
                        ExtraId = extraTypeId ?? 0,
                        CenterData = centerExtraData.Any() ? JsonSerializer.Serialize(centerExtraData) : "{}"
                    };

                    // Apply typed center properties
                    foreach (var kvp in centerListFieldValues)
                    {
                        if (kvp.Key.CanWrite)
                        {
                            kvp.Key.SetValue(centerList, kvp.Value);
                        }
                    }

                    if (isReserveRemark && centerList.Quantity <= 0 && itemQty > 0)
                    {
                        centerList.Quantity = itemQty;
                        centerList.NRQuantity = itemQty;
                    }

                    // Reserve centers are associated with a specific center code and must be saved in catchGroup.Centers
                    if (!isExtraRow || isReserveRemark)
                    {
                        catchGroup.Centers.Add(centerList);
                    }

                    // Process Extra Envelopes
                    var extraEnvelopes = ProcessExtraEnvelopes(
                        itemCatchNo,
                        itemCenterCode,
                        itemQty,
                        itemNrQty,
                        projectId,
                        extraConfigs,
                        extraTypesByName
                    );
                    extraEnvelopesToAdd.AddRange(extraEnvelopes);
                }

                // 5. Finalize NrData1 dynamic JSON serialization
                var nrData1List = new List<NrData1>();
                var allCenterListToAdd = new List<CenterList>();

                foreach (var group in catchGroups.Values)
                {
                    if (group.UniqueExtraData.Any())
                    {
                        group.NrData1.NRDatas = JsonSerializer.Serialize(group.UniqueExtraData);
                    }
                    nrData1List.Add(group.NrData1);
                }

                // 6. If re-uploading (Batch > 1), deactivate existing active center records for catches being updated
                if (batchId > 1 && existingNrData1List.Any())
                {
                    var incomingCatchNos = nrData1List
                        .Select(x => x.CatchNo)
                        .Where(x => !string.IsNullOrWhiteSpace(x))
                        .ToHashSet(StringComparer.OrdinalIgnoreCase);

                    var previousNrData1Rows = existingNrData1List
                        .Where(x => x.CatchNo != null && incomingCatchNos.Contains(x.CatchNo))
                        .ToList();

                    var previousNrDataIds = previousNrData1Rows.Select(x => x.Id).ToList();

                    if (previousNrDataIds.Any())
                    {
                        foreach (var oldNr in previousNrData1Rows)
                        {
                            oldNr.Status = false;
                            _context.NrData1.Update(oldNr);
                        }

                        var oldCenters = await _context.CenterList
                            .Where(c => c.ProjectId == projectId && previousNrDataIds.Contains(c.NRDataId) && c.Status)
                            .ToListAsync();

                        foreach (var oldCenter in oldCenters)
                        {
                            oldCenter.Status = false;
                            _context.CenterList.Update(oldCenter);
                        }
                    }
                }

                // 7. Save NrData1 records first to generate their Identity IDs
                if (nrData1List.Any())
                {
                    await _context.NrData1.AddRangeAsync(nrData1List);
                    await _context.SaveChangesAsync();
                }

                // 8. Associate NRDataId with CenterList records and save them
                foreach (var group in catchGroups.Values)
                {
                    int generatedNrDataId = group.NrData1.Id;
                    foreach (var center in group.Centers)
                    {
                        center.NRDataId = generatedNrDataId;
                        allCenterListToAdd.Add(center);
                    }
                }

                if (allCenterListToAdd.Any())
                {
                    await _context.CenterList.AddRangeAsync(allCenterListToAdd);
                }

                // 9. Process and save Extra Envelopes
                var groupedExtras = GroupAndProcessExtras(
                    extraEnvelopesToAdd,
                    projectId,
                    extraConfigs
                );

                if (groupedExtras.Any())
                {
                    // Deactivate overlapping existing extras
                    var existingExtras = await _context.ExtrasEnvelope
                        .Where(e => e.ProjectId == projectId && e.Status == 1)
                        .ToListAsync();

                    foreach (var ex in existingExtras)
                    {
                        if (groupedExtras.Any(ge => ge.ExtraId == ex.ExtraId && ge.CatchNo == ex.CatchNo))
                        {
                            ex.Status = 0;
                            _context.ExtrasEnvelope.Update(ex);
                        }
                    }

                    await _context.ExtrasEnvelope.AddRangeAsync(groupedExtras);
                }

                await _context.SaveChangesAsync();

                // 10. Logging
                await _loggerService.LogEventAsync(
                    batchId == 1
                        ? $"Initial Upload Processed (Batch 1) into nrdata1 and centerlist"
                        : $"Upload Processed (Batch {batchId}) into nrdata1 and centerlist",
                    "NrData1",
                    LogHelper.GetTriggeredBy(User),
                    projectId
                );

                return Ok(new
                {
                    message = batchId == 1 ? "Initial data uploaded successfully" : "Data processed successfully",
                    Batch = batchId,
                    NewCatches = nrData1List.Count,
                    NewCenters = allCenterListToAdd.Count
                });
            }
            catch (Exception ex)
            {
                await _loggerService.LogErrorAsync(
                    "NrData1 Upload Error",
                    ex.ToString(),
                    nameof(NrData1Controller)
                );

                return StatusCode(500, ex.Message);
            }
        }

        private static string NormalizeKey(string? s)
        {
            if (string.IsNullOrWhiteSpace(s)) return "";
            return s.Replace(" ", "").Replace("_", "").ToLowerInvariant().Trim();
        }

        private static string GetFieldValue(CenterList c, NrData1? p, string fieldName)
        {
            var target = NormalizeKey(fieldName);

            // 1. Check CenterList direct properties
            var cProp = typeof(CenterList).GetProperties(BindingFlags.Public | BindingFlags.Instance)
                .FirstOrDefault(pr => NormalizeKey(pr.Name) == target);
            if (cProp != null)
            {
                var val = cProp.GetValue(c)?.ToString()?.Trim();
                if (!string.IsNullOrEmpty(val)) return val;
            }

            // 2. Check CenterList.CenterData JSON
            if (!string.IsNullOrWhiteSpace(c.CenterData))
            {
                try
                {
                    var dict = JsonSerializer.Deserialize<Dictionary<string, string>>(c.CenterData);
                    if (dict != null)
                    {
                        var kvp = dict.FirstOrDefault(k => NormalizeKey(k.Key) == target);
                        if (!string.IsNullOrEmpty(kvp.Value)) return kvp.Value.Trim();
                    }
                }
                catch { }
            }

            // 3. Check NrData1 direct properties
            if (p != null)
            {
                var pProp = typeof(NrData1).GetProperties(BindingFlags.Public | BindingFlags.Instance)
                    .FirstOrDefault(pr => NormalizeKey(pr.Name) == target);
                if (pProp != null)
                {
                    var val = pProp.GetValue(p)?.ToString()?.Trim();
                    if (!string.IsNullOrEmpty(val)) return val;
                }

                // 4. Check NrData1.NRDatas JSON
                if (!string.IsNullOrWhiteSpace(p.NRDatas))
                {
                    try
                    {
                        var dict = JsonSerializer.Deserialize<Dictionary<string, string>>(p.NRDatas);
                        if (dict != null)
                        {
                            var kvp = dict.FirstOrDefault(k => NormalizeKey(k.Key) == target);
                            if (!string.IsNullOrEmpty(kvp.Value)) return kvp.Value.Trim();
                        }
                    }
                    catch { }
                }
            }

            return "";
        }

        [HttpPost("MergeFields")]
        public async Task<IActionResult> MergeFields([FromQuery] int ProjectId, [FromQuery] int? batchId = null, [FromQuery] int? lotNo = null)
        {
            try
            {
                Console.WriteLine($"[NrData1Controller] MergeFields received ProjectId: {ProjectId}, batchId: {batchId}, lotNo: {lotNo}");

                IQueryable<NrData1> query = _context.NrData1
                    .Where(p => p.ProjectId == ProjectId && p.Status);

                // 1. Try with batchId and lotNo if provided, prioritizing STEP_UPLOADED (0)
                var attemptQuery = query;
                if (batchId.HasValue && batchId.Value > 0)
                {
                    Console.WriteLine($"[NrData1Controller] Filtering by batchId: {batchId.Value}");
                    attemptQuery = attemptQuery.Where(p => p.Batch == batchId.Value);
                }
                if (lotNo.HasValue && lotNo.Value > 0)
                {
                    Console.WriteLine($"[NrData1Controller] Filtering by lotNo: {lotNo.Value}");
                    attemptQuery = attemptQuery.Where(p => p.LotNo == lotNo.Value || p.EnvLotNo == lotNo.Value);
                }

                var data = await attemptQuery.Where(p => p.Steps == PipelineNavigator.STEP_UPLOADED).ToListAsync();

                // 2. If no records found with lot filter, try with batchId + STEP_UPLOADED (records might have LotNo = 0)
                if (!data.Any() && lotNo.HasValue && lotNo.Value > 0)
                {
                    var batchQuery = query;
                    if (batchId.HasValue && batchId.Value > 0)
                    {
                        batchQuery = batchQuery.Where(p => p.Batch == batchId.Value);
                    }
                    data = await batchQuery.Where(p => p.Steps == PipelineNavigator.STEP_UPLOADED).ToListAsync();
                }

                // 3. If still no records found, find ANY records with STEP_UPLOADED for this ProjectId
                if (!data.Any())
                {
                    data = await query.Where(p => p.Steps == PipelineNavigator.STEP_UPLOADED).ToListAsync();
                }

                // 4. Fallback for re-running on specific batch if no STEP_UPLOADED found
                if (!data.Any() && batchId.HasValue && batchId.Value > 0)
                {
                    data = await attemptQuery.ToListAsync();
                    if (!data.Any())
                    {
                        data = await query.Where(p => p.Batch == batchId.Value).ToListAsync();
                    }
                }

                // 5. Fallback for re-running across project
                if (!data.Any())
                {
                    data = await query.ToListAsync();
                }

                Console.WriteLine($"[NrData1Controller] Found {data.Count} records to process");

                if (!data.Any())
                    return BadRequest("No data found to process duplication. All data is already processed or no new data added.");

                var hasPendingDuplicateRows = data.Any(d => d.Steps == PipelineNavigator.STEP_UPLOADED);
                var projectconfig = await _context.ProjectConfigs
                    .FirstOrDefaultAsync(p => p.ProjectId == ProjectId);
                if (projectconfig == null)
                    return NotFound("Project config not exists for this project");

                var mergeFieldIds = projectconfig.DuplicateCriteria ?? new List<int>();
                if (!mergeFieldIds.Any())
                    return BadRequest("Duplicate criteria is not configured for this project.");

                var allFields = await _context.Fields.AsNoTracking().ToListAsync();
                var fieldNames = allFields
                    .Where(f => mergeFieldIds.Contains(f.FieldId))
                    .Select(f => f.Name.Trim())
                    .ToList();

                if (!fieldNames.Any())
                    return BadRequest("Duplicate criteria fields not found.");

                // Map of all catches for this project
                var nrDataIds = data.Select(d => d.Id).ToList();
                var nrDataMap = await _context.NrData1
                    .Where(d => d.ProjectId == ProjectId)
                    .ToDictionaryAsync(d => d.Id);

                // Fetch active regular CenterList records (ExtraId == 0) for these catches
                var activeCenters = await _context.CenterList
                    .Where(c => c.ProjectId == ProjectId && nrDataIds.Contains(c.NRDataId) && c.Status && c.ExtraId == 0)
                    .ToListAsync();

                if (!activeCenters.Any())
                {
                    return BadRequest("No active center list data found to process duplication.");
                }

                // Group active centers by duplicate criteria
                var centerGroups = activeCenters.GroupBy(c =>
                {
                    var p = nrDataMap.TryGetValue(c.NRDataId, out var nr) ? nr : null;
                    var keyParts = fieldNames.Select(fn => NormalizeKey(GetFieldValue(c, p, fn)));
                    return string.Join("|", keyParts);
                }).ToList();

                int mergedCount = 0;
                var deletedCenters = new List<CenterList>();
                var newCenters = new List<CenterList>();
                var affectedNrDataIds = new HashSet<int>();

                foreach (var cg in centerGroups)
                {
                    if (cg.Count() <= 1)
                        continue;

                    var firstCenter = cg.First();
                    int totalNrQty = cg.Sum(c => c.NRQuantity);
                    int finalQty = Math.Max(totalNrQty, cg.Max(c => c.Quantity));

                    // Merge dynamic CenterData JSON
                    var mergedCenterData = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    var remarksList = new List<string>();
                    var importRowNos = new List<string>();

                    foreach (var c in cg)
                    {
                        if (!string.IsNullOrWhiteSpace(c.CenterData))
                        {
                            try
                            {
                                var dict = JsonSerializer.Deserialize<Dictionary<string, string>>(c.CenterData);
                                if (dict != null)
                                {
                                    foreach (var kvp in dict)
                                    {
                                        if (string.Equals(kvp.Key, "Quantity", StringComparison.OrdinalIgnoreCase) ||
                                            string.Equals(kvp.Key, "NRQuantity", StringComparison.OrdinalIgnoreCase))
                                            continue;

                                        if (string.Equals(kvp.Key, "Remark", StringComparison.OrdinalIgnoreCase) ||
                                            string.Equals(kvp.Key, "Remarks", StringComparison.OrdinalIgnoreCase))
                                        {
                                            if (!string.IsNullOrWhiteSpace(kvp.Value))
                                                remarksList.Add(kvp.Value.Trim());
                                            continue;
                                        }

                                        if (string.Equals(kvp.Key, "ImportRowNo", StringComparison.OrdinalIgnoreCase))
                                        {
                                            if (!string.IsNullOrWhiteSpace(kvp.Value))
                                                importRowNos.Add(kvp.Value.Trim());
                                            continue;
                                        }

                                        if (!mergedCenterData.ContainsKey(kvp.Key) || string.IsNullOrWhiteSpace(mergedCenterData[kvp.Key]))
                                        {
                                            mergedCenterData[kvp.Key] = kvp.Value;
                                        }
                                    }
                                }
                            }
                            catch { }
                        }
                    }

                    if (remarksList.Any())
                        mergedCenterData["Remark"] = string.Join(" / ", remarksList.Distinct(StringComparer.OrdinalIgnoreCase));
                    if (importRowNos.Any())
                        mergedCenterData["ImportRowNo"] = string.Join(" / ", importRowNos.Distinct(StringComparer.OrdinalIgnoreCase));
                    mergedCenterData["Quantity"] = finalQty.ToString();
                    mergedCenterData["NRQuantity"] = totalNrQty.ToString();

                    int targetNrDataId = firstCenter.NRDataId;

                    // Create new merged CenterList record
                    var newMergedCenter = new CenterList
                    {
                        ProjectId = ProjectId,
                        NRDataId = targetNrDataId,
                        CenterCode = firstCenter.CenterCode,
                        NodalCode = firstCenter.NodalCode,
                        Route = firstCenter.Route,
                        RouteSort = firstCenter.RouteSort,
                        CenterSort = firstCenter.CenterSort,
                        NodalSort = firstCenter.NodalSort,
                        District = firstCenter.District,
                        DistrictSort = firstCenter.DistrictSort,
                        ExtraId = 0,
                        Quantity = finalQty,
                        NRQuantity = totalNrQty,
                        Status = true,
                        CenterData = JsonSerializer.Serialize(mergedCenterData)
                    };

                    await _context.CenterList.AddAsync(newMergedCenter);
                    newCenters.Add(newMergedCenter);

                    // Deactivate old duplicate centers
                    foreach (var c in cg)
                    {
                        c.Status = false;
                        _context.CenterList.Update(c);
                        deletedCenters.Add(c);
                        affectedNrDataIds.Add(c.NRDataId);
                    }

                    // Check if duplicate centers belonged to multiple distinct catches
                    var distinctCatchIds = cg.Select(c => c.NRDataId).Distinct().ToList();
                    if (distinctCatchIds.Count > 1)
                    {
                        var primaryCatchId = targetNrDataId;
                        var secondaryCatchIds = distinctCatchIds.Where(id => id != primaryCatchId).ToList();

                        if (nrDataMap.TryGetValue(primaryCatchId, out var primaryCatch))
                        {
                            var secondaryCatches = secondaryCatchIds
                                .Where(id => nrDataMap.ContainsKey(id))
                                .Select(id => nrDataMap[id])
                                .ToList();

                            var allSubjects = new[] { primaryCatch.SubjectName }
                                .Concat(secondaryCatches.Select(s => s.SubjectName))
                                .Where(s => !string.IsNullOrWhiteSpace(s))
                                .Distinct(StringComparer.OrdinalIgnoreCase)
                                .ToList();
                            if (allSubjects.Count > 1)
                                primaryCatch.SubjectName = string.Join(" / ", allSubjects);

                            var allCourses = new[] { primaryCatch.CourseName }
                                .Concat(secondaryCatches.Select(s => s.CourseName))
                                .Where(s => !string.IsNullOrWhiteSpace(s))
                                .Distinct(StringComparer.OrdinalIgnoreCase)
                                .ToList();
                            if (allCourses.Count > 1)
                                primaryCatch.CourseName = string.Join(" / ", allCourses);

                            _context.NrData1.Update(primaryCatch);

                            // Re-point remaining active centers to primary catch
                            var otherCenters = await _context.CenterList
                                .Where(c => c.ProjectId == ProjectId && secondaryCatchIds.Contains(c.NRDataId) && c.Status)
                                .ToListAsync();

                            foreach (var oc in otherCenters)
                            {
                                if (!deletedCenters.Any(d => d.Id == oc.Id))
                                {
                                    oc.NRDataId = primaryCatchId;
                                    _context.CenterList.Update(oc);
                                }
                            }

                            // Remove secondary catches from NrData1
                            foreach (var sc in secondaryCatches)
                            {
                                _context.NrData1.Remove(sc);
                            }
                        }
                    }

                    mergedCount += cg.Count() - 1;
                }

                // Advance steps to STEP_DUP_PARTIAL for data catches
                foreach (var row in data)
                {
                    if (_context.Entry(row).State == EntityState.Detached || _context.Entry(row).State == EntityState.Deleted)
                        continue;

                    row.Steps = PipelineNavigator.STEP_DUP_PARTIAL;
                    _context.NrData1.Update(row);
                }

                await _context.SaveChangesAsync();

                // Recalculate catch dynamic quantities for affected catches
                foreach (var catchId in affectedNrDataIds)
                {
                    if (nrDataMap.TryGetValue(catchId, out var catchRow) &&
                        _context.Entry(catchRow).State != EntityState.Deleted &&
                        _context.Entry(catchRow).State != EntityState.Detached)
                    {
                        var curCenters = await _context.CenterList
                            .Where(c => c.ProjectId == ProjectId && c.NRDataId == catchId && c.Status && c.ExtraId == 0)
                            .ToListAsync();

                        if (curCenters.Any())
                        {
                            var dyn = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                            if (!string.IsNullOrWhiteSpace(catchRow.NRDatas))
                            {
                                try
                                {
                                    dyn = JsonSerializer.Deserialize<Dictionary<string, string>>(catchRow.NRDatas)
                                        ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                                }
                                catch { }
                            }

                            dyn["Quantity"] = curCenters.Sum(c => c.Quantity).ToString();
                            dyn["NRQuantity"] = curCenters.Sum(c => c.NRQuantity).ToString();
                            catchRow.NRDatas = JsonSerializer.Serialize(dyn);
                            _context.NrData1.Update(catchRow);
                        }
                    }
                }

                await _context.SaveChangesAsync();

                var triggeredBy = LogHelper.GetTriggeredBy(User);

                await _loggerService.LogEventAsync(
                    "Duplicates merged and new rows created (CenterList / NrData1)",
                    "Duplicates",
                    triggeredBy,
                    ProjectId,
                    string.Empty,
                    LogHelper.ToJson(new { ProjectId, MergedCount = mergedCount })
                );

                // ================= REPORT =================
                var reportPath = FileStorageHelper.GetProjectFolder(ProjectId);

                var distinctLots = (lotNo.HasValue && lotNo.Value > 0)
                    ? new List<int> { lotNo.Value }
                    : data.Where(r => r.EnvLotNo > 0).Select(r => r.EnvLotNo).Distinct().OrderBy(l => l).ToList();
                var lotStr = distinctLots.Any() ? string.Join("_", distinctLots) : "All";

                var fileName = ReportVersionHelper.GetNextVersionFileName(reportPath, $"DuplicateTool_NrData1_{lotStr}.xlsx");
                var filePath = Path.Combine(reportPath, fileName);

                // Build report rows: include all centers involved in duplicate groups + new merged centers
                var reportCenters = new List<CenterList>();
                reportCenters.AddRange(deletedCenters);
                reportCenters.AddRange(newCenters);

                // If no duplicates found, show active centers
                if (!reportCenters.Any())
                {
                    reportCenters.AddRange(activeCenters);
                }

                // Collect extra dynamic headers
                var extraHeaders = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var c in reportCenters)
                {
                    if (!string.IsNullOrWhiteSpace(c.CenterData))
                    {
                        try
                        {
                            var dict = JsonSerializer.Deserialize<Dictionary<string, string>>(c.CenterData);
                            if (dict != null)
                            {
                                foreach (var key in dict.Keys)
                                {
                                    if (!key.Equals("Quantity", StringComparison.OrdinalIgnoreCase) &&
                                        !key.Equals("NRQuantity", StringComparison.OrdinalIgnoreCase))
                                    {
                                        extraHeaders.Add(key);
                                    }
                                }
                            }
                        }
                        catch { }
                    }
                }

                var extraHeaderList = extraHeaders.OrderBy(x => x).ToList();

                ExcelPackage.License.SetNonCommercialPersonal("Tools");
                using (var package = new ExcelPackage())
                {
                    // Sheet 1: Merge Report
                    var ws = package.Workbook.Worksheets.Add("Merge Report");

                    var baseHeaders = new List<string>
                    {
                        "CatchNo", "SubjectName", "CourseName", "NodalCode", "CenterCode",
                        "Route", "Quantity", "NRQuantity", "Status"
                    };

                    int col = 1;
                    foreach (var h in baseHeaders)
                    {
                        ws.Cells[1, col].Value = h;
                        ws.Cells[1, col].Style.Font.Bold = true;
                        col++;
                    }

                    foreach (var key in extraHeaderList)
                    {
                        ws.Cells[1, col].Value = key;
                        ws.Cells[1, col].Style.Font.Bold = true;
                        col++;
                    }

                    int rowIdx = 2;
                    foreach (var item in reportCenters)
                    {
                        nrDataMap.TryGetValue(item.NRDataId, out var catchInfo);

                        col = 1;
                        ws.Cells[rowIdx, col++].Value = catchInfo?.CatchNo ?? "";
                        ws.Cells[rowIdx, col++].Value = catchInfo?.SubjectName ?? "";
                        ws.Cells[rowIdx, col++].Value = catchInfo?.CourseName ?? "";
                        ws.Cells[rowIdx, col++].Value = item.NodalCode ?? "";
                        ws.Cells[rowIdx, col++].Value = item.CenterCode ?? "";
                        ws.Cells[rowIdx, col++].Value = item.Route ?? "";
                        ws.Cells[rowIdx, col++].Value = item.Quantity;
                        ws.Cells[rowIdx, col++].Value = item.NRQuantity;
                        ws.Cells[rowIdx, col++].Value = item.Status ? "Active" : "Inactive";

                        Dictionary<string, string>? cDict = null;
                        if (!string.IsNullOrWhiteSpace(item.CenterData))
                        {
                            try
                            {
                                cDict = JsonSerializer.Deserialize<Dictionary<string, string>>(item.CenterData);
                            }
                            catch { }
                        }

                        foreach (var key in extraHeaderList)
                        {
                            string val = "";
                            if (cDict != null && cDict.TryGetValue(key, out var dynVal))
                            {
                                val = dynVal ?? "";
                            }
                            ws.Cells[rowIdx, col++].Value = val;
                        }

                        using (var range = ws.Cells[rowIdx, 1, rowIdx, col - 1])
                        {
                            range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            if (deletedCenters.Any(x => x.Id == item.Id))
                                range.Style.Fill.BackgroundColor.SetColor(Color.LightCoral);
                            else if (newCenters.Any(x => x.Id == item.Id))
                                range.Style.Fill.BackgroundColor.SetColor(Color.LightGreen);
                        }

                        rowIdx++;
                    }

                    if (ws.Dimension != null)
                    {
                        ws.Cells[ws.Dimension.Address].AutoFitColumns();
                        ws.View.FreezePanes(2, 1);
                    }

                    // Sheet 2: Clean Data (Active Centers)
                    var wsClean = package.Workbook.Worksheets.Add("Clean Data");
                    col = 1;
                    foreach (var h in baseHeaders)
                    {
                        wsClean.Cells[1, col].Value = h;
                        wsClean.Cells[1, col].Style.Font.Bold = true;
                        col++;
                    }
                    foreach (var key in extraHeaderList)
                    {
                        wsClean.Cells[1, col].Value = key;
                        wsClean.Cells[1, col].Style.Font.Bold = true;
                        col++;
                    }

                    var cleanCenters = await _context.CenterList
                        .Where(x => x.ProjectId == ProjectId && x.Status && x.ExtraId == 0)
                        .ToListAsync();

                    int cleanRow = 2;
                    foreach (var item in cleanCenters)
                    {
                        nrDataMap.TryGetValue(item.NRDataId, out var catchInfo);

                        col = 1;
                        wsClean.Cells[cleanRow, col++].Value = catchInfo?.CatchNo ?? "";
                        wsClean.Cells[cleanRow, col++].Value = catchInfo?.SubjectName ?? "";
                        wsClean.Cells[cleanRow, col++].Value = catchInfo?.CourseName ?? "";
                        wsClean.Cells[cleanRow, col++].Value = item.NodalCode ?? "";
                        wsClean.Cells[cleanRow, col++].Value = item.CenterCode ?? "";
                        wsClean.Cells[cleanRow, col++].Value = item.Route ?? "";
                        wsClean.Cells[cleanRow, col++].Value = item.Quantity;
                        wsClean.Cells[cleanRow, col++].Value = item.NRQuantity;
                        wsClean.Cells[cleanRow, col++].Value = "Active";

                        Dictionary<string, string>? cDict = null;
                        if (!string.IsNullOrWhiteSpace(item.CenterData))
                        {
                            try
                            {
                                cDict = JsonSerializer.Deserialize<Dictionary<string, string>>(item.CenterData);
                            }
                            catch { }
                        }

                        foreach (var key in extraHeaderList)
                        {
                            string val = "";
                            if (cDict != null && cDict.TryGetValue(key, out var dynVal))
                            {
                                val = dynVal ?? "";
                            }
                            wsClean.Cells[cleanRow, col++].Value = val;
                        }

                        cleanRow++;
                    }

                    if (wsClean.Dimension != null)
                    {
                        wsClean.Cells[wsClean.Dimension.Address].AutoFitColumns();
                        wsClean.View.FreezePanes(2, 1);
                    }

                    package.SaveAs(new FileInfo(filePath));
                }

                // Record in ExcelReports table
                var dupVersionMatch = System.Text.RegularExpressions.Regex.Match(fileName, @"_v(\d+)\.xlsx$", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                int dupVersion = dupVersionMatch.Success ? int.Parse(dupVersionMatch.Groups[1].Value) : 1;
                await ExcelReportHelper.RecordExcelReportAsync(
                    _context,
                    ProjectId,
                    1, // Module 1 (Duplicate Tool)
                    dupVersion,
                    lotNo.HasValue && lotNo.Value > 0 ? lotNo : null,
                    filePath,
                    true,
                    LogHelper.GetTriggeredBy(User)
                );

                return Ok(new
                {
                    MergedRows = mergedCount,
                    message = mergedCount > 0
                        ? "Duplicates processed successfully for NrData1."
                        : "No pending duplicate rows. Report regenerated from active data.",
                    fileName,
                    filePath
                });
            }
            catch (Exception ex)
            {
                await _loggerService.LogErrorAsync("Error solving duplicates for NrData1", ex.ToString(), nameof(NrData1Controller));
                return StatusCode(500, "Internal server error: " + ex.Message);
            }
        }

        [HttpPost("Enhancement")]
        public async Task<IActionResult> ApplyEnhancement([FromQuery] int ProjectId, [FromQuery] int? batch = null, [FromQuery] int? lotNo = null)
        {
            try
            {
                Console.WriteLine($"[NrData1Controller] ApplyEnhancement API called for ProjectId: {ProjectId}, batch: {batch}, lotNo: {lotNo}");

                IQueryable<NrData1> query = _context.NrData1
                    .Where(p => p.ProjectId == ProjectId && p.Status);

                if (batch.HasValue && batch.Value > 0)
                {
                    query = query.Where(p => p.Batch == batch.Value);
                }
                else
                {
                    query = query.Where(p => p.Batch == 1 || p.Steps == PipelineNavigator.STEP_DUP_PARTIAL);
                }

                if (lotNo.HasValue && lotNo.Value > 0)
                {
                    Console.WriteLine($"[NrData1Controller] Filtering by lotNo: {lotNo.Value}");
                    query = query.Where(p => p.LotNo == lotNo.Value);
                }

                var data = await query.ToListAsync();

                var projectconfig = await _context.ProjectConfigs
                    .Where(p => p.ProjectId == ProjectId).FirstOrDefaultAsync();

                if (projectconfig == null)
                {
                    return NotFound("Project config not exists for this project");
                }

                int smallestInner = 0;
                var innerEnv = projectconfig.Envelope;

                if (!string.IsNullOrEmpty(innerEnv))
                {
                    try
                    {
                        var envelopeDict = JsonSerializer.Deserialize<Dictionary<string, string>>(innerEnv);
                        if (envelopeDict != null && envelopeDict.TryGetValue("Inner", out var innerValue) &&
                            !string.IsNullOrWhiteSpace(innerValue))
                        {
                            var innerSizes = innerValue
                                .Split(',', StringSplitOptions.RemoveEmptyEntries)
                                .Select(e => e.Trim().ToUpper().Replace("E", ""))
                                .Where(x => int.TryParse(x, out _))
                                .Select(int.Parse)
                                .OrderBy(x => x)
                                .ToList();

                            if (innerSizes.Any())
                            {
                                smallestInner = innerSizes.First();
                            }
                        }
                        else if (envelopeDict != null && envelopeDict.TryGetValue("Outer", out var outerValue) &&
                                 !string.IsNullOrWhiteSpace(outerValue))
                        {
                            var outerSizes = outerValue
                                .Split(',', StringSplitOptions.RemoveEmptyEntries)
                                .Select(e => e.Trim().ToUpper().Replace("E", ""))
                                .Where(x => int.TryParse(x, out _))
                                .Select(int.Parse)
                                .OrderBy(x => x)
                                .ToList();

                            if (outerSizes.Any())
                            {
                                smallestInner = outerSizes.First();
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        await _loggerService.LogErrorAsync("Error in serializing Envelope", ex.Message, nameof(NrData1Controller));
                        smallestInner = 0;
                    }
                }

                // Consolidated calculation logic: Round Before (Optional) -> Enhance -> Round After
                if (data.Any())
                {
                    var nrDataIds = data.Select(x => x.Id).ToList();
                    var relatedCenters = await _context.CenterList
                        .Where(c => c.ProjectId == ProjectId && nrDataIds.Contains(c.NRDataId) && c.Status)
                        .ToListAsync();

                    var centersByNrDataId = relatedCenters
                        .GroupBy(c => c.NRDataId)
                        .ToDictionary(g => g.Key, g => g.ToList());

                    foreach (var d in data)
                    {
                        d.Steps = PipelineNavigator.STEP_ENHANCEMENT;

                        if (centersByNrDataId.TryGetValue(d.Id, out var centers))
                        {
                            foreach (var center in centers)
                            {
                                int cNrQty = center.NRQuantity;
                                if (cNrQty <= 0 && !string.IsNullOrWhiteSpace(center.CenterData))
                                {
                                    try
                                    {
                                        var cDict = JsonSerializer.Deserialize<Dictionary<string, string>>(center.CenterData);
                                        if (cDict != null && cDict.TryGetValue("NRQuantity", out var cNrQtyStr) &&
                                            int.TryParse(cNrQtyStr, out int parsedVal))
                                        {
                                            cNrQty = parsedVal;
                                        }
                                    }
                                    catch { }
                                }

                                if (cNrQty > 0)
                                {
                                    double cInitQty = cNrQty;
                                    if (projectconfig.RoundOffBeforeEnhancement && smallestInner > 0)
                                    {
                                        cInitQty = Math.Ceiling(cInitQty / (double)smallestInner) * smallestInner;
                                    }

                                    double cEnhVal = projectconfig.Enhancement > 0
                                        ? (projectconfig.Enhancement * cInitQty) / 100.0
                                        : 0;
                                    double cTotal = cInitQty + cEnhVal;
                                    int cEnhancedQty = smallestInner > 0
                                        ? (int)Math.Ceiling(cTotal / (double)smallestInner) * smallestInner
                                        : (int)Math.Round(cTotal);

                                    center.Quantity = cEnhancedQty;

                                    if (!string.IsNullOrWhiteSpace(center.CenterData))
                                    {
                                        try
                                        {
                                            var cDict = JsonSerializer.Deserialize<Dictionary<string, string>>(center.CenterData) ?? new Dictionary<string, string>();
                                            cDict["Quantity"] = cEnhancedQty.ToString();
                                            center.CenterData = JsonSerializer.Serialize(cDict);
                                        }
                                        catch { }
                                    }

                                    _context.CenterList.Update(center);
                                }
                            }
                        }
                    }

                    await _context.SaveChangesAsync();
                }

                var filePath = string.Empty;

                if (data.Any())
                {
                    var reportPath = FileStorageHelper.GetProjectFolder(ProjectId);

                    var distinctLots = (lotNo.HasValue && lotNo.Value > 0)
                        ? new List<int> { lotNo.Value }
                        : data.Where(r => r.EnvLotNo > 0).Select(r => r.EnvLotNo).Distinct().OrderBy(l => l).ToList();
                    var lotStr = distinctLots.Any() ? string.Join("_", distinctLots) : "All";

                    var fileName = ReportVersionHelper.GetNextVersionFileName(reportPath, $"EnhancementReport_NrData1_{lotStr}.xlsx");
                    filePath = Path.Combine(reportPath, fileName);

                    var baseProperties = typeof(NrData1).GetProperties(BindingFlags.Public | BindingFlags.Instance)
                        .Where(p => p.Name != "NRDatas" && p.Name != "ProjectId")
                        .ToList();

                    ExcelPackage.License.SetNonCommercialPersonal("Tools");
                    using (var package = new ExcelPackage())
                    {
                        var ws = package.Workbook.Worksheets.Add("Enhancement Report");

                        int col = 1;
                        foreach (var prop in baseProperties)
                        {
                            ws.Cells[1, col].Value = prop.Name;
                            ws.Cells[1, col].Style.Font.Bold = true;
                            col++;
                        }

                        int row = 2;
                        foreach (var item in data)
                        {
                            col = 1;
                            foreach (var prop in baseProperties)
                            {
                                ws.Cells[row, col++].Value = prop.GetValue(item)?.ToString();
                            }
                            row++;
                        }

                        ws.Cells[ws.Dimension.Address].AutoFitColumns();
                        ws.View.FreezePanes(2, 1);

                        package.SaveAs(new FileInfo(filePath));

                        // Record in ExcelReports table
                        var enhVersionMatch = System.Text.RegularExpressions.Regex.Match(fileName, @"_v(\d+)\.xlsx$", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                        int enhVersion = enhVersionMatch.Success ? int.Parse(enhVersionMatch.Groups[1].Value) : 1;
                        await ExcelReportHelper.RecordExcelReportAsync(
                            _context,
                            ProjectId,
                            2, // Module 2 (Envelope Setup and Enhancement)
                            enhVersion,
                            lotNo.HasValue && lotNo.Value > 0 ? lotNo : null,
                            filePath,
                            true,
                            LogHelper.GetTriggeredBy(User)
                        );
                    }

                    // Logging
                    var triggeredBy = LogHelper.GetTriggeredBy(User);

                    await _loggerService.LogEventAsync(
                        "Enhancement applied with report (NrData1)",
                        "Enhancement",
                        triggeredBy,
                        ProjectId,
                        string.Empty,
                        LogHelper.ToJson(new { ProjectId })
                    );
                }

                try
                {
                    Console.WriteLine("[NrData1Controller] Envelope breaking is called");
                    await EnvelopeConfiguration(ProjectId, bypassDispatch: true);
                }
                catch (Exception ex)
                {
                    await _loggerService.LogErrorAsync("Error running combined EnvelopeConfiguration", ex.Message, nameof(NrData1Controller));
                }

                return Ok(new
                {
                    EnhancementApplied = data.Any() ? projectconfig.Enhancement : 0.0,
                    ReportPath = filePath,
                    fileName = !string.IsNullOrEmpty(filePath) ? Path.GetFileName(filePath) : null
                });
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[NrData1Controller ApplyEnhancement] Exception thrown: {ex}");
                await _loggerService.LogErrorAsync("Error applying enhancement for NrData1", ex.ToString(), nameof(NrData1Controller));
                return StatusCode(500, "Internal server error: " + ex.Message);
            }
        }

        [HttpPost("EnvelopeConfiguration")]
        public async Task<IActionResult> EnvelopeConfiguration([FromQuery] int ProjectId, [FromQuery] bool bypassDispatch = false)
        {
            try
            {
                var projectConfig = await _context.ProjectConfigs
                    .FirstOrDefaultAsync(s => s.ProjectId == ProjectId);

                if (projectConfig == null)
                    return NotFound("Project config not found for this project.");

                var envelopesJson = projectConfig.Envelope;
                if (string.IsNullOrWhiteSpace(envelopesJson))
                    return BadRequest("Envelope configuration not found in ProjectConfig.");

                var envelopeDict = JsonSerializer.Deserialize<Dictionary<string, string>>(envelopesJson);
                if (envelopeDict == null)
                    return BadRequest("Invalid envelope JSON format.");

                var innerSizes = envelopeDict["Inner"]
                    .Split(',', StringSplitOptions.RemoveEmptyEntries)
                    .Select(e => e.Trim().ToUpper().Replace("E", ""))
                    .Where(x => int.TryParse(x, out _))
                    .Select(int.Parse)
                    .OrderByDescending(x => x)
                    .ToList();

                var outerSizes = envelopeDict["Outer"]
                    .Split(',', StringSplitOptions.RemoveEmptyEntries)
                    .Select(e => e.Trim().ToUpper().Replace("E", ""))
                    .Where(x => int.TryParse(x, out _))
                    .Select(int.Parse)
                    .OrderByDescending(x => x)
                    .ToList();

                // 1. Fetch enhanced NrData1 records for Project
                var nrDataList = await _context.NrData1
                    .Where(s => s.ProjectId == ProjectId && s.Status && s.Steps == PipelineNavigator.STEP_ENHANCEMENT)
                    .ToListAsync();

                if (!nrDataList.Any())
                {
                    return Ok("Envelope breakdown report has been successfully created (no pending enhanced records).");
                }

                var nrDataIds = nrDataList.Select(r => r.Id).ToList();

                // 2. Fetch active centers for these NrData1 records
                var centerList = await _context.CenterList
                    .Where(c => c.ProjectId == ProjectId && nrDataIds.Contains(c.NRDataId) && c.Status)
                    .ToListAsync();

                bool hasExtraConfig = await _context.ExtraConfigurations.AnyAsync(e => e.ProjectId == ProjectId);

                // 3. Deactivate existing active NewEnvelopeBreakage entries for active centers / NrDataIds
                var centerIds = centerList.Select(c => c.Id).ToList();
                var existingBreakages = await _context.NewEnvelopeBreakages
                    .Where(p => p.ProjectId == ProjectId && centerIds.Contains(p.CenterListId) && p.Status == true)
                    .ToListAsync();

                if (existingBreakages.Any())
                {
                    foreach (var item in existingBreakages)
                    {
                        item.Status = false;
                        _context.NewEnvelopeBreakages.Update(item);
                    }
                    await _context.SaveChangesAsync();

                    await _loggerService.LogEventAsync(
                        $"Successfully deactivated {existingBreakages.Count} NewEnvelopeBreakage entries (Status=0) for ProjectID {ProjectId}",
                        "NewEnvelopeBreakages",
                        LogHelper.GetTriggeredBy(User),
                        ProjectId
                    );
                }

                var breakagesToAdd = new List<NewEnvelopeBreakage>();
                var legacyBreakagesToAdd = new List<EnvelopeBreakage>();

                // 4. Calculate envelope breaking for EACH center
                foreach (var center in centerList)
                {
                    int quantity = center.Quantity > 0 ? center.Quantity : center.NRQuantity;
                    if (quantity <= 0 && !string.IsNullOrWhiteSpace(center.CenterData))
                    {
                        try
                        {
                            var cDict = JsonSerializer.Deserialize<Dictionary<string, string>>(center.CenterData);
                            if (cDict != null && cDict.TryGetValue("Quantity", out var qStr) && int.TryParse(qStr, out int qVal))
                            {
                                quantity = qVal;
                            }
                            else if (cDict != null && cDict.TryGetValue("NRQuantity", out var nrqStr) && int.TryParse(nrqStr, out int nrqVal))
                            {
                                quantity = nrqVal;
                            }
                        }
                        catch { }
                    }

                    if (quantity <= 0) continue;

                    // Inner breakdown
                    int remaining = quantity;
                    var innerBreakdown = new Dictionary<string, string>();

                    if (innerSizes.Any())
                    {
                        foreach (var size in innerSizes)
                        {
                            int count = remaining / size;
                            if (count > 0)
                            {
                                innerBreakdown[$"E{size}"] = count.ToString();
                                remaining -= count * size;
                            }
                        }

                        if (remaining > 0)
                        {
                            int smallestInner = innerSizes.Last();
                            int count = (int)Math.Ceiling((double)remaining / smallestInner);
                            string key = $"E{smallestInner}";
                            innerBreakdown[key] = (innerBreakdown.ContainsKey(key)
                                ? int.Parse(innerBreakdown[key]) + count
                                : count).ToString();
                        }
                    }

                    // Outer breakdown
                    remaining = quantity;
                    var outerBreakdown = new Dictionary<string, string>();
                    int totalOuterCount = 0;

                    if (outerSizes.Any())
                    {
                        foreach (var size in outerSizes)
                        {
                            int count = remaining / size;
                            if (count > 0)
                            {
                                outerBreakdown[$"E{size}"] = count.ToString();
                                totalOuterCount += count;
                                remaining -= count * size;
                            }
                        }

                        if (remaining > 0)
                        {
                            int smallestSize = outerSizes.Last();
                            int count = (int)Math.Ceiling((double)remaining / smallestSize);

                            string key = $"E{smallestSize}";
                            if (outerBreakdown.ContainsKey(key))
                            {
                                outerBreakdown[key] = (int.Parse(outerBreakdown[key]) + count).ToString();
                            }
                            else
                            {
                                outerBreakdown[key] = count.ToString();
                            }

                            totalOuterCount += count;
                            remaining = 0;
                        }
                    }

                    if (!innerBreakdown.Any() && outerBreakdown.Any())
                    {
                        innerBreakdown = new Dictionary<string, string>(outerBreakdown);
                    }

                    var newBreakage = new NewEnvelopeBreakage
                    {
                        ProjectId = ProjectId,
                        NrDataId = center.NRDataId,
                        CenterListId = center.Id,
                        InnerEnvelope = JsonSerializer.Serialize(innerBreakdown),
                        OuterEnvelope = JsonSerializer.Serialize(outerBreakdown),
                        TotalEnvelope = totalOuterCount,
                        Status = true
                    };
                    breakagesToAdd.Add(newBreakage);

                    var legacyBreakage = new EnvelopeBreakage
                    {
                        ProjectId = ProjectId,
                        NrDataId = center.NRDataId,
                        InnerEnvelope = JsonSerializer.Serialize(innerBreakdown),
                        OuterEnvelope = JsonSerializer.Serialize(outerBreakdown),
                        TotalEnvelope = totalOuterCount,
                        Status = true
                    };
                    legacyBreakagesToAdd.Add(legacyBreakage);
                }

                if (breakagesToAdd.Any())
                {
                    await _context.NewEnvelopeBreakages.AddRangeAsync(breakagesToAdd);
                    await _context.EnvelopeBreakages.AddRangeAsync(legacyBreakagesToAdd);

                    // Update steps for NrData1
                    foreach (var nr in nrDataList)
                    {
                        if (hasExtraConfig)
                        {
                            nr.Steps = PipelineNavigator.STEP_ENV_BREAKING;
                        }
                        else
                        {
                            nr.Steps = PipelineNavigator.STEP_AWAITING_EXTRA;
                        }
                    }

                    await _context.SaveChangesAsync();
                    await _loggerService.LogEventAsync(
                        $"Created {breakagesToAdd.Count} NewEnvelopeBreakage entries for ProjectID {ProjectId} and updated steps",
                        "NewEnvelopeBreakages",
                        LogHelper.GetTriggeredBy(User),
                        ProjectId
                    );
                }

                return Ok("Envelope breakdown report has been successfully created.");
            }
            catch (Exception ex)
            {
                await _loggerService.LogErrorAsync("Error creating NewEnvelopeBreakage", ex.ToString(), nameof(NrData1Controller));
                return StatusCode(500, "Internal Server Error: " + ex.Message);
            }
        }

        [HttpPost("PostExtraEnvelopes")]
        public async Task<ActionResult> PostExtraEnvelopes([FromQuery] int ProjectId, [FromQuery] int? uploadId = null, [FromQuery] int? batchNo = null, [FromQuery] int? lotNo = null)
        {
            try
            {
                var projectConfig = await _context.ProjectConfigs
                    .FirstOrDefaultAsync(p => p.ProjectId == ProjectId);

                var eligibleSteps = PipelineNavigator.GetEligiblePickupSteps(PipelineNavigator.STEP_AWAITING_EXTRA);

                var query = _context.NrData1
                    .Where(d => d.ProjectId == ProjectId && d.Status == true && eligibleSteps.Contains(d.Steps) && d.Batch == (batchNo ?? 1));

                if (lotNo.HasValue && lotNo.Value > 0)
                {
                    query = query.Where(d => d.LotNo == lotNo.Value);
                }

                var nrDataList = await query.ToListAsync();

                if (!nrDataList.Any())
                {
                    // Fallback to all active records in this batch
                    query = _context.NrData1.Where(d => d.ProjectId == ProjectId && d.Status == true && d.Batch == (batchNo ?? 1));
                    if (lotNo.HasValue && lotNo.Value > 0)
                        query = query.Where(d => d.LotNo == lotNo.Value);
                    nrDataList = await query.ToListAsync();
                }

                if (!nrDataList.Any())
                    return BadRequest("No suitable data found for Extra Configuration. Either no data is at the required step or all data is already processed.");

                var extraConfig = await _context.ExtraConfigurations
                    .Where(c => c.ProjectId == ProjectId)
                    .ToListAsync();

                var allExtraTypes = await _context.ExtraType
                    .OrderBy(e => e.Order)
                    .ToListAsync();
                var extraTypesMap = allExtraTypes.ToDictionary(e => e.ExtraTypeId, e => e.Type);
                var extraTypeOrderMap = allExtraTypes.ToDictionary(e => e.ExtraTypeId, e => e.Order);

                // Sort extra configurations according to ExtraType table Order
                extraConfig = extraConfig
                    .OrderBy(c => extraTypeOrderMap.TryGetValue(c.ExtraType, out var ord) ? ord : c.ExtraType)
                    .ToList();

                var nrDataIds = nrDataList.Select(d => d.Id).ToList();
                var existingCenters = await _context.CenterList
                    .Where(c => c.ProjectId == ProjectId && nrDataIds.Contains(c.NRDataId) && c.ExtraId == 0 && c.Status)
                    .ToListAsync();

                bool hasExistingExtraCenters = await _context.CenterList
                    .AnyAsync(c => c.ProjectId == ProjectId && c.Status &&
                        (c.ExtraId > 0 || (c.CenterData != null && (c.CenterData.Contains("Reserve") || c.CenterData.Contains("Extra"))) || (c.CenterSort > 0 && Math.Abs(c.CenterSort - Math.Round(c.CenterSort)) > 0.001)));

                bool isOnlyReport = extraConfig.Any(c => c.IsExtraProcessingAsPerNR) || (!extraConfig.Any() && hasExistingExtraCenters);
                var envelopesToUse = new List<ExtraEnvelopeReportDto>();

                if (isOnlyReport)
                {
                    var extraCenters = await _context.CenterList
                        .Where(c => c.ProjectId == ProjectId && c.ExtraId > 0 && c.Status)
                        .ToListAsync();

                    if (!extraCenters.Any())
                    {
                        var centersWithExtraName = await _context.CenterList
                            .Where(c => c.ProjectId == ProjectId && c.Status && c.CenterCode != null)
                            .ToListAsync();

                        var matchedCenters = centersWithExtraName
                            .Where(c => allExtraTypes.Any(et => !string.IsNullOrEmpty(et.Type) && c.CenterCode!.Contains(et.Type, StringComparison.OrdinalIgnoreCase)) ||
                                        c.CenterCode!.Contains("Extra", StringComparison.OrdinalIgnoreCase))
                            .ToList();

                        if (matchedCenters.Any())
                        {
                            foreach (var mc in matchedCenters)
                            {
                                if (mc.ExtraId == 0)
                                {
                                    var matchedType = allExtraTypes.FirstOrDefault(et =>
                                        !string.IsNullOrEmpty(et.Type) && mc.CenterCode!.Contains(et.Type, StringComparison.OrdinalIgnoreCase));
                                    mc.ExtraId = matchedType?.ExtraTypeId ?? 1;
                                    _context.CenterList.Update(mc);
                                }
                            }
                            await _context.SaveChangesAsync();
                            extraCenters = matchedCenters;
                        }
                    }

                    // Check for Option 2: Reserve Quantity rows specified in Remark column or decimal CenterSort
                    if (!extraCenters.Any())
                    {
                        var potentialReserveCenters = await _context.CenterList
                            .Where(c => c.ProjectId == ProjectId && nrDataIds.Contains(c.NRDataId) && c.Status)
                            .ToListAsync();

                        var matchedReserveCenters = new List<CenterList>();
                        foreach (var c in potentialReserveCenters)
                        {
                            bool isReserve = false;
                            int reserveQty = c.Quantity > 0 ? c.Quantity : c.NRQuantity;

                            if (!string.IsNullOrWhiteSpace(c.CenterData))
                            {
                                try
                                {
                                    using var doc = JsonDocument.Parse(c.CenterData);
                                    foreach (var prop in doc.RootElement.EnumerateObject())
                                    {
                                        string pName = prop.Name;
                                        string pVal = prop.Value.ToString();

                                        if (pName.Contains("Remark", StringComparison.OrdinalIgnoreCase) &&
                                            pVal.Contains("Reserve", StringComparison.OrdinalIgnoreCase))
                                        {
                                            isReserve = true;
                                        }

                                        if (reserveQty <= 0 &&
                                            (pName.Equals("Qty", StringComparison.OrdinalIgnoreCase) ||
                                             pName.Equals("Quantity", StringComparison.OrdinalIgnoreCase) ||
                                             pName.Contains("Reserve", StringComparison.OrdinalIgnoreCase)) &&
                                            int.TryParse(pVal, out int parsedQty) && parsedQty > 0)
                                        {
                                            reserveQty = parsedQty;
                                        }
                                    }
                                }
                                catch { }
                            }

                            // Decimal CenterSort like 1.1 indicates a reserve center
                            if (!isReserve && c.CenterSort > 0 && Math.Abs(c.CenterSort - Math.Round(c.CenterSort)) > 0.001)
                            {
                                isReserve = true;
                            }

                            if (isReserve)
                            {
                                if (c.ExtraId == 0)
                                {
                                    c.ExtraId = 1; // Default to Nodal Extra / Reserve
                                }
                                if (reserveQty > 0)
                                {
                                    c.Quantity = reserveQty;
                                    c.NRQuantity = reserveQty;
                                }
                                _context.CenterList.Update(c);
                                matchedReserveCenters.Add(c);
                            }
                        }

                        if (matchedReserveCenters.Any())
                        {
                            await _context.SaveChangesAsync();
                            extraCenters = matchedReserveCenters;
                        }
                    }

                    if (extraCenters.Any())
                    {
                        var extraCenterIds = extraCenters.Select(c => c.Id).ToList();
                        var breakages = await _context.NewEnvelopeBreakages
                            .Where(b => b.ProjectId == ProjectId && extraCenterIds.Contains(b.CenterListId) && b.Status)
                            .ToListAsync();

                        // Ensure every extra center has NewEnvelopeBreakage & EnvelopeBreakage records for downstream processing
                        var newBreakagesToInsert = new List<NewEnvelopeBreakage>();
                        var legBreakagesToInsert = new List<EnvelopeBreakage>();

                        var envConfigObj = extraConfig.FirstOrDefault(ec => ec.IsExtraProcessingAsPerNR)
                            ?? extraConfig.FirstOrDefault();
                        string? envJson = envConfigObj?.EnvelopeType ?? projectConfig?.Envelope;

                        int outerCapacity = 30;
                        int innerCapacity = 0;
                        if (!string.IsNullOrWhiteSpace(envJson))
                        {
                            try
                            {
                                var envDict = JsonSerializer.Deserialize<Dictionary<string, string>>(envJson);
                                if (envDict != null)
                                {
                                    if (envDict.TryGetValue("Outer", out var oVal))
                                    {
                                        var cleanO = oVal.Trim().ToUpper().Replace("E", "");
                                        if (int.TryParse(cleanO, out int oCap) && oCap > 0) outerCapacity = oCap;
                                    }
                                    if (envDict.TryGetValue("Inner", out var iVal))
                                    {
                                        var cleanI = iVal.Trim().ToUpper().Replace("E", "");
                                        if (int.TryParse(cleanI, out int iCap) && iCap > 0) innerCapacity = iCap;
                                    }
                                }
                            }
                            catch { }
                        }

                        foreach (var c in extraCenters)
                        {
                            var b = breakages.FirstOrDefault(x => x.CenterListId == c.Id);
                            if (b == null)
                            {
                                int centerQty = c.Quantity > 0 ? c.Quantity : c.NRQuantity;
                                if (centerQty <= 0) centerQty = outerCapacity;

                                int totalEnv = 1;
                                if (!string.IsNullOrWhiteSpace(c.CenterData))
                                {
                                    try
                                    {
                                        using var doc = JsonDocument.Parse(c.CenterData);
                                        foreach (var prop in doc.RootElement.EnumerateObject())
                                        {
                                            if (prop.Name.Contains("Env", StringComparison.OrdinalIgnoreCase) ||
                                                prop.Name.Contains("Total", StringComparison.OrdinalIgnoreCase))
                                            {
                                                if (int.TryParse(prop.Value.ToString(), out int parsedEnv) && parsedEnv > 0)
                                                {
                                                    totalEnv = parsedEnv;
                                                    break;
                                                }
                                            }
                                        }
                                    }
                                    catch { }
                                }

                                var innerBreakdown = new Dictionary<string, string>();
                                var outerBreakdown = new Dictionary<string, string>();

                                string outerKey = $"E{outerCapacity}";
                                outerBreakdown[outerKey] = totalEnv.ToString();

                                if (innerCapacity > 0)
                                {
                                    string innerKey = $"E{innerCapacity}";
                                    innerBreakdown[innerKey] = totalEnv.ToString();
                                }
                                else
                                {
                                    innerBreakdown[outerKey] = totalEnv.ToString();
                                }

                                string innerStr = JsonSerializer.Serialize(innerBreakdown);
                                string outerStr = JsonSerializer.Serialize(outerBreakdown);

                                var newB = new NewEnvelopeBreakage
                                {
                                    ProjectId = ProjectId,
                                    NrDataId = c.NRDataId,
                                    CenterListId = c.Id,
                                    InnerEnvelope = innerStr,
                                    OuterEnvelope = outerStr,
                                    TotalEnvelope = totalEnv,
                                    Status = true
                                };
                                newBreakagesToInsert.Add(newB);

                                var legB = new EnvelopeBreakage
                                {
                                    ProjectId = ProjectId,
                                    NrDataId = c.NRDataId,
                                    InnerEnvelope = innerStr,
                                    OuterEnvelope = outerStr,
                                    TotalEnvelope = totalEnv,
                                    Status = true
                                };
                                legBreakagesToInsert.Add(legB);
                            }
                        }

                        if (newBreakagesToInsert.Any())
                        {
                            await _context.NewEnvelopeBreakages.AddRangeAsync(newBreakagesToInsert);
                            await _context.EnvelopeBreakages.AddRangeAsync(legBreakagesToInsert);
                            await _context.SaveChangesAsync();

                            breakages.AddRange(newBreakagesToInsert);
                        }

                        var nrLookup = nrDataList.ToDictionary(n => n.Id, n => n.CatchNo);

                        foreach (var c in extraCenters)
                        {
                            var b = breakages.FirstOrDefault(x => x.CenterListId == c.Id);
                            nrLookup.TryGetValue(c.NRDataId, out var catchNo);

                            envelopesToUse.Add(new ExtraEnvelopeReportDto
                            {
                                CatchNo = catchNo,
                                NodalCode = c.NodalCode,
                                ExtraId = c.ExtraId,
                                Quantity = c.Quantity > 0 ? c.Quantity : c.NRQuantity,
                                InnerEnvelope = b?.InnerEnvelope,
                                OuterEnvelope = b?.OuterEnvelope
                            });
                        }
                    }
                    else
                    {
                        // Check ExtrasEnvelope table as alternative storage from upload
                        var extrasEnvelopeData = await _context.ExtrasEnvelope
                            .Where(e => e.ProjectId == ProjectId && e.Status == 1)
                            .ToListAsync();

                        if (extrasEnvelopeData.Any())
                        {
                            foreach (var ex in extrasEnvelopeData)
                            {
                                envelopesToUse.Add(new ExtraEnvelopeReportDto
                                {
                                    CatchNo = ex.CatchNo,
                                    NodalCode = ex.NodalCode,
                                    ExtraId = ex.ExtraId,
                                    Quantity = ex.Quantity,
                                    InnerEnvelope = ex.InnerEnvelope,
                                    OuterEnvelope = ex.OuterEnvelope
                                });
                            }
                        }
                    }

                    if (!envelopesToUse.Any())
                    {
                        bool hasConfiguredCalculations = extraConfig.Any(c =>
                            (!string.IsNullOrWhiteSpace(c.Value) && c.Value != "0") ||
                            !string.IsNullOrWhiteSpace(c.nodalValue) ||
                            !string.IsNullOrWhiteSpace(c.RangeConfig));

                        if (!hasConfiguredCalculations)
                        {
                            return BadRequest("No existing Extra Envelope data found in CenterList or NR data. 'Extra Processing as per NR' requires extra packet rows (e.g., University Extra, Nodal Extra, or Reserve Quantity) to be present in the uploaded NR file. If your NR file does not contain extra rows, please uncheck 'Extra Processing as per NR' in Project Configuration, configure envelopes and quantities, and re-run.");
                        }

                        // If user has configured calculations, switch off isOnlyReport and proceed to calculation
                        isOnlyReport = false;
                    }
                }

                if (!isOnlyReport)
                {
                    if (!extraConfig.Any())
                        return BadRequest("No ExtraConfiguration found");

                    var envelopesToAdd = new List<ExtraEnvelopeReportDto>();

                    // Pre-fetch all active extra centers and their breakages once to avoid DB calls in loop
                    var existingExtraCenters = await _context.CenterList
                        .Where(c => c.ProjectId == ProjectId && c.ExtraId > 0 && c.Status)
                        .ToListAsync();

                    var existingExtraCenterIds = existingExtraCenters.Select(c => c.Id).ToList();
                    var existingNewBreakages = await _context.NewEnvelopeBreakages
                        .Where(b => b.ProjectId == ProjectId && existingExtraCenterIds.Contains(b.CenterListId) && b.Status)
                        .ToListAsync();

                    var allNodals = await _context.CenterList
                        .Where(n => n.ProjectId == ProjectId && n.Status && !string.IsNullOrEmpty(n.NodalCode))
                        .Select(n => n.NodalCode!.Trim())
                        .Distinct()
                        .ToListAsync();

                    var newCentersToInsert = new List<CenterList>();
                    var pendingBreakages = new List<(CenterList CenterRecord, int NrDataId, Dictionary<string, string> InnerBreakdown, Dictionary<string, string> OuterBreakdown, int TotalEnvelope)>();

                    foreach (var config in extraConfig)
                    {
                        var extraTypeEntity = allExtraTypes.FirstOrDefault(e => e.ExtraTypeId == config.ExtraType);
                        string extraTypeName = extraTypeEntity?.Type ?? (extraTypesMap.TryGetValue(config.ExtraType, out var t) ? t : "");
                        bool isNodalExtra = (!string.IsNullOrWhiteSpace(extraTypeName) && extraTypeName.Contains("Nodal", StringComparison.OrdinalIgnoreCase))
                            || (extraTypeEntity != null && extraTypeEntity.Order == 1);
                        bool useNodal = isNodalExtra || !string.IsNullOrWhiteSpace(config.nodalValue);

                        EnvelopeTypeConfig? envelopeType = null;
                        if (!string.IsNullOrWhiteSpace(config.EnvelopeType))
                        {
                            try { envelopeType = JsonSerializer.Deserialize<EnvelopeTypeConfig>(config.EnvelopeType); }
                            catch { }
                        }

                        int innerCapacity = envelopeType != null && !string.IsNullOrWhiteSpace(envelopeType.Inner) ? GetEnvelopeCapacity(envelopeType.Inner) : 1;
                        int outerCapacity = envelopeType != null && !string.IsNullOrWhiteSpace(envelopeType.Outer) ? GetEnvelopeCapacity(envelopeType.Outer) : 1;

                        // Group data
                        var groupedData = useNodal
                            ? nrDataList
                                .SelectMany(nr =>
                                {
                                    var centersForNr = existingCenters.Where(c => c.NRDataId == nr.Id && (!isNodalExtra || !string.IsNullOrWhiteSpace(c.NodalCode))).ToList();
                                    if (centersForNr.Any())
                                    {
                                        return centersForNr.Select(c => new
                                        {
                                            nr.CatchNo,
                                            NodalCode = (c.NodalCode ?? "").Trim(),
                                            Quantity = c.Quantity > 0 ? c.Quantity : c.NRQuantity,
                                            NrDataId = nr.Id
                                        });
                                    }
                                    return new[]
                                    {
                                        new
                                        {
                                            nr.CatchNo,
                                            NodalCode = "",
                                            Quantity = 0,
                                            NrDataId = nr.Id
                                        }
                                    };
                                })
                                .GroupBy(x => new { x.CatchNo, x.NodalCode })
                                .Select(g => new
                                {
                                    g.Key.CatchNo,
                                    g.Key.NodalCode,
                                    Quantity = g.Sum(x => x.Quantity),
                                    NrDataId = g.First().NrDataId
                                }).ToList()
                            : nrDataList
                                .GroupBy(x => x.CatchNo)
                                .Select(g =>
                                {
                                    var catchNrIds = g.Select(x => x.Id).ToList();
                                    var centersForCatch = existingCenters.Where(c => catchNrIds.Contains(c.NRDataId)).ToList();
                                    int totalCatchQty = centersForCatch.Sum(c => c.Quantity > 0 ? c.Quantity : c.NRQuantity);
                                    return new
                                    {
                                        CatchNo = g.Key,
                                        NodalCode = (string?)null,
                                        Quantity = totalCatchQty,
                                        NrDataId = g.First().Id
                                    };
                                }).ToList();

                        if (isNodalExtra && config.AttachExtraForEachCatchForAllNodal)
                        {
                            var distinctCatches = nrDataList.Select(x => x.CatchNo).Where(c => !string.IsNullOrEmpty(c)).Distinct().ToList();

                            var existingPairs = new HashSet<(string?, string)>(groupedData.Select(g => (g.CatchNo, g.NodalCode ?? "")));
                            foreach (var cNo in distinctCatches)
                            {
                                var matchingNr = nrDataList.FirstOrDefault(x => x.CatchNo == cNo);
                                foreach (var nCode in allNodals)
                                {
                                    if (!existingPairs.Contains((cNo, nCode)))
                                    {
                                        groupedData.Add(new
                                        {
                                            CatchNo = cNo,
                                            NodalCode = (string?)nCode,
                                            Quantity = 0,
                                            NrDataId = matchingNr?.Id ?? 0
                                        });
                                        existingPairs.Add((cNo, nCode));
                                    }
                                }
                            }
                            groupedData.RemoveAll(g => string.IsNullOrWhiteSpace(g.NodalCode));
                        }
                        else if (isNodalExtra)
                        {
                            groupedData.RemoveAll(g => string.IsNullOrWhiteSpace(g.NodalCode));
                        }

                        // Deduplicate within run so only 1 record per (NrDataId, NodalCode) is processed
                        groupedData = groupedData
                            .GroupBy(x => new { x.NrDataId, Nodal = (x.NodalCode ?? "").Trim() })
                            .Select(g => new
                            {
                                CatchNo = g.First().CatchNo,
                                NodalCode = g.First().NodalCode,
                                Quantity = g.Sum(x => x.Quantity),
                                NrDataId = g.Key.NrDataId
                            })
                            .ToList();

                        // Parse nodal configs if any
                        List<NodalValueConfig>? nodalConfigs = null;
                        if (useNodal && !string.IsNullOrWhiteSpace(config.nodalValue))
                        {
                            try
                            {
                                nodalConfigs = JsonSerializer.Deserialize<List<NodalValueConfig>>(
                                    config.nodalValue,
                                    new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
                            }
                            catch { }
                        }

                        foreach (var data in groupedData)
                        {
                            if (isNodalExtra && string.IsNullOrWhiteSpace(data.NodalCode))
                                continue;

                            int calculatedQuantity = 0;

                            if (useNodal && nodalConfigs != null)
                            {
                                var match = nodalConfigs.FirstOrDefault(nc =>
                                    !string.IsNullOrEmpty(nc.NodalCodes) &&
                                    nc.NodalCodes.Split(',')
                                        .Select(n => n.Trim())
                                        .Contains(data.NodalCode, StringComparer.OrdinalIgnoreCase));

                                if (match != null)
                                {
                                    switch (config.Mode)
                                    {
                                        case "Fixed":
                                            if (int.TryParse(match.Value, out var fv))
                                                calculatedQuantity = fv;
                                            break;

                                        case "Percentage":
                                            if (decimal.TryParse(match.Value, out var pv))
                                                calculatedQuantity = (int)Math.Round((double)(data.Quantity * pv) / 100);
                                            break;
                                    }
                                }
                                else
                                {
                                    switch (config.Mode)
                                    {
                                        case "Fixed":
                                            if (int.TryParse(config.Value, out var fqv))
                                                calculatedQuantity = fqv;
                                            break;

                                        case "Percentage":
                                            if (decimal.TryParse(config.Value, out var percent))
                                                calculatedQuantity = (int)Math.Round((double)(data.Quantity * percent) / 100);
                                            break;

                                        case "Range":
                                            if (!string.IsNullOrEmpty(config.RangeConfig))
                                            {
                                                var rangeConfig = JsonSerializer.Deserialize<RangeConfigModel>(config.RangeConfig);
                                                var range = rangeConfig?.ranges?
                                                    .FirstOrDefault(r => data.Quantity >= r.from && data.Quantity <= r.to);
                                                if (range != null)
                                                    calculatedQuantity = range.value;
                                            }
                                            break;
                                    }
                                }
                            }
                            else
                            {
                                switch (config.Mode)
                                {
                                    case "Fixed":
                                        if (int.TryParse(config.Value, out var fqv))
                                            calculatedQuantity = fqv;
                                        break;

                                    case "Percentage":
                                        if (decimal.TryParse(config.Value, out var percent))
                                            calculatedQuantity = (int)Math.Round((double)(data.Quantity * percent) / 100);
                                        break;

                                    case "Range":
                                        if (!string.IsNullOrEmpty(config.RangeConfig))
                                        {
                                            var rangeConfig = JsonSerializer.Deserialize<RangeConfigModel>(config.RangeConfig);
                                            var range = rangeConfig?.ranges?
                                                .FirstOrDefault(r => data.Quantity >= r.from && data.Quantity <= r.to);
                                            if (range != null)
                                                calculatedQuantity = range.value;
                                        }
                                        break;
                                }
                            }

                            int innerCount = innerCapacity > 0
                                ? (int)Math.Ceiling((double)calculatedQuantity / innerCapacity)
                                : 0;

                            int outerCount = outerCapacity > 0
                                ? (int)Math.Ceiling((double)calculatedQuantity / outerCapacity)
                                : 0;

                            var extraReportItem = new ExtraEnvelopeReportDto
                            {
                                CatchNo = data.CatchNo,
                                NodalCode = data.NodalCode,
                                ExtraId = config.ExtraType,
                                Quantity = calculatedQuantity,
                                InnerEnvelope = innerCount.ToString(),
                                OuterEnvelope = outerCount.ToString()
                            };
                            envelopesToAdd.Add(extraReportItem);

                            // 1. Post entry in CenterList (centerlist table) dynamically from ExtraType table
                            string extraCenterName = !string.IsNullOrWhiteSpace(extraTypeName)
                                ? extraTypeName
                                : $"Extra {config.ExtraType}";

                            var centerExtraDict = new Dictionary<string, string>
                            {
                                { "CatchNo", data.CatchNo ?? "" },
                                { "Quantity", calculatedQuantity.ToString() },
                                { "NRQuantity", calculatedQuantity.ToString() },
                                { "InnerEnvelope", innerCount.ToString() },
                                { "OuterEnvelope", outerCount.ToString() }
                            };

                            var matchingCenterForSave = existingCenters.FirstOrDefault(c => c.NRDataId == data.NrDataId && (!string.IsNullOrWhiteSpace(data.NodalCode) && c.NodalCode == data.NodalCode))
                                ?? existingCenters.FirstOrDefault(c => c.NRDataId == data.NrDataId);

                            double extraNodalSort = matchingCenterForSave?.NodalSort ?? 0;
                            double extraCenterSort = 0;
                            int extraRouteSort = matchingCenterForSave?.RouteSort ?? 0;

                            switch (config.ExtraType)
                            {
                                case 1:
                                    extraCenterSort = 10000;
                                    break;
                                case 2:
                                    extraNodalSort = 10000;
                                    extraCenterSort = 100000;
                                    extraRouteSort = 10000;
                                    break;
                                case 3:
                                    extraNodalSort = 100000;
                                    extraCenterSort = 1000000;
                                    extraRouteSort = 100000;
                                    break;
                            }

                            // Ensure only 1 active record can exist for combination (ExtraId + NodalCode + NRDataId)
                            string targetNodal = (data.NodalCode ?? "").Trim();
                            var oldMatchingCenters = existingExtraCenters
                                .Where(c => c.ExtraId == config.ExtraType &&
                                            c.NRDataId == data.NrDataId &&
                                            c.Status &&
                                            ((c.NodalCode == null && string.IsNullOrEmpty(targetNodal)) ||
                                             (c.NodalCode ?? "").Trim().Equals(targetNodal, StringComparison.OrdinalIgnoreCase)))
                                .ToList();

                            foreach (var oldCenter in oldMatchingCenters)
                            {
                                oldCenter.Status = false;
                                _context.CenterList.Update(oldCenter);

                                var oldNewBreakages = existingNewBreakages
                                    .Where(b => b.CenterListId == oldCenter.Id && b.Status)
                                    .ToList();
                                foreach (var ob in oldNewBreakages)
                                {
                                    ob.Status = false;
                                    _context.NewEnvelopeBreakages.Update(ob);
                                }
                            }

                            var extraCenterRecord = new CenterList
                            {
                                ProjectId = ProjectId,
                                NRDataId = data.NrDataId,
                                CenterCode = extraCenterName,
                                NodalCode = data.NodalCode,
                                ExtraId = config.ExtraType,
                                Quantity = calculatedQuantity,
                                NRQuantity = calculatedQuantity,
                                CenterSort = extraCenterSort,
                                NodalSort = extraNodalSort,
                                RouteSort = extraRouteSort,
                                Status = true,
                                CenterData = JsonSerializer.Serialize(centerExtraDict)
                            };
                            newCentersToInsert.Add(extraCenterRecord);

                            // 2. Prepare envelope breaking for batch insert
                            var innerBreakdown = new Dictionary<string, string>();
                            if (innerCapacity > 0 && innerCount > 0)
                            {
                                innerBreakdown[$"E{innerCapacity}"] = innerCount.ToString();
                            }

                            var outerBreakdown = new Dictionary<string, string>();
                            if (outerCapacity > 0 && outerCount > 0)
                            {
                                outerBreakdown[$"E{outerCapacity}"] = outerCount.ToString();
                            }

                            pendingBreakages.Add((extraCenterRecord, data.NrDataId, innerBreakdown, outerBreakdown, outerCount));
                        }
                    }

                    // 1. Batch insert new centers and commit deactivations in 1 roundtrip
                    if (newCentersToInsert.Any())
                    {
                        await _context.CenterList.AddRangeAsync(newCentersToInsert);
                    }
                    await _context.SaveChangesAsync(); // Generated IDs now populated on extraCenterRecord.Id

                    // 2. Batch insert NewEnvelopeBreakage & EnvelopeBreakage
                    var newEnvBreakages = new List<NewEnvelopeBreakage>();
                    var envBreakages = new List<EnvelopeBreakage>();

                    foreach (var pb in pendingBreakages)
                    {
                        newEnvBreakages.Add(new NewEnvelopeBreakage
                        {
                            ProjectId = ProjectId,
                            NrDataId = pb.NrDataId,
                            CenterListId = pb.CenterRecord.Id,
                            InnerEnvelope = JsonSerializer.Serialize(pb.InnerBreakdown),
                            OuterEnvelope = JsonSerializer.Serialize(pb.OuterBreakdown),
                            TotalEnvelope = pb.TotalEnvelope,
                            Status = true
                        });

                        envBreakages.Add(new EnvelopeBreakage
                        {
                            ProjectId = ProjectId,
                            NrDataId = pb.NrDataId,
                            InnerEnvelope = JsonSerializer.Serialize(pb.InnerBreakdown),
                            OuterEnvelope = JsonSerializer.Serialize(pb.OuterBreakdown),
                            TotalEnvelope = pb.TotalEnvelope,
                            Status = true
                        });
                    }

                    if (newEnvBreakages.Any())
                    {
                        var targetCenterIds = newEnvBreakages.Select(b => b.CenterListId).Distinct().ToList();
                        var existingActiveForCenters = await _context.NewEnvelopeBreakages
                            .Where(b => b.ProjectId == ProjectId && targetCenterIds.Contains(b.CenterListId) && b.Status)
                            .ToListAsync();

                        if (existingActiveForCenters.Any())
                        {
                            foreach (var oldB in existingActiveForCenters)
                            {
                                oldB.Status = false;
                                _context.NewEnvelopeBreakages.Update(oldB);
                            }
                        }

                        await _context.NewEnvelopeBreakages.AddRangeAsync(newEnvBreakages);
                    }
                    if (envBreakages.Any())
                    {
                        await _context.EnvelopeBreakages.AddRangeAsync(envBreakages);
                    }

                    envelopesToUse = envelopesToAdd;
                }

                // Update steps for NrData1
                foreach (var nr in nrDataList)
                {
                    nr.Steps = PipelineNavigator.GetNextStep(PipelineNavigator.STEP_AWAITING_EXTRA, projectConfig?.Modules);
                    _context.NrData1.Update(nr);
                }

                await _context.SaveChangesAsync();

                // Generate Excel Report
                var extraConfigs = await _context.ExtraConfigurations
                    .Where(x => x.ProjectId == ProjectId)
                    .ToListAsync();

                var extraHeaders = new HashSet<string>();
                var innerKeys = new HashSet<string>();
                var outerKeys = new HashSet<string>();

                var allRows = new List<Dictionary<string, object>>();

                foreach (var extra in envelopesToUse)
                {
                    var baseRow = nrDataList.FirstOrDefault(x =>
                        x.CatchNo == extra.CatchNo);

                    var config = extraConfigs.FirstOrDefault(c => c.ExtraType == extra.ExtraId);
                    if (baseRow == null || (!isOnlyReport && config == null)) continue;

                    var matchingCenter = existingCenters.FirstOrDefault(c => c.NRDataId == baseRow.Id && (!string.IsNullOrWhiteSpace(extra.NodalCode) && c.NodalCode == extra.NodalCode))
                        ?? existingCenters.FirstOrDefault(c => c.NRDataId == baseRow.Id);

                    double baseNodalSort = matchingCenter?.NodalSort ?? 0;
                    int baseRouteSort = matchingCenter?.RouteSort ?? 0;

                    if (baseNodalSort == 0 && !string.IsNullOrWhiteSpace(baseRow.NRDatas))
                    {
                        try
                        {
                            var dyn = JsonSerializer.Deserialize<Dictionary<string, string>>(baseRow.NRDatas);
                            if (dyn != null)
                            {
                                if (dyn.TryGetValue("NodalSort", out var nsStr) && double.TryParse(nsStr, out var nsVal))
                                    baseNodalSort = nsVal;
                                if (baseRouteSort == 0 && dyn.TryGetValue("RouteSort", out var rsStr) && int.TryParse(rsStr, out var rsVal))
                                    baseRouteSort = rsVal;
                            }
                        }
                        catch { }
                    }

                    var dict = new Dictionary<string, object>();
                    dict["CatchNo"] = baseRow.CatchNo ?? "";
                    dict["CourseName"] = baseRow.CourseName ?? "";
                    dict["SubjectName"] = baseRow.SubjectName ?? "";
                    dict["Quantity"] = extra.Quantity;
                    dict["NodalCode"] = extra.NodalCode ?? "";
                    dict["ExtraId"] = extra.ExtraId;
                    dict["InnerEnvelope"] = extra.InnerEnvelope ?? "";
                    dict["OuterEnvelope"] = extra.OuterEnvelope ?? "";

                    // Set CenterCode
                    var matchingExtraCenter = await _context.CenterList
                        .Where(c => c.ProjectId == ProjectId && c.NRDataId == baseRow.Id && c.ExtraId == extra.ExtraId && c.Status)
                        .FirstOrDefaultAsync();

                    if (matchingExtraCenter != null && !string.IsNullOrWhiteSpace(matchingExtraCenter.CenterCode) &&
                        !matchingExtraCenter.CenterCode.Contains("Extra", StringComparison.OrdinalIgnoreCase))
                    {
                        dict["CenterCode"] = matchingExtraCenter.CenterCode;
                        dict["CenterSort"] = matchingExtraCenter.CenterSort;
                        dict["NodalSort"] = matchingExtraCenter.NodalSort > 0 ? matchingExtraCenter.NodalSort : baseNodalSort;
                        dict["RouteSort"] = matchingExtraCenter.RouteSort > 0 ? matchingExtraCenter.RouteSort : baseRouteSort;
                        dict["Remark"] = "Reserve Quantity";
                    }
                    else
                    {
                        switch (extra.ExtraId)
                        {
                            case 1:
                                dict["CenterCode"] = "Nodal Extra";
                                dict["NodalSort"] = baseNodalSort;
                                dict["CenterSort"] = 10000;
                                dict["RouteSort"] = baseRouteSort;
                                break;
                            case 2:
                                dict["CenterCode"] = "University Extra";
                                dict["NodalSort"] = 10000;
                                dict["CenterSort"] = 100000;
                                dict["RouteSort"] = 10000;
                                break;
                            case 3:
                                dict["CenterCode"] = "Office Extra";
                                dict["NodalSort"] = 100000;
                                dict["CenterSort"] = 1000000;
                                dict["RouteSort"] = 100000;
                                break;
                            default:
                                dict["CenterCode"] = "Extra";
                                break;
                        }
                    }

                    allRows.Add(dict);
                }

                var path = FileStorageHelper.GetProjectFolder(ProjectId);
                var distinctLots = (lotNo.HasValue && lotNo.Value > 0)
                    ? new List<int> { lotNo.Value }
                    : nrDataList.Where(r => r.EnvLotNo > 0).Select(r => r.EnvLotNo).Distinct().OrderBy(l => l).ToList();
                var lotStr = distinctLots.Any() ? string.Join("_", distinctLots) : "All";

                var fileName = uploadId.HasValue
                    ? $"ExtrasCalculation_NrData1_{lotStr}_v{uploadId}.xlsx"
                    : ReportVersionHelper.GetNextVersionFileName(path, $"ExtrasCalculation_NrData1_{lotStr}.xlsx");
                var filePath = Path.Combine(path, fileName);

                ExcelPackage.License.SetNonCommercialPersonal("Tools");
                using (var package = new ExcelPackage())
                {
                    var ws = package.Workbook.Worksheets.Add("Extra Envelope");
                    var headers = new List<string> { "CatchNo", "CourseName", "SubjectName", "CenterCode", "Quantity", "NodalCode", "ExtraId", "NodalSort", "CenterSort", "RouteSort", "InnerEnvelope", "OuterEnvelope" };
                    foreach (var rowDict in allRows)
                    {
                        foreach (var k in rowDict.Keys)
                        {
                            if (!headers.Contains(k))
                                headers.Add(k);
                        }
                    }

                    for (int i = 0; i < headers.Count; i++)
                        ws.Cells[1, i + 1].Value = headers[i];

                    int row = 2;
                    foreach (var rowDict in allRows)
                    {
                        for (int col = 0; col < headers.Count; col++)
                        {
                            rowDict.TryGetValue(headers[col], out object? val);
                            ws.Cells[row, col + 1].Value = val?.ToString();
                        }
                        row++;
                    }

                    if (ws.Dimension != null)
                        ws.Cells[ws.Dimension.Address].AutoFitColumns();
                    package.SaveAs(new FileInfo(filePath));
                }

                await ExcelReportHelper.RecordExcelReportAsync(
                    _context,
                    ProjectId,
                    3, // Module 3 (Extras Calculation)
                    uploadId ?? 1,
                    lotNo,
                    filePath,
                    true,
                    LogHelper.GetTriggeredBy(User)
                );

                return Ok(new
                {
                    message = isOnlyReport
                        ? "Report generated from existing data"
                        : "Calculated and report generated into centerlist and newenvelopebreakage",
                    data = envelopesToUse,
                    fileName = fileName
                });
            }
            catch (Exception ex)
            {
                await _loggerService.LogErrorAsync("Error creating ExtraEnvelope for NrData1", ex.ToString(), nameof(NrData1Controller));
                return StatusCode(500, "Internal server error: " + ex.Message);
            }
        }

        [HttpPost("ProcessEnvelopeBreaking")]
        public async Task<IActionResult> ProcessEnvelopeBreaking([FromQuery] int ProjectId, [FromQuery] int triggeredBy = 0, [FromQuery] bool skipReset = false, [FromQuery] int? lotNo = null, [FromQuery] string? catchNo = null, [FromQuery] bool bypassDispatch = false, [FromQuery] int? batchNo = null)
        {
            var procController = new EnvelopeBreakageProcessingController(_context, _loggerService, _apiSettings, _dispatchService)
            {
                ControllerContext = this.ControllerContext
            };
            return await procController.ProcessEnvelopeBreaking(ProjectId, triggeredBy, skipReset, lotNo, catchNo, bypassDispatch, batchNo);
        }

        [HttpPost("ProcessBoxBreaking")]
        public async Task<IActionResult> ProcessBoxBreaking([FromQuery] int ProjectId, [FromQuery] List<int>? LotNo = null, [FromQuery] bool skipReset = false, [FromQuery] bool bypassDispatch = false, [FromQuery] bool runBoth = false, [FromQuery] int? batchNo = null)
        {
            var boxController = new BoxBreakingProcessingController(_context, _loggerService, _apiSettings, _dispatchService)
            {
                ControllerContext = this.ControllerContext
            };
            return await boxController.ProcessBoxBreaking(ProjectId, LotNo ?? new List<int>(), skipReset, bypassDispatch, runBoth, batchNo);
        }

        [HttpPost("missing-data")]
        public async Task<ActionResult> SaveMissingData([FromBody] NRDatasController.MissingDataSaveRequest request, [FromQuery] int? lotNo = null)
        {
            var legacy = new NRDatasController(_context, _loggerService, _apiSettings, _dispatchService)
            {
                ControllerContext = this.ControllerContext
            };
            return await legacy.SaveMissingData(request, lotNo);
        }

        #region Helper Models
        public class NodalValueConfig
        {
            public string? NodalCodes { get; set; }
            public string? Value { get; set; }
        }

        public class RangeConfigModel
        {
            public List<RangeItem>? ranges { get; set; }
        }

        public class RangeItem
        {
            public int from { get; set; }
            public int to { get; set; }
            public int value { get; set; }
        }
        public class ExtraEnvelopeReportDto
        {
            public string? CatchNo { get; set; }
            public string? NodalCode { get; set; }
            public int ExtraId { get; set; }
            public int Quantity { get; set; }
            public string? InnerEnvelope { get; set; }
            public string? OuterEnvelope { get; set; }
        }
        #endregion

        private class CatchGroupData
        {
            public NrData1 NrData1 { get; set; } = null!;
            public Dictionary<string, string> UniqueExtraData { get; set; } = new(StringComparer.OrdinalIgnoreCase);
            public List<CenterList> Centers { get; set; } = new();
        }

        private static object? ConvertPropertyValue(PropertyInfo propInfo, string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return null;

            try
            {
                var targetType = Nullable.GetUnderlyingType(propInfo.PropertyType) ?? propInfo.PropertyType;

                if (targetType == typeof(bool))
                {
                    if (bool.TryParse(value, out bool bVal)) return bVal;
                    if (value.Equals("1") || value.Equals("true", StringComparison.OrdinalIgnoreCase)) return true;
                    if (value.Equals("0") || value.Equals("false", StringComparison.OrdinalIgnoreCase)) return false;
                    return null;
                }

                if (targetType == typeof(int))
                {
                    if (int.TryParse(value, out int iVal)) return iVal;
                    if (double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out double dVal)) return (int)dVal;
                    return null;
                }

                if (targetType == typeof(double))
                {
                    if (double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out double dVal)) return dVal;
                    return null;
                }

                if (targetType == typeof(DateTime))
                {
                    if (DateTime.TryParse(value, out DateTime dtVal)) return dtVal;
                    return null;
                }

                return Convert.ChangeType(value, targetType, CultureInfo.InvariantCulture);
            }
            catch
            {
                return null;
            }
        }

        private class EnvelopeTypeConfig
        {
            public string? Inner { get; set; }
            public string? Outer { get; set; }
        }

        private List<ExtraEnvelopes> ProcessExtraEnvelopes(
            string? catchNo,
            string? centerCode,
            int quantity,
            int nrQuantity,
            int projectId,
            List<ExtrasConfiguration> extraConfigs,
            Dictionary<string, int>? extraTypesByName = null)
        {
            var extraEnvelopesToAdd = new List<ExtraEnvelopes>();
            int? extraTypeId = null;
            if (!string.IsNullOrWhiteSpace(centerCode) && extraTypesByName != null && extraTypesByName.TryGetValue(centerCode.Trim(), out var foundId))
            {
                extraTypeId = foundId;
            }

            if (extraTypeId.HasValue)
            {
                var config = extraConfigs.FirstOrDefault(x => x.ExtraType == extraTypeId);
                if (config != null)
                {
                    EnvelopeTypeConfig? envelopeType = null;
                    if (!string.IsNullOrWhiteSpace(config.EnvelopeType))
                    {
                        try
                        {
                            envelopeType = JsonSerializer.Deserialize<EnvelopeTypeConfig>(config.EnvelopeType);
                        }
                        catch { }
                    }

                    int innerCapacity = envelopeType != null && !string.IsNullOrWhiteSpace(envelopeType.Inner) ? GetEnvelopeCapacity(envelopeType.Inner) : 1;
                    int outerCapacity = envelopeType != null && !string.IsNullOrWhiteSpace(envelopeType.Outer) ? GetEnvelopeCapacity(envelopeType.Outer) : 1;
                    int baseQty = nrQuantity > 0 ? nrQuantity : quantity;
                    int roundedQty = baseQty;

                    if (innerCapacity > 0)
                        roundedQty = (int)Math.Ceiling((double)baseQty / innerCapacity) * innerCapacity;
                    else if (outerCapacity > 0)
                        roundedQty = (int)Math.Ceiling((double)baseQty / outerCapacity) * outerCapacity;

                    extraEnvelopesToAdd.Add(new ExtraEnvelopes
                    {
                        ProjectId = projectId,
                        CatchNo = catchNo,
                        ExtraId = extraTypeId.Value,
                        Quantity = roundedQty,
                        InnerEnvelope = innerCapacity > 0 ? Math.Ceiling((double)roundedQty / innerCapacity).ToString() : "0",
                        OuterEnvelope = outerCapacity > 0 ? Math.Ceiling((double)roundedQty / outerCapacity).ToString() : "0",
                        Status = 1
                    });
                }
            }

            return extraEnvelopesToAdd;
        }

        private List<ExtraEnvelopes> GroupAndProcessExtras(
            List<ExtraEnvelopes> extras,
            int projectId,
            List<ExtrasConfiguration> extraConfigs)
        {
            var grouped = extras
                .GroupBy(e => new { e.ExtraId, e.CatchNo })
                .Select(g => new ExtraEnvelopes
                {
                    ProjectId = projectId,
                    ExtraId = g.Key.ExtraId,
                    CatchNo = g.Key.CatchNo,
                    Quantity = g.Sum(e => e.Quantity),
                    Status = 1
                })
                .ToList();

            foreach (var extra in grouped)
            {
                var config = extraConfigs.FirstOrDefault(x => x.ExtraType == extra.ExtraId);
                if (config != null)
                {
                    EnvelopeTypeConfig? envelopeType = null;
                    if (!string.IsNullOrWhiteSpace(config.EnvelopeType))
                    {
                        try
                        {
                            envelopeType = JsonSerializer.Deserialize<EnvelopeTypeConfig>(config.EnvelopeType);
                        }
                        catch { }
                    }

                    int innerCap = envelopeType != null && !string.IsNullOrWhiteSpace(envelopeType.Inner) ? GetEnvelopeCapacity(envelopeType.Inner) : 1;
                    int outerCap = envelopeType != null && !string.IsNullOrWhiteSpace(envelopeType.Outer) ? GetEnvelopeCapacity(envelopeType.Outer) : 1;

                    extra.InnerEnvelope = innerCap > 0 ? Math.Ceiling((double)extra.Quantity / innerCap).ToString() : "0";
                    extra.OuterEnvelope = outerCap > 0 ? Math.Ceiling((double)extra.Quantity / outerCap).ToString() : "0";
                }
            }

            return grouped;
        }

        private int GetEnvelopeCapacity(string envelopeCode)
        {
            if (string.IsNullOrWhiteSpace(envelopeCode))
                return 0;

            var numberPart = new string(envelopeCode.Where(char.IsDigit).ToArray());
            return int.TryParse(numberPart, out var capacity) ? capacity : 0;
        }

        // DELETE: api/NrData1/5
        [HttpDelete("{id}")]
        public async Task<IActionResult> DeleteNrData1(int id)
        {
            try
            {
                var nr1 = await _context.NrData1.FindAsync(id);
                if (nr1 == null)
                {
                    return NotFound();
                }

                nr1.Status = false;
                _context.NrData1.Update(nr1);

                // Update referencing CenterList rows to Status = false
                var centers = await _context.CenterList
                    .Where(c => c.NRDataId == nr1.Id && c.Status)
                    .ToListAsync();
                foreach (var c in centers)
                {
                    c.Status = false;
                    _context.CenterList.Update(c);
                }

                await _context.SaveChangesAsync();
                await _loggerService.LogEventAsync($"Soft deleted NrData1 record {id} and referencing CenterList records", "NrData1", LogHelper.GetTriggeredBy(User), nr1.ProjectId);
                return NoContent();
            }
            catch (Exception ex)
            {
                await _loggerService.LogErrorAsync("Error soft deleting NrData1 record", ex.Message, nameof(NrData1Controller));
                return StatusCode(500, "Internal Server Error");
            }
        }

        // DELETE: api/NrData1/DeleteByProject/5
        [HttpDelete("DeleteByProject/{projectId}")]
        public async Task<IActionResult> DeleteByProject(int projectId, [FromQuery] int? lotNo = null)
        {
            try
            {
                var query = _context.NrData1.Where(d => d.ProjectId == projectId);
                if (lotNo.HasValue)
                {
                    query = query.Where(d => d.LotNo == lotNo.Value);
                }
                var nr1List = await query.ToListAsync();
                if (!nr1List.Any())
                {
                    return NotFound($"No NrData1 records found for ProjectId {projectId}");
                }

                var nr1Ids = nr1List.Select(n => n.Id).ToList();

                foreach (var item in nr1List)
                {
                    item.Status = false;
                    _context.NrData1.Update(item);
                }

                var centers = await _context.CenterList
                    .Where(c => c.ProjectId == projectId && nr1Ids.Contains(c.NRDataId) && c.Status)
                    .ToListAsync();

                foreach (var c in centers)
                {
                    c.Status = false;
                    _context.CenterList.Update(c);
                }

                await _context.SaveChangesAsync();
                await _loggerService.LogEventAsync($"Soft deleted all NrData1 and CenterList records for ProjectId {projectId}", "NrData1", LogHelper.GetTriggeredBy(User), projectId);
                return Ok(new { message = "Project NrData1 and CenterList records soft deleted successfully.", projectId, count = nr1List.Count });
            }
            catch (Exception ex)
            {
                await _loggerService.LogErrorAsync("Error soft deleting project NrData1 records", ex.Message, nameof(NrData1Controller));
                return StatusCode(500, "Internal Server Error");
            }
        }

        // DELETE: api/NrData1/DeleteCatchNo/5/Catch123
        [HttpDelete("DeleteCatchNo/{projectId}/{*catchNo}")]
        [HttpDelete("DeleteCatchNo/{projectId}")]
        public async Task<IActionResult> DeleteCatchNo(int projectId, string? catchNo, [FromQuery] string? queryCatchNo = null)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(catchNo))
                    catchNo = queryCatchNo;

                if (string.IsNullOrWhiteSpace(catchNo))
                    return BadRequest("Catch number cannot be empty");

                var nr1List = await _context.NrData1
                    .Where(d => d.ProjectId == projectId && d.CatchNo == catchNo)
                    .ToListAsync();

                if (!nr1List.Any())
                    return NotFound($"Catch number {catchNo} not found for project {projectId}");

                var nr1Ids = nr1List.Select(n => n.Id).ToList();

                foreach (var item in nr1List)
                {
                    item.Status = false;
                    _context.NrData1.Update(item);
                }

                var centers = await _context.CenterList
                    .Where(c => c.ProjectId == projectId && nr1Ids.Contains(c.NRDataId) && c.Status)
                    .ToListAsync();

                foreach (var c in centers)
                {
                    c.Status = false;
                    _context.CenterList.Update(c);
                }

                await _context.SaveChangesAsync();
                await _loggerService.LogEventAsync($"Soft deleted CatchNo {catchNo} in NrData1 and CenterList", "NrData1", LogHelper.GetTriggeredBy(User), projectId);
                return Ok(new { message = $"Catch number {catchNo} soft deleted successfully in NrData1 and CenterList." });
            }
            catch (Exception ex)
            {
                await _loggerService.LogErrorAsync("Error deleting CatchNo in NrData1", ex.Message, nameof(NrData1Controller));
                return StatusCode(500, "Internal Server Error");
            }
        }
    }
}
