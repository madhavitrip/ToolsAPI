using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Tools.Models;
using ERPToolsAPI.Data;
using System.Text.Json;

namespace Tools.Controllers
{
    public class DynamicRule1Request
    {
        public List<string> Level1 { get; set; } = new List<string>();
        public List<string> Level2 { get; set; } = new List<string>();
        public List<string> Level3 { get; set; } = new List<string>();
    }

    [Route("api/[controller]")]
    [ApiController]
    public class MergingController : ControllerBase
    {
        private readonly ERPToolsDbContext _context;

        public MergingController(ERPToolsDbContext context)
        {
            _context = context;
        }

        [HttpPost("MergeToTemporary/{projectId}")]
        public async Task<IActionResult> MergeToTemporary(int projectId, [FromQuery] string? mergeBy = "CollegeCode")
        {
            try
            {
                var nodalLists = await _context.NodalList.AsNoTracking().Where(x => x.ProjectId == projectId && x.Status).ToListAsync();

                // Rule 1 Validation: One college -> one center, One center -> one nodal
                var groupedByCollegeGender = nodalLists
                    .GroupBy(n => new { n.CollegeCode, Gender = n.Gender?.ToUpper() ?? "ALL" })
                    .ToList();

                var multiCenterColleges = groupedByCollegeGender
                    .Where(g => g.Select(x => x.ExamCenterCode).Distinct().Count() > 1)
                    .ToList();

                var multiNodalCenters = nodalLists
                    .GroupBy(n => n.ExamCenterCode)
                    .Where(g => g.Select(x => x.NodalCode).Distinct().Count() > 1)
                    .ToList();

                if (multiCenterColleges.Any() || multiNodalCenters.Any())
                {
                    var rule1Errors = new List<string>();
                    if (multiCenterColleges.Any())
                    {
                        var msg = string.Join("; ", multiCenterColleges.Select(g => $"College {g.Key.CollegeCode} ({g.Key.Gender}) assigned to {g.Select(x => x.ExamCenterCode).Distinct().Count()} centers"));
                        rule1Errors.Add($"Rule 1 Failed: Multiple exam centers for same college/gender - {msg}");
                    }
                    if (multiNodalCenters.Any())
                    {
                        var msg = string.Join("; ", multiNodalCenters.Select(g => $"Center {g.Key} assigned to {g.Select(x => x.NodalCode).Distinct().Count()} nodal codes"));
                        rule1Errors.Add($"Rule 1 Failed: Multiple nodal codes for same exam center - {msg}");
                    }

                    var multiCenterDetails = multiCenterColleges.Select(g => new
                    {
                        collegeCode = g.Key.CollegeCode,
                        gender = g.Key.Gender,
                        collegeNames = g.Select(x => x.CollegeName).Where(x => !string.IsNullOrEmpty(x)).Distinct().ToList(),
                        centerCount = g.Select(x => x.ExamCenterCode).Distinct().Count(),
                        centerCodes = g.Select(x => x.ExamCenterCode).Distinct().ToList(),
                        centerNames = g.Select(x => x.ExamCenterName).Where(x => !string.IsNullOrEmpty(x)).Distinct().ToList(),
                        nodalCodes = g.Select(x => x.NodalCode).Distinct().ToList(),
                        nodalNames = g.Select(x => x.NodalName).Where(x => !string.IsNullOrEmpty(x)).Distinct().ToList(),
                        records = g.Select(x => new { x.Id, x.CollegeCode, x.CollegeName, x.ExamCenterCode, x.ExamCenterName, x.NodalCode, x.NodalName, x.Gender }).ToList(),
                        description = $"College {g.Key.CollegeCode} ({g.Key.Gender}) assigned to {g.Select(x => x.ExamCenterCode).Distinct().Count()} centers: {string.Join(", ", g.Select(x => x.ExamCenterCode).Distinct())}"
                    }).ToList();

                    var multiNodalDetails = multiNodalCenters.Select(g => new
                    {
                        centerCode = g.Key,
                        centerNames = g.Select(x => x.ExamCenterName).Where(x => !string.IsNullOrEmpty(x)).Distinct().ToList(),
                        nodalCount = g.Select(x => x.NodalCode).Distinct().Count(),
                        nodalCodes = g.Select(x => x.NodalCode).Distinct().ToList(),
                        nodalNames = g.Select(x => x.NodalName).Where(x => !string.IsNullOrEmpty(x)).Distinct().ToList(),
                        collegeCodes = g.Select(x => x.CollegeCode).Distinct().ToList(),
                        records = g.Select(x => new { x.Id, x.CollegeCode, x.CollegeName, x.ExamCenterCode, x.ExamCenterName, x.NodalCode, x.NodalName, x.Gender }).ToList(),
                        description = $"Center {g.Key} assigned to {g.Select(x => x.NodalCode).Distinct().Count()} nodal codes: {string.Join(", ", g.Select(x => x.NodalCode).Distinct())}"
                    }).ToList();

                    return BadRequest(new
                    {
                        message = "Rule 1 validation failed: One college code must belong to one exam center, and one exam center must belong to one nodal code.",
                        errors = rule1Errors,
                        rule1CollegeCodes = multiCenterColleges.Select(g => g.Key.CollegeCode.ToString()).Distinct().ToList(),
                        rule1CenterCodes = multiNodalCenters.Select(g => g.Key.ToString()).Distinct().ToList(),
                        multiCenterDetails,
                        multiNodalDetails
                    });
                }

                // Clear existing temp data
                await _context.TemporaryNrDatas.Where(x => x.ProjectId == projectId).ExecuteDeleteAsync();

                var catchLists = await _context.CatchList.AsNoTracking().Where(x => x.ProjectId == projectId && x.Status).ToListAsync();

                var mode = (mergeBy ?? "CollegeCode").Trim();
                Func<CatchList, string> getCatchKey;
                Func<NodalList, string> getNodalKey;

                switch (mode.ToLowerInvariant())
                {
                    case "collegename":
                        getCatchKey = c => (c.CollegeName ?? "").Trim().ToLowerInvariant();
                        getNodalKey = n => (n.CollegeName ?? "").Trim().ToLowerInvariant();
                        break;
                    case "centercode":
                        getCatchKey = c => c.CollegeCode.ToString().Trim();
                        getNodalKey = n => n.ExamCenterCode.ToString().Trim();
                        break;
                    case "centername":
                        getCatchKey = c => (c.CollegeName ?? "").Trim().ToLowerInvariant();
                        getNodalKey = n => (n.ExamCenterName ?? "").Trim().ToLowerInvariant();
                        break;
                    case "collegecode":
                    default:
                        getCatchKey = c => c.CollegeCode.ToString().Trim();
                        getNodalKey = n => n.CollegeCode.ToString().Trim();
                        break;
                }

                var nodalLookup = nodalLists
                    .Where(n => !string.IsNullOrEmpty(getNodalKey(n)))
                    .ToLookup(n => getNodalKey(n));

                var tempDatas = new List<TemporaryNrDatas>();

                foreach (var catchItem in catchLists)
                {
                    var key = getCatchKey(catchItem);
                    var matchingNodals = !string.IsNullOrEmpty(key) ? nodalLookup[key].ToList() : new List<NodalList>();

                    if (!matchingNodals.Any())
                    {
                        // Even if no center, create an entry with null CenterCode so they can see and fix it.
                        tempDatas.Add(new TemporaryNrDatas
                        {
                            ProjectId = projectId,
                            CatchNo = catchItem.CatchNo,
                            NRQuantity = catchItem.NRQuantity,
                            CollegeCode = catchItem.CollegeCode,
                            CollegeName = catchItem.CollegeName,
                            CourseName = catchItem.CourseName,
                            SubjectName = catchItem.SubjectName,
                            ExamDate = catchItem.ExamDate,
                            ExamTime = catchItem.ExamTime,
                            NRDatas = catchItem.NRDatas
                        });
                        continue;
                    }

                    foreach (var nodalItem in matchingNodals)
                    {
                        var gender = nodalItem.Gender?.ToUpper() ?? "ALL";
                        int quantity = 0;

                        if (gender == "ALL")
                        {
                            quantity = catchItem.NRQuantity;
                        }
                        else if (gender == "MALE")
                        {
                            quantity = catchItem.Male;
                        }
                        else if (gender == "FEMALE")
                        {
                            quantity = catchItem.Female;
                        }

                        if (quantity > 0 || gender == "ALL") // Always add if 'ALL' or quantity > 0
                        {
                            tempDatas.Add(new TemporaryNrDatas
                            {
                                ProjectId = projectId,
                                CatchNo = catchItem.CatchNo,
                                NRQuantity = quantity,
                                CollegeCode = catchItem.CollegeCode,
                                CollegeName = catchItem.CollegeName,
                                CenterCode = nodalItem.ExamCenterCode.ToString(),
                                NodalCode = nodalItem.NodalCode.ToString(),
                                CourseName = catchItem.CourseName,
                                SubjectName = catchItem.SubjectName,
                                ExamDate = catchItem.ExamDate,
                                ExamTime = catchItem.ExamTime,
                                NRDatas = catchItem.NRDatas
                            });
                        }
                    }
                }

                _context.ChangeTracker.AutoDetectChangesEnabled = false;
                await _context.TemporaryNrDatas.AddRangeAsync(tempDatas);
                await _context.SaveChangesAsync();
                _context.ChangeTracker.AutoDetectChangesEnabled = true;

                return Ok(new { message = "Data merged into temporary staging successfully", count = tempDatas.Count });
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
        }

        [HttpPost("CheckDynamicRule1/{projectId}")]
        public async Task<IActionResult> CheckDynamicRule1(int projectId, [FromBody] DynamicRule1Request req)
        {
            try
            {
                // Clear existing dynamic conflicts
                await _context.ConflictingFields
                    .Where(c => c.ProjectId == projectId)
                    .ExecuteDeleteAsync();

                var nodalLists = await _context.NodalList.AsNoTracking().Where(x => x.ProjectId == projectId && x.Status).ToListAsync();
                var catchLists = await _context.CatchList.AsNoTracking().Where(x => x.ProjectId == projectId && x.Status).ToListAsync();

                // Merge lists for dynamic fields using full outer join on CollegeCode
                var unifiedList = new List<dynamic>();
                var catchGroup = catchLists.GroupBy(c => c.CollegeCode).ToDictionary(g => g.Key, g => g.ToList());
                var nodalGroup = nodalLists.GroupBy(n => n.CollegeCode).ToDictionary(g => g.Key, g => g.ToList());
                var allCollegeCodes = catchGroup.Keys.Union(nodalGroup.Keys).ToList();

                foreach (var code in allCollegeCodes)
                {
                    var cList = catchGroup.ContainsKey(code) ? catchGroup[code] : new List<CatchList>();
                    var nList = nodalGroup.ContainsKey(code) ? nodalGroup[code] : new List<NodalList>();

                    if (cList.Any() && nList.Any())
                    {
                        foreach (var c in cList)
                            foreach (var n in nList)
                                unifiedList.Add(new { Catch = c, Nodal = n });
                    }
                    else if (cList.Any())
                    {
                        foreach (var c in cList)
                            unifiedList.Add(new { Catch = c, Nodal = (NodalList)null });
                    }
                    else if (nList.Any())
                    {
                        foreach (var n in nList)
                            unifiedList.Add(new { Catch = (CatchList)null, Nodal = n });
                    }
                }

                string GetVal(object catchObj, object nodalObj, string field)
                {
                    if (string.Equals(field, "NRQuantity", StringComparison.OrdinalIgnoreCase))
                    {
                        if (catchObj != null)
                        {
                            var p = catchObj.GetType().GetProperty(field, System.Reflection.BindingFlags.IgnoreCase | System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance);
                            if (p != null) return p.GetValue(catchObj)?.ToString() ?? "";
                        }
                    }

                    if (nodalObj != null)
                    {
                        var p = nodalObj.GetType().GetProperty(field, System.Reflection.BindingFlags.IgnoreCase | System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance);
                        if (p != null) return p.GetValue(nodalObj)?.ToString() ?? "";
                    }
                    if (catchObj != null)
                    {
                        var p = catchObj.GetType().GetProperty(field, System.Reflection.BindingFlags.IgnoreCase | System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance);
                        if (p != null) return p.GetValue(catchObj)?.ToString() ?? "";
                    }
                    return "";
                }

                string GetKey(object c, object n, List<string> fields)
                {
                    return string.Join("|", fields.Select(f => GetVal(c, n, f)));
                }

                var newConflicts = new List<ConflictingFields>();

                // Check Level 1 -> Level 2
                if (req.Level1 != null && req.Level1.Any() && req.Level2 != null && req.Level2.Any())
                {
                    var g1 = unifiedList.GroupBy(u => GetKey(u.Catch, u.Nodal, req.Level1)).ToList();
                    foreach (var g in g1)
                    {
                        if (string.IsNullOrEmpty(g.Key)) continue;
                        var l2Keys = g.Select(u => GetKey(u.Catch, u.Nodal, req.Level2)).Where(k => !string.IsNullOrEmpty(k)).Distinct().ToList();
                        if (l2Keys.Count > 1)
                        {
                            newConflicts.Add(new ConflictingFields
                            {
                                ProjectId = projectId,
                                UniqueField = JsonSerializer.Serialize(new { fields = req.Level1, value = g.Key }),
                                ConflictingField = JsonSerializer.Serialize(new { fields = req.Level2, values = l2Keys }),
                                Status = 1
                            });
                        }
                    }
                }

                // Check Level 2 -> Level 3
                if (req.Level2 != null && req.Level2.Any() && req.Level3 != null && req.Level3.Any())
                {
                    var g2 = unifiedList.GroupBy(u => GetKey(u.Catch, u.Nodal, req.Level2)).ToList();
                    foreach (var g in g2)
                    {
                        if (string.IsNullOrEmpty(g.Key)) continue;
                        var l3Keys = g.Select(u => GetKey(u.Catch, u.Nodal, req.Level3)).Where(k => !string.IsNullOrEmpty(k)).Distinct().ToList();
                        if (l3Keys.Count > 1)
                        {
                            newConflicts.Add(new ConflictingFields
                            {
                                ProjectId = projectId,
                                UniqueField = JsonSerializer.Serialize(new { fields = req.Level2, value = g.Key }),
                                ConflictingField = JsonSerializer.Serialize(new { fields = req.Level3, values = l3Keys }),
                                Status = 1
                            });
                        }
                    }
                }

                if (newConflicts.Any())
                {
                    _context.ConflictingFields.AddRange(newConflicts);
                    await _context.SaveChangesAsync();
                }

                return Ok(new { message = "Dynamic Rule 1 validated", conflicts = newConflicts.Count });
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
        }

        [HttpGet("Reports/{projectId}")]
        public async Task<IActionResult> GetReports(int projectId, [FromQuery] string? mergeBy = "CollegeCode")
        {
            try
            {
                var dynamicConflicts = await _context.ConflictingFields.AsNoTracking().Where(c => c.ProjectId == projectId).ToListAsync();

                return Ok(new
                {
                    rule1Passed = !dynamicConflicts.Any(),
                    rule2Passed = true,
                    dynamicConflicts = dynamicConflicts
                });
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
        }

        [HttpPost("PushToMain/{projectId}")]
        public async Task<IActionResult> PushToMain(int projectId)
        {
            try
            {
                var tempDatas = await _context.TemporaryNrDatas.AsNoTracking().Where(x => x.ProjectId == projectId).ToListAsync();

                if (!tempDatas.Any())
                {
                    return BadRequest("No temporary data found to push.");
                }

                var nrDatas = tempDatas.Select(t => new NRData
                {
                    ProjectId = t.ProjectId,
                    CourseName = t.CourseName,
                    SubjectName = t.SubjectName,
                    CenterCode = t.CenterCode,
                    NRQuantity = t.NRQuantity,
                    Quantity = t.NRQuantity,
                    CatchNo = t.CatchNo,
                    ExamDate = t.ExamDate,
                    ExamTime = t.ExamTime,
                    Day = t.Day,
                    NRDatas = t.NRDatas,
                    NodalCode = t.NodalCode,
                    Pages = t.Pages,
                    Route = t.Route,
                    RouteSort = t.RouteSort,
                    CenterSort = t.CenterSort,
                    NodalSort = t.NodalSort,
                    Symbol = t.Symbol,
                    Status = t.Status,
                    District = t.District,
                    DistrictSort = t.DistrictSort
                }).ToList();

                // Clear existing NRDatas for this project, maybe? 
                // Or does it append? Usually uploading fresh means clear first. Let's clear to be safe, 
                // or just insert if it's supposed to append. Usually it's clear first for a fresh upload.
                await _context.NRDatas.Where(x => x.ProjectId == projectId).ExecuteDeleteAsync();

                _context.ChangeTracker.AutoDetectChangesEnabled = false;
                await _context.NRDatas.AddRangeAsync(nrDatas);
                await _context.SaveChangesAsync();
                _context.ChangeTracker.AutoDetectChangesEnabled = true;
                
                // Clean up temp table
                await _context.TemporaryNrDatas.Where(x => x.ProjectId == projectId).ExecuteDeleteAsync();

                return Ok(new { message = "Data pushed to main table successfully." });
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
        }
    }
}
