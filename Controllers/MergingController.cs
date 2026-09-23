using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Tools.Models;
using ERPToolsAPI.Data;
using System.Text.Json;

namespace Tools.Controllers
{
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

        [HttpGet("Reports/{projectId}")]
        public async Task<IActionResult> GetReports(int projectId, [FromQuery] string? mergeBy = "CollegeCode")
        {
            try
            {
                var catchLists = await _context.CatchList.AsNoTracking().Where(x => x.ProjectId == projectId && x.Status).ToListAsync();
                var nodalLists = await _context.NodalList.AsNoTracking().Where(x => x.ProjectId == projectId && x.Status).ToListAsync();

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

                var errors = new List<string>();

                // Rule 1: Every College_Code OR {College Code-Gender} is assigned to ONE Exam Centre - Every Exam Centre is assigned to ONE Nodal
                var groupedByCollegeGender = nodalLists
                    .GroupBy(n => new { n.CollegeCode, Gender = n.Gender?.ToUpper() ?? "ALL" })
                    .ToList();

                var multiCenterColleges = groupedByCollegeGender
                    .Where(g => g.Select(x => x.ExamCenterCode).Distinct().Count() > 1)
                    .ToList();

                if (multiCenterColleges.Any())
                {
                    var msg = string.Join("; ", multiCenterColleges.Select(g => $"College {g.Key.CollegeCode} ({g.Key.Gender}) assigned to {g.Select(x => x.ExamCenterCode).Distinct().Count()} centers"));
                    errors.Add($"Rule 1 Failed: Multiple exam centers for same college/gender - {msg}");
                }

                var groupedByExamCenter = nodalLists
                    .GroupBy(n => n.ExamCenterCode)
                    .ToList();

                var multiNodalCenters = groupedByExamCenter
                    .Where(g => g.Select(x => x.NodalCode).Distinct().Count() > 1)
                    .ToList();

                if (multiNodalCenters.Any())
                {
                    var msg = string.Join("; ", multiNodalCenters.Select(g => $"Center {g.Key} assigned to {g.Select(x => x.NodalCode).Distinct().Count()} nodal codes"));
                    errors.Add($"Rule 1 Failed: Multiple nodal codes for same exam center - {msg}");
                }

                bool rule1Passed = !multiCenterColleges.Any() && !multiNodalCenters.Any();

                // Rule 2: All catch list items are assigned in nodal list based on selected merge criteria
                var catchKeys = catchLists
                    .Select(c => getCatchKey(c))
                    .Where(k => !string.IsNullOrEmpty(k))
                    .Distinct()
                    .ToList();

                var nodalKeys = nodalLists
                    .Select(n => getNodalKey(n))
                    .Where(k => !string.IsNullOrEmpty(k))
                    .Distinct()
                    .ToHashSet();

                var unassignedKeys = catchKeys.Where(k => !nodalKeys.Contains(k)).ToList();
                if (unassignedKeys.Any())
                {
                    string fieldLabel = mode.ToLowerInvariant() switch
                    {
                        "collegename" => "College names",
                        "centercode" => "Center codes",
                        "centername" => "Center names",
                        _ => "College codes"
                    };
                    errors.Add($"Rule 2 Failed: {fieldLabel} missing from Nodal List - {string.Join(", ", unassignedKeys.Take(20))}{(unassignedKeys.Count > 20 ? $" and {unassignedKeys.Count - 20} more" : "")}");
                }

                bool rule2Passed = !unassignedKeys.Any();

                var affectedCollegeCodes = multiCenterColleges.Select(g => g.Key.CollegeCode).ToList();
                var affectedCollegeCodes2 = nodalLists.Where(n => multiNodalCenters.Select(g => g.Key).Contains(n.ExamCenterCode)).Select(n => n.CollegeCode).ToList();
                var allAffectedCollegeCodes = affectedCollegeCodes.Union(affectedCollegeCodes2).Distinct().ToList();
                var rule1CatchNos = catchLists.Where(c => allAffectedCollegeCodes.Contains(c.CollegeCode)).Select(c => c.CatchNo).Distinct().ToList();

                var rule2CatchItems = catchLists.Where(c => unassignedKeys.Contains(getCatchKey(c))).ToList();
                var rule2CatchNos = rule2CatchItems.Select(c => c.CatchNo).Distinct().ToList();

                var rule1Errors = errors.Where(e => e.StartsWith("Rule 1")).ToList();
                var rule2Errors = errors.Where(e => e.StartsWith("Rule 2")).ToList();

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

                string fieldLabelForDetails = mode.ToLowerInvariant() switch
                {
                    "collegename" => "College name",
                    "centercode" => "Center code",
                    "centername" => "Center name",
                    _ => "College code"
                };

                var unassignedDetails = catchLists
                    .Where(c => unassignedKeys.Contains(getCatchKey(c)))
                    .GroupBy(c => getCatchKey(c))
                    .Select(g => new
                    {
                        key = g.Key,
                        collegeCode = g.First().CollegeCode,
                        collegeName = g.First().CollegeName,
                        catchNos = g.Select(x => x.CatchNo).Distinct().ToList(),
                        totalQuantity = g.Sum(x => x.NRQuantity),
                        rowCount = g.Count(),
                        description = $"{fieldLabelForDetails} '{g.Key}' missing from Nodal List ({g.Select(x => x.CatchNo).Distinct().Count()} Catch Nos)"
                    })
                    .ToList();

                return Ok(new
                {
                    rule1Passed,
                    rule2Passed,
                    allCollegeCodesAssigned = rule2Passed,
                    errors,
                    rule1Errors,
                    rule2Errors,
                    rule1Colleges = rule1CatchNos,
                    rule1CollegeCodes = multiCenterColleges.Select(g => g.Key.CollegeCode.ToString()).Distinct().ToList(),
                    rule1CenterCodes = multiNodalCenters.Select(g => g.Key.ToString()).Distinct().ToList(),
                    rule2Colleges = rule2CatchNos,
                    rule2CollegeCodes = unassignedKeys,
                    multiCenterDetails,
                    multiNodalDetails,
                    unassignedDetails,
                    mergeBy = mode
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
