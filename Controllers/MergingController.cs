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
        public async Task<IActionResult> MergeToTemporary(int projectId)
        {
            try
            {
                // Clear existing temp data
                await _context.TemporaryNrDatas.Where(x => x.ProjectId == projectId).ExecuteDeleteAsync();

                var catchLists = await _context.CatchList.AsNoTracking().Where(x => x.ProjectId == projectId).ToListAsync();
                var nodalLists = await _context.NodalList.AsNoTracking().Where(x => x.ProjectId == projectId).ToListAsync();
                var nodalLookup = nodalLists.ToLookup(n => n.CollegeCode);

                var tempDatas = new List<TemporaryNrDatas>();

                foreach (var catchItem in catchLists)
                {
                    var matchingNodals = nodalLookup[catchItem.CollegeCode].ToList();

                    if (!matchingNodals.Any())
                    {
                        // Even if no center, we might want to still stage it or leave it out? 
                        // The report will flag it if it's missing from NodalList.
                        // Let's create an entry with null CenterCode so they can see and fix it.
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
                            // Transgender logic skipped for now based on user's request
                        }
                        else if (gender == "FEMALE")
                        {
                            quantity = catchItem.Female;
                            // Transgender logic skipped for now based on user's request
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
        public async Task<IActionResult> GetReports(int projectId)
        {
            try
            {
                var catchLists = await _context.CatchList.AsNoTracking().Where(x => x.ProjectId == projectId).ToListAsync();
                var nodalLists = await _context.NodalList.AsNoTracking().Where(x => x.ProjectId == projectId).ToListAsync();

                var errors = new List<string>();

                // Rule 1: All college code is assigned a centre
                var catchCollegeCodes = catchLists.Select(c => c.CollegeCode).Distinct().ToList();
                var nodalCollegeCodes = nodalLists.Select(n => n.CollegeCode).Distinct().ToList();
                
                var unassignedColleges = catchCollegeCodes.Except(nodalCollegeCodes).ToList();
                if (unassignedColleges.Any())
                {
                    errors.Add($"Rule 1 Failed: College codes missing from Nodal List - {string.Join(", ", unassignedColleges)}");
                }

                // Rule 2: Every College_Code OR {College Code-Gender} is assigned to ONE Exam Centre-Every Exam Centre is assigned to ONE Nodal
                var groupedByCollegeGender = nodalLists
                    .GroupBy(n => new { n.CollegeCode, Gender = n.Gender?.ToUpper() ?? "ALL" })
                    .ToList();
                
                var multiCenterColleges = groupedByCollegeGender
                    .Where(g => g.Select(x => x.ExamCenterCode).Distinct().Count() > 1)
                    .ToList();

                if (multiCenterColleges.Any())
                {
                    var msg = string.Join("; ", multiCenterColleges.Select(g => $"College {g.Key.CollegeCode} ({g.Key.Gender}) assigned to {g.Select(x => x.ExamCenterCode).Distinct().Count()} centers"));
                    errors.Add($"Rule 2 Failed: Multiple exam centers for same college/gender - {msg}");
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
                    errors.Add($"Rule 2 Failed: Multiple nodal codes for same exam center - {msg}");
                }

                var affectedCollegeCodes = multiCenterColleges.Select(g => g.Key.CollegeCode).ToList();
                var affectedCollegeCodes2 = nodalLists.Where(n => multiNodalCenters.Select(g => g.Key).Contains(n.ExamCenterCode)).Select(n => n.CollegeCode).ToList();
                var allAffectedCollegeCodes = affectedCollegeCodes.Union(affectedCollegeCodes2).Distinct().ToList();
                var rule2CatchNos = catchLists.Where(c => allAffectedCollegeCodes.Contains(c.CollegeCode)).Select(c => c.CatchNo).Distinct().ToList();

                return Ok(new
                {
                    allCollegeCodesAssigned = !unassignedColleges.Any(),
                    rule2Passed = !multiCenterColleges.Any() && !multiNodalCenters.Any(),
                    errors,
                    rule1Colleges = catchLists.Where(c => unassignedColleges.Contains(c.CollegeCode)).Select(c => c.CatchNo).Distinct().ToList(),
                    rule2Colleges = rule2CatchNos,
                    rule1CollegeCodes = unassignedColleges.Select(c => c.ToString()).ToList(),
                    rule2CollegeCodes = multiCenterColleges.Select(g => g.Key.CollegeCode.ToString()).Distinct().ToList(),
                    rule2CenterCodes = multiNodalCenters.Select(g => g.Key.ToString()).Distinct().ToList()
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
