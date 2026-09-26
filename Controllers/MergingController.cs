using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Tools.Models;
using ERPToolsAPI.Data;
using System.Text.Json;
using System.Text;
using System.Globalization;

namespace Tools.Controllers
{
    public class DynamicRule1Request
    {
        public List<string> Level1 { get; set; } = new List<string>();
        public List<string> Level2 { get; set; } = new List<string>();
        public List<string> Level3 { get; set; } = new List<string>();
    }

    public class ResolveDynamicConflictDto
    {
        public int ProjectId { get; set; }
        public int ConflictId { get; set; }
        public string TargetField { get; set; } = "";
        public string TargetValue { get; set; } = "";
        public string TargetName { get; set; } = "";
        public string MatchField { get; set; } = "";
        public string MatchValue { get; set; } = "";
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

        private static string? ExtractCodeNumber(string? input)
        {
            if (string.IsNullOrWhiteSpace(input)) return null;
            var trimmed = input.Trim();
            if (trimmed == "0") return null;

            var match = System.Text.RegularExpressions.Regex.Match(trimmed, @"(?:\b(?:college\s+code|college|center|centre|nodal|col|code)\b\s*[:\-]?\s*)?(\d+)", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            if (match.Success && match.Groups[1].Success)
            {
                if (int.TryParse(match.Groups[1].Value, out int val) && val > 0)
                {
                    return val.ToString();
                }
                return match.Groups[1].Value;
            }

            var digitMatch = System.Text.RegularExpressions.Regex.Match(trimmed, @"\d+");
            if (digitMatch.Success)
            {
                if (int.TryParse(digitMatch.Value, out int val) && val > 0)
                {
                    return val.ToString();
                }
                return digitMatch.Value;
            }

            return trimmed.ToLowerInvariant();
        }

        [HttpPost("MergeToTemporary/{projectId}")]
        public async Task<IActionResult> MergeToTemporary(int projectId, [FromQuery] string? mergeBy = "CollegeCode")
        {
            try
            {
                // Check if any unresolved active dynamic conflicts exist in database
                var activeConflicts = await _context.ConflictingFields
                    .AsNoTracking()
                    .Where(c => c.ProjectId == projectId && c.Status == 1)
                    .AnyAsync();

                if (activeConflicts)
                {
                    return BadRequest(new
                    {
                        message = "Validation failed: Unresolved conflicts exist in Conflict Report. Please resolve them first.",
                        errors = new List<string> { "Please resolve all pending conflicts in the Conflict Report tab before generating preview." }
                    });
                }

                var nodalLists = await _context.NodalList.AsNoTracking().Where(x => x.ProjectId == projectId && x.Status).ToListAsync();

                static string GetEffectiveCollegeCode(NodalList n)
                {
                    if (n.CollegeCode != 0) return n.CollegeCode.ToString();
                    if (!string.IsNullOrEmpty(n.CollegeName))
                    {
                        var code = ExtractCodeNumber(n.CollegeName);
                        if (code != null && code != "0") return code;
                    }
                    return GetEffectiveCenterCode(n);
                }

                static string GetEffectiveCenterCode(NodalList n)
                {
                    if (n.ExamCenterCode != 0) return n.ExamCenterCode.ToString();
                    if (!string.IsNullOrEmpty(n.ExamCenterName))
                    {
                        var code = ExtractCodeNumber(n.ExamCenterName);
                        if (code != null) return code;
                        return n.ExamCenterName.Trim().ToLowerInvariant();
                    }
                    return $"center_{n.Id}";
                }

                // Rule 1 Validation: One college -> one center, One center -> one nodal
                var groupedByCollegeGender = nodalLists
                    .GroupBy(n => new { CollegeKey = GetEffectiveCollegeCode(n), Gender = n.Gender?.ToUpper() ?? "ALL" })
                    .ToList();

                var multiCenterColleges = groupedByCollegeGender
                    .Where(g => g.Select(x => GetEffectiveCenterCode(x)).Distinct().Count() > 1)
                    .ToList();

                var multiNodalCenters = nodalLists
                    .GroupBy(n => GetEffectiveCenterCode(n))
                    .Where(g => g.Select(x => x.NodalCode).Distinct().Count() > 1)
                    .ToList();

                if (multiCenterColleges.Any() || multiNodalCenters.Any())
                {
                    var rule1Errors = new List<string>();
                    if (multiCenterColleges.Any())
                    {
                        var msg = string.Join("; ", multiCenterColleges.Select(g => $"College {g.Key.CollegeKey} ({g.Key.Gender}) assigned to {g.Select(x => GetEffectiveCenterCode(x)).Distinct().Count()} centers"));
                        rule1Errors.Add($"Rule 1 Failed: Multiple exam centers for same college/gender - {msg}");
                    }
                    if (multiNodalCenters.Any())
                    {
                        var msg = string.Join("; ", multiNodalCenters.Select(g => $"Center {g.Key} assigned to {g.Select(x => x.NodalCode).Distinct().Count()} nodal codes"));
                        rule1Errors.Add($"Rule 1 Failed: Multiple nodal codes for same exam center - {msg}");
                    }

                    var multiCenterDetails = multiCenterColleges.Select(g => new
                    {
                        collegeCode = g.Key.CollegeKey,
                        gender = g.Key.Gender,
                        collegeNames = g.Select(x => x.CollegeName).Where(x => !string.IsNullOrEmpty(x)).Distinct().ToList(),
                        centerCount = g.Select(x => GetEffectiveCenterCode(x)).Distinct().Count(),
                        centerCodes = g.Select(x => GetEffectiveCenterCode(x)).Distinct().ToList(),
                        centerNames = g.Select(x => x.ExamCenterName).Where(x => !string.IsNullOrEmpty(x)).Distinct().ToList(),
                        nodalCodes = g.Select(x => x.NodalCode).Distinct().ToList(),
                        nodalNames = g.Select(x => x.NodalName).Where(x => !string.IsNullOrEmpty(x)).Distinct().ToList(),
                        records = g.Select(x => new { x.Id, x.CollegeCode, x.CollegeName, x.ExamCenterCode, x.ExamCenterName, x.NodalCode, x.NodalName, x.Gender }).ToList(),
                        description = $"College {g.Key.CollegeKey} ({g.Key.Gender}) assigned to {g.Select(x => GetEffectiveCenterCode(x)).Distinct().Count()} centers: {string.Join(", ", g.Select(x => GetEffectiveCenterCode(x)).Distinct())}"
                    }).ToList();

                    var multiNodalDetails = multiNodalCenters.Select(g => new
                    {
                        centerCode = g.Key,
                        centerNames = g.Select(x => x.ExamCenterName).Where(x => !string.IsNullOrEmpty(x)).Distinct().ToList(),
                        nodalCount = g.Select(x => x.NodalCode).Distinct().Count(),
                        nodalCodes = g.Select(x => x.NodalCode).Distinct().ToList(),
                        nodalNames = g.Select(x => x.NodalName).Where(x => !string.IsNullOrEmpty(x)).Distinct().ToList(),
                        collegeCodes = g.Select(x => GetEffectiveCollegeCode(x)).Distinct().ToList(),
                        records = g.Select(x => new { x.Id, x.CollegeCode, x.CollegeName, x.ExamCenterCode, x.ExamCenterName, x.NodalCode, x.NodalName, x.Gender }).ToList(),
                        description = $"Center {g.Key} assigned to {g.Select(x => x.NodalCode).Distinct().Count()} nodal codes: {string.Join(", ", g.Select(x => x.NodalCode).Distinct())}"
                    }).ToList();

                    return BadRequest(new
                    {
                        message = "Rule 1 validation failed: One college code must belong to one exam center, and one exam center must belong to one nodal code.",
                        errors = rule1Errors,
                        rule1CollegeCodes = multiCenterColleges.Select(g => g.Key.CollegeKey).Distinct().ToList(),
                        rule1CenterCodes = multiNodalCenters.Select(g => g.Key).Distinct().ToList(),
                        multiCenterDetails,
                        multiNodalDetails
                    });
                }

                // Clear existing temp data
                await _context.TemporaryNrDatas.Where(x => x.ProjectId == projectId).ExecuteDeleteAsync();

                var catchLists = await _context.CatchList.AsNoTracking().Where(x => x.ProjectId == projectId && x.Status).ToListAsync();

                bool allCatchCollegeCodesZero = catchLists.Any() && catchLists.All(c => c.CollegeCode == 0 && string.IsNullOrEmpty(ExtractCodeNumber(c.CollegeName)));
                var requestedMode = (mergeBy ?? "CollegeCode").Trim();
                var mode = (allCatchCollegeCodesZero && (string.Equals(requestedMode, "collegecode", StringComparison.OrdinalIgnoreCase) || string.IsNullOrEmpty(requestedMode)))
                    ? "centercode"
                    : requestedMode;

                Func<CatchList, string> getCatchKey;
                Func<NodalList, string> getNodalKey;

                switch (mode.ToLowerInvariant())
                {
                    case "collegename":
                        getCatchKey = c => (c.CollegeName ?? "").Trim().ToLowerInvariant();
                        getNodalKey = n => (n.CollegeName ?? "").Trim().ToLowerInvariant();
                        break;
                    case "centercode":
                        getCatchKey = c => (c.CenterCode != 0 ? c.CenterCode.ToString() : (ExtractCodeNumber(c.CenterName) ?? c.CollegeCode.ToString())).Trim();
                        getNodalKey = n => GetEffectiveCenterCode(n);
                        break;
                    case "centername":
                        getCatchKey = c => (c.CollegeName ?? "").Trim().ToLowerInvariant();
                        getNodalKey = n => (n.ExamCenterName ?? "").Trim().ToLowerInvariant();
                        break;
                    case "collegecode":
                    default:
                        getCatchKey = c => {
                            if (c.CollegeCode != 0) return c.CollegeCode.ToString();
                            var code = ExtractCodeNumber(c.CollegeName);
                            if (!string.IsNullOrEmpty(code) && code != "0") return code;
                            return (c.CollegeName ?? "").Trim().ToLowerInvariant();
                        };
                        getNodalKey = n => GetEffectiveCollegeCode(n);
                        break;
                }

                static string? GetValidCenterCode(NodalList n)
                {
                    if (n.ExamCenterCode != 0) return n.ExamCenterCode.ToString();
                    if (!string.IsNullOrWhiteSpace(n.ExamCenterName))
                    {
                        var code = ExtractCodeNumber(n.ExamCenterName);
                        if (code != null && !code.StartsWith("center_")) return code;
                    }
                    return null;
                }

                static string? GetValidNodalCode(NodalList n)
                {
                    if (n.NodalCode != 0) return n.NodalCode.ToString();
                    if (!string.IsNullOrWhiteSpace(n.NodalName))
                    {
                        var code = ExtractCodeNumber(n.NodalName);
                        if (code != null && !code.StartsWith("nodal_")) return code;
                    }
                    return null;
                }

                var nodalLookup = nodalLists
                    .Where(n => !string.IsNullOrEmpty(getNodalKey(n)))
                    .ToLookup(n => getNodalKey(n));

                var tempDatas = new List<TemporaryNrDatas>(catchLists.Count);

                foreach (var catchItem in catchLists)
                {
                    var key = getCatchKey(catchItem);
                    var matchingNodals = !string.IsNullOrEmpty(key) ? nodalLookup[key].ToList() : new List<NodalList>();

                    if (!matchingNodals.Any() && allCatchCollegeCodesZero)
                    {
                        var catchCenterKey = catchItem.CenterCode != 0 ? catchItem.CenterCode.ToString() : ExtractCodeNumber(catchItem.CenterName);
                        if (!string.IsNullOrEmpty(catchCenterKey))
                        {
                            var centerNodals = nodalLists.Where(n => 
                                GetEffectiveCenterCode(n) == catchCenterKey || 
                                n.ExamCenterCode.ToString() == catchCenterKey || 
                                GetValidCenterCode(n) == catchCenterKey ||
                                GetEffectiveCollegeCode(n) == catchCenterKey
                            ).ToList();
                            if (centerNodals.Any())
                            {
                                matchingNodals = centerNodals;
                            }
                        }
                    }

                    if (!matchingNodals.Any())
                    {
                        var centerCodeStr = catchItem.CenterCode != 0 ? catchItem.CenterCode.ToString() : (ExtractCodeNumber(catchItem.CenterName) ?? "");
                        tempDatas.Add(new TemporaryNrDatas
                        {
                            ProjectId = projectId,
                            CatchNo = catchItem.CatchNo,
                            NRQuantity = catchItem.NRQuantity,
                            CollegeCode = catchItem.CollegeCode,
                            CollegeName = catchItem.CollegeName,
                            CenterCode = centerCodeStr,
                            NodalCode = centerCodeStr,
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

                        if (quantity > 0 || gender == "ALL")
                        {
                            var centerCodeStr = GetValidCenterCode(nodalItem) ?? (nodalItem.ExamCenterCode != 0 ? nodalItem.ExamCenterCode.ToString() : (catchItem.CenterCode != 0 ? catchItem.CenterCode.ToString() : ""));
                            var nodalCodeStr = GetValidNodalCode(nodalItem) ?? (nodalItem.NodalCode != 0 ? nodalItem.NodalCode.ToString() : centerCodeStr);

                            tempDatas.Add(new TemporaryNrDatas
                            {
                                ProjectId = projectId,
                                CatchNo = catchItem.CatchNo,
                                NRQuantity = quantity,
                                CollegeCode = catchItem.CollegeCode,
                                CollegeName = catchItem.CollegeName,
                                CenterCode = centerCodeStr,
                                NodalCode = nodalCodeStr,
                                CourseName = catchItem.CourseName,
                                SubjectName = catchItem.SubjectName,
                                ExamDate = catchItem.ExamDate,
                                ExamTime = catchItem.ExamTime,
                                NRDatas = catchItem.NRDatas
                            });
                        }
                    }
                }

                await BulkInsertTemporaryNrDatasAsync(tempDatas);

                // Rule 2 Validation: Check if every college is assigned an exam center and nodal
                var unassignedCatchItems = catchLists
                    .Where(c => {
                        var key = getCatchKey(c);
                        var nodals = !string.IsNullOrEmpty(key) ? nodalLookup[key].ToList() : new List<NodalList>();

                        if (!nodals.Any() && allCatchCollegeCodesZero)
                        {
                            var catchCenterKey = c.CenterCode != 0 ? c.CenterCode.ToString() : ExtractCodeNumber(c.CenterName);
                            if (!string.IsNullOrEmpty(catchCenterKey))
                            {
                                var centerNodals = nodalLists.Where(n => 
                                    GetEffectiveCenterCode(n) == catchCenterKey || 
                                    n.ExamCenterCode.ToString() == catchCenterKey || 
                                    GetValidCenterCode(n) == catchCenterKey ||
                                    GetEffectiveCollegeCode(n) == catchCenterKey
                                ).ToList();
                                if (centerNodals.Any())
                                {
                                    nodals = centerNodals;
                                }
                            }
                        }

                        if (!nodals.Any())
                        {
                            return true;
                        }

                        var validNodals = nodals
                            .Where(n => GetValidCenterCode(n) != null || GetValidNodalCode(n) != null || n.ExamCenterCode != 0 || n.NodalCode != 0)
                            .ToList();

                        if (!validNodals.Any()) return true;

                        bool needsMale = c.Male > 0;
                        bool needsFemale = c.Female > 0;

                        if (needsMale)
                        {
                            bool maleOk = validNodals.Any(n => 
                                string.Equals(n.Gender, "MALE", StringComparison.OrdinalIgnoreCase) || 
                                string.Equals(n.Gender, "ALL", StringComparison.OrdinalIgnoreCase) || 
                                string.IsNullOrWhiteSpace(n.Gender));
                            if (!maleOk) return true;
                        }

                        if (needsFemale)
                        {
                            bool femaleOk = validNodals.Any(n => 
                                string.Equals(n.Gender, "FEMALE", StringComparison.OrdinalIgnoreCase) || 
                                string.Equals(n.Gender, "ALL", StringComparison.OrdinalIgnoreCase) || 
                                string.IsNullOrWhiteSpace(n.Gender));
                            if (!femaleOk) return true;
                        }

                        return false;
                    })
                    .ToList();

                bool rule2Passed = !unassignedCatchItems.Any();

                if (!rule2Passed)
                {
                    var rule2CollegeCodes = unassignedCatchItems
                        .Select(c => {
                            var k = getCatchKey(c);
                            if (!string.IsNullOrEmpty(k) && k != "0") return k;
                            return c.CenterCode != 0 ? c.CenterCode.ToString() : (ExtractCodeNumber(c.CenterName) ?? k);
                        })
                        .Where(k => !string.IsNullOrEmpty(k) && k != "0")
                        .Distinct()
                        .ToList();

                    var unassignedDetails = unassignedCatchItems
                        .GroupBy(c => getCatchKey(c))
                        .Select(g => {
                            var firstItem = g.First();
                            var keyVal = g.Key ?? "";
                            var cKey = (keyVal == "0" || string.IsNullOrEmpty(keyVal)) 
                                ? (firstItem.CenterCode != 0 ? firstItem.CenterCode.ToString() : (ExtractCodeNumber(firstItem.CenterName) ?? "")) 
                                : keyVal;
                            var cName = !string.IsNullOrEmpty(firstItem.CollegeName) ? firstItem.CollegeName : (firstItem.CenterName ?? "");
                            
                            string displayCollege;
                            if (string.IsNullOrEmpty(cName))
                            {
                                displayCollege = !string.IsNullOrEmpty(cKey) ? $"Center {cKey}" : "Unassigned Record";
                            }
                            else if (!string.IsNullOrEmpty(cKey) && cName.ToLowerInvariant().Contains(cKey.ToLowerInvariant()))
                            {
                                displayCollege = cName;
                            }
                            else if (!string.IsNullOrEmpty(cKey) && cKey.ToLowerInvariant().Contains(cName.ToLowerInvariant()))
                            {
                                displayCollege = cKey;
                            }
                            else if (!string.IsNullOrEmpty(cKey))
                            {
                                displayCollege = $"{cKey} - {cName}";
                            }
                            else
                            {
                                displayCollege = cName;
                            }
                            return new
                            {
                                collegeCode = cKey,
                                collegeName = cName,
                                catchNos = g.Select(x => x.CatchNo).Where(x => !string.IsNullOrEmpty(x)).Distinct().ToList(),
                                courses = g.Select(x => x.CourseName).Where(x => !string.IsNullOrEmpty(x)).Distinct().ToList(),
                                subjects = g.Select(x => x.SubjectName).Where(x => !string.IsNullOrEmpty(x)).Distinct().ToList(),
                                centerCodes = g.Select(x => x.CenterCode != 0 ? x.CenterCode.ToString() : ExtractCodeNumber(x.CenterName)).Where(x => !string.IsNullOrEmpty(x) && x != "0").Distinct().ToList(),
                                centerNames = g.Select(x => !string.IsNullOrEmpty(x.CenterName) ? x.CenterName : (x.CenterCode != 0 ? $"Center {x.CenterCode}" : "")).Where(x => !string.IsNullOrEmpty(x)).Distinct().ToList(),
                                totalQuantity = g.Sum(x => x.NRQuantity),
                                description = $"{displayCollege} is not assigned to any Exam Center in Nodal List."
                            };
                        })
                        .ToList();

                    var rule2Errors = new List<string> { $"Rule 2 Failed: {rule2CollegeCodes.Count} college(s) are not assigned to an exam center: {string.Join(", ", rule2CollegeCodes)}" };

                    return Ok(new
                    {
                        message = "Data merged into temporary staging with Rule 2 warnings. Every college must be assigned an exam center before pushing to main NR Data.",
                        count = tempDatas.Count,
                        rule1Passed = true,
                        rule2Passed = false,
                        rule2Errors = rule2Errors,
                        rule2CollegeCodes = rule2CollegeCodes,
                        unassignedDetails = unassignedDetails,
                        errors = rule2Errors
                    });
                }

                return Ok(new { message = "Data merged into temporary staging successfully", count = tempDatas.Count, rule1Passed = true, rule2Passed = true });
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
                var nodalLists = await _context.NodalList.AsNoTracking().Where(x => x.ProjectId == projectId && x.Status).ToListAsync();
                var catchLists = await _context.CatchList.AsNoTracking().Where(x => x.ProjectId == projectId && x.Status).ToListAsync();

                // Build pair list
                var unifiedList = new List<(CatchList? Catch, NodalList? Nodal)>();
                var catchGroup = catchLists.GroupBy(c => c.CollegeCode).ToDictionary(g => g.Key, g => g.ToList());
                var nodalGroup = nodalLists.GroupBy(n => n.CollegeCode).ToDictionary(g => g.Key, g => g.ToList());
                var allCollegeCodes = catchGroup.Keys.Union(nodalGroup.Keys).ToList();

                foreach (var code in allCollegeCodes)
                {
                    var cList = catchGroup.ContainsKey(code) ? catchGroup[code] : new List<CatchList>();
                    var nList = nodalGroup.ContainsKey(code) ? nodalGroup[code] : new List<NodalList>();

                    if (nList.Any())
                    {
                        var firstCatch = cList.FirstOrDefault();
                        foreach (var n in nList)
                        {
                            unifiedList.Add((firstCatch, n));
                        }
                    }
                    else if (cList.Any())
                    {
                        foreach (var c in cList)
                        {
                            unifiedList.Add((c, null));
                        }
                    }
                }

                static string GetVal(CatchList? c, NodalList? n, string field)
                {
                    var f = field.Trim().ToLowerInvariant().Replace(" ", "").Replace("_", "");
                    switch (f)
                    {
                        case "collegecode":
                            if (n != null && n.CollegeCode != 0) return n.CollegeCode.ToString();
                            if (c != null && c.CollegeCode != 0) return c.CollegeCode.ToString();
                            if (n != null && !string.IsNullOrEmpty(n.CollegeName))
                            {
                                var match = System.Text.RegularExpressions.Regex.Match(n.CollegeName, @"^(\d+)");
                                if (match.Success) return match.Groups[1].Value;
                            }
                            return "";
                        case "collegename":
                            if (n != null && !string.IsNullOrEmpty(n.CollegeName)) return n.CollegeName;
                            if (c != null && !string.IsNullOrEmpty(c.CollegeName)) return c.CollegeName;
                            return "";
                        case "examcentercode":
                        case "centercode":
                            if (n != null && n.ExamCenterCode != 0) return n.ExamCenterCode.ToString();
                            if (c != null && c.CenterCode != 0) return c.CenterCode.ToString();
                            if (n != null && !string.IsNullOrEmpty(n.ExamCenterName))
                            {
                                var match = System.Text.RegularExpressions.Regex.Match(n.ExamCenterName, @"^(\d+)");
                                if (match.Success) return match.Groups[1].Value;
                            }
                            return "";
                        case "examcentername":
                        case "centername":
                            if (n != null && !string.IsNullOrEmpty(n.ExamCenterName)) return n.ExamCenterName;
                            if (c != null && !string.IsNullOrEmpty(c.CenterName)) return c.CenterName;
                            return "";
                        case "gender":
                            if (n != null && !string.IsNullOrEmpty(n.Gender)) return n.Gender.Trim().ToUpper();
                            return "";
                        case "nodalcode":
                            if (n != null && n.NodalCode != 0) return n.NodalCode.ToString();
                            if (n != null && !string.IsNullOrEmpty(n.NodalName))
                            {
                                var match = System.Text.RegularExpressions.Regex.Match(n.NodalName, @"^(\d+)");
                                if (match.Success) return match.Groups[1].Value;
                            }
                            return "";
                        case "nodalname":
                            if (n != null && !string.IsNullOrEmpty(n.NodalName)) return n.NodalName;
                            return "";
                        case "catchno":
                            if (c != null && !string.IsNullOrEmpty(c.CatchNo)) return c.CatchNo;
                            return "";
                        case "nrquantity":
                        case "quantity":
                            if (c != null) return c.NRQuantity.ToString();
                            return "";
                        case "coursename":
                        case "course":
                            if (c != null && !string.IsNullOrEmpty(c.CourseName)) return c.CourseName;
                            return "";
                        case "subjectname":
                        case "subject":
                            if (c != null && !string.IsNullOrEmpty(c.SubjectName)) return c.SubjectName;
                            return "";
                        case "examdate":
                            if (c != null && !string.IsNullOrEmpty(c.ExamDate)) return c.ExamDate;
                            return "";
                        case "examtime":
                            if (c != null && !string.IsNullOrEmpty(c.ExamTime)) return c.ExamTime;
                            return "";
                        default:
                            if (n != null)
                            {
                                var p = typeof(NodalList).GetProperty(field, System.Reflection.BindingFlags.IgnoreCase | System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance);
                                if (p != null) return p.GetValue(n)?.ToString() ?? "";
                            }
                            if (c != null)
                            {
                                var p = typeof(CatchList).GetProperty(field, System.Reflection.BindingFlags.IgnoreCase | System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance);
                                if (p != null) return p.GetValue(c)?.ToString() ?? "";
                            }
                            return "";
                    }
                }

                static string GetKey(CatchList? c, NodalList? n, List<string> fields)
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
                                Status = 1,
                                Rule = 1
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
                                Status = 1,
                                Rule = 1
                            });
                        }
                    }
                }

                await _context.ConflictingFields.Where(c => c.ProjectId == projectId && c.Rule == 1).ExecuteDeleteAsync();
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

        [HttpDelete("ClearDynamicRule1/{projectId}")]
        public async Task<IActionResult> ClearDynamicRule1(int projectId)
        {
            try
            {
                await _context.ConflictingFields.Where(c => c.ProjectId == projectId && c.Rule == 1).ExecuteDeleteAsync();
                return Ok(new { message = "Dynamic conflicts cleared successfully" });
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
                var dynamicConflicts = await _context.ConflictingFields
                    .AsNoTracking()
                    .Where(c => c.ProjectId == projectId && c.Status == 1 && c.Rule == 1)
                    .ToListAsync();

                var catchLists = await _context.CatchList
                    .AsNoTracking()
                    .Where(x => x.ProjectId == projectId && x.Status)
                    .ToListAsync();

                var nodalLists = await _context.NodalList
                    .AsNoTracking()
                    .Where(x => x.ProjectId == projectId && x.Status)
                    .ToListAsync();

                static string GetEffectiveCollegeCode(NodalList n)
                {
                    if (n.CollegeCode != 0) return n.CollegeCode.ToString();
                    if (!string.IsNullOrEmpty(n.CollegeName))
                    {
                        var code = ExtractCodeNumber(n.CollegeName);
                        if (code != null && code != "0") return code;
                    }
                    return GetEffectiveCenterCode(n);
                }

                static string GetEffectiveCenterCode(NodalList n)
                {
                    if (n.ExamCenterCode != 0) return n.ExamCenterCode.ToString();
                    if (!string.IsNullOrEmpty(n.ExamCenterName))
                    {
                        var code = ExtractCodeNumber(n.ExamCenterName);
                        if (code != null) return code;
                        return n.ExamCenterName.Trim().ToLowerInvariant();
                    }
                    return $"center_{n.Id}";
                }

                // Rule 1 Validation: One college -> one center, One center -> one nodal
                var groupedByCollegeGender = nodalLists
                    .GroupBy(n => new { CollegeKey = GetEffectiveCollegeCode(n), Gender = n.Gender?.ToUpper() ?? "ALL" })
                    .ToList();

                var multiCenterCollegesGroup = groupedByCollegeGender
                    .Where(g => g.Select(x => GetEffectiveCenterCode(x)).Distinct().Count() > 1)
                    .ToList();

                var multiNodalCentersGroup = nodalLists
                    .GroupBy(n => GetEffectiveCenterCode(n))
                    .Where(g => g.Select(x => x.NodalCode).Distinct().Count() > 1)
                    .ToList();

                var rule1Errors = new List<string>();
                if (multiCenterCollegesGroup.Any())
                {
                    var msg = string.Join("; ", multiCenterCollegesGroup.Select(g => $"College {g.Key.CollegeKey} ({g.Key.Gender}) assigned to {g.Select(x => GetEffectiveCenterCode(x)).Distinct().Count()} centers"));
                    rule1Errors.Add($"Rule 1 Failed: Multiple exam centers for same college/gender - {msg}");
                }
                if (multiNodalCentersGroup.Any())
                {
                    var msg = string.Join("; ", multiNodalCentersGroup.Select(g => $"Center {g.Key} assigned to {g.Select(x => x.NodalCode).Distinct().Count()} nodal codes"));
                    rule1Errors.Add($"Rule 1 Failed: Multiple nodal codes for same exam center - {msg}");
                }

                var multiCenterDetails = multiCenterCollegesGroup.Select(g => new
                {
                    collegeCode = g.Key.CollegeKey,
                    gender = g.Key.Gender,
                    collegeNames = g.Select(x => x.CollegeName).Where(x => !string.IsNullOrEmpty(x)).Distinct().ToList(),
                    centerCount = g.Select(x => GetEffectiveCenterCode(x)).Distinct().Count(),
                    centerCodes = g.Select(x => GetEffectiveCenterCode(x)).Distinct().ToList(),
                    centerNames = g.Select(x => x.ExamCenterName).Where(x => !string.IsNullOrEmpty(x)).Distinct().ToList(),
                    nodalCodes = g.Select(x => x.NodalCode).Distinct().ToList(),
                    nodalNames = g.Select(x => x.NodalName).Where(x => !string.IsNullOrEmpty(x)).Distinct().ToList(),
                    records = g.Select(x => new { x.Id, x.CollegeCode, x.CollegeName, x.ExamCenterCode, x.ExamCenterName, x.NodalCode, x.NodalName, x.Gender }).ToList(),
                    description = $"College {g.Key.CollegeKey} ({g.Key.Gender}) assigned to {g.Select(x => GetEffectiveCenterCode(x)).Distinct().Count()} centers: {string.Join(", ", g.Select(x => GetEffectiveCenterCode(x)).Distinct())}"
                }).Cast<object>().ToList();

                var multiNodalDetails = multiNodalCentersGroup.Select(g => new
                {
                    centerCode = g.Key,
                    centerNames = g.Select(x => x.ExamCenterName).Where(x => !string.IsNullOrEmpty(x)).Distinct().ToList(),
                    nodalCount = g.Select(x => x.NodalCode).Distinct().Count(),
                    nodalCodes = g.Select(x => x.NodalCode).Distinct().ToList(),
                    nodalNames = g.Select(x => x.NodalName).Where(x => !string.IsNullOrEmpty(x)).Distinct().ToList(),
                    collegeCodes = g.Select(x => GetEffectiveCollegeCode(x)).Distinct().ToList(),
                    records = g.Select(x => new { x.Id, x.CollegeCode, x.CollegeName, x.ExamCenterCode, x.ExamCenterName, x.NodalCode, x.NodalName, x.Gender }).ToList(),
                    description = $"Center {g.Key} assigned to {g.Select(x => x.NodalCode).Distinct().Count()} nodal codes: {string.Join(", ", g.Select(x => x.NodalCode).Distinct())}"
                }).Cast<object>().ToList();

                var multiCenterColleges = multiCenterCollegesGroup.Select(g => g.Key.CollegeKey).Distinct().ToList();
                var multiNodalCenters = multiNodalCentersGroup.Select(g => g.Key).Distinct().ToList();

                // Rule 2 Validation
                bool allCatchCollegeCodesZero = catchLists.Any() && catchLists.All(c => c.CollegeCode == 0 && string.IsNullOrEmpty(ExtractCodeNumber(c.CollegeName)));
                var requestedMode = (mergeBy ?? "CollegeCode").Trim();
                var mode = (allCatchCollegeCodesZero && (string.Equals(requestedMode, "collegecode", StringComparison.OrdinalIgnoreCase) || string.IsNullOrEmpty(requestedMode)))
                    ? "centercode"
                    : requestedMode;

                Func<CatchList, string> getCatchKey;
                Func<NodalList, string> getNodalKey;

                switch (mode.ToLowerInvariant())
                {
                    case "collegename":
                        getCatchKey = c => (c.CollegeName ?? "").Trim().ToLowerInvariant();
                        getNodalKey = n => (n.CollegeName ?? "").Trim().ToLowerInvariant();
                        break;
                    case "centercode":
                        getCatchKey = c => (c.CenterCode != 0 ? c.CenterCode.ToString() : (ExtractCodeNumber(c.CenterName) ?? c.CollegeCode.ToString())).Trim();
                        getNodalKey = n => GetEffectiveCenterCode(n);
                        break;
                    case "centername":
                        getCatchKey = c => (c.CollegeName ?? "").Trim().ToLowerInvariant();
                        getNodalKey = n => (n.ExamCenterName ?? "").Trim().ToLowerInvariant();
                        break;
                    case "collegecode":
                    default:
                        getCatchKey = c => {
                            if (c.CollegeCode != 0) return c.CollegeCode.ToString();
                            var code = ExtractCodeNumber(c.CollegeName);
                            if (!string.IsNullOrEmpty(code) && code != "0") return code;
                            return (c.CollegeName ?? "").Trim().ToLowerInvariant();
                        };
                        getNodalKey = n => GetEffectiveCollegeCode(n);
                        break;
                }

                static string? GetValidCenterCode(NodalList n)
                {
                    if (n.ExamCenterCode != 0) return n.ExamCenterCode.ToString();
                    if (!string.IsNullOrWhiteSpace(n.ExamCenterName))
                    {
                        var code = ExtractCodeNumber(n.ExamCenterName);
                        if (code != null && !code.StartsWith("center_")) return code;
                    }
                    return null;
                }

                static string? GetValidNodalCode(NodalList n)
                {
                    if (n.NodalCode != 0) return n.NodalCode.ToString();
                    if (!string.IsNullOrWhiteSpace(n.NodalName))
                    {
                        var code = ExtractCodeNumber(n.NodalName);
                        if (code != null && !code.StartsWith("nodal_")) return code;
                    }
                    return null;
                }

                // Rule 2 Validation: Evaluated ONLY when data is present in TemporaryNrDatas table
                bool hasTempData = await _context.TemporaryNrDatas.AnyAsync(x => x.ProjectId == projectId);

                var rule2CollegeCodes = new List<string>();
                var unassignedDetails = new List<object>();
                var rule2Errors = new List<string>();
                bool rule2Passed = true;

                if (hasTempData)
                {
                    var tempDatas = await _context.TemporaryNrDatas.AsNoTracking().Where(x => x.ProjectId == projectId).ToListAsync();

                    List<NodalList> FindMatchingNodals(TemporaryNrDatas t)
                    {
                        // 1. Try matching by College Code / College Name first
                        string collegeKey = "";
                        if (t.CollegeCode != 0)
                        {
                            collegeKey = t.CollegeCode.ToString();
                        }
                        else
                        {
                            var code = ExtractCodeNumber(t.CollegeName);
                            if (!string.IsNullOrEmpty(code) && code != "0") collegeKey = code;
                        }

                        if (!string.IsNullOrEmpty(collegeKey))
                        {
                            var matchedByCollege = nodalLists.Where(n => 
                                GetEffectiveCollegeCode(n) == collegeKey || 
                                n.CollegeCode.ToString() == collegeKey ||
                                (n.CollegeName ?? "").Trim().ToLowerInvariant() == collegeKey.ToLowerInvariant()
                            ).ToList();

                            if (matchedByCollege.Any()) return matchedByCollege;
                        }

                        // 2. If no college match (or no college code), try matching by Center Code / Center Name
                        string catchCenterKey = !string.IsNullOrEmpty(t.CenterCode) && t.CenterCode != "0" 
                            ? t.CenterCode 
                            : ExtractCodeNumber(t.CollegeName);

                        if (!string.IsNullOrEmpty(catchCenterKey) && catchCenterKey != "0")
                        {
                            var matchedByCenter = nodalLists.Where(n => 
                                GetEffectiveCenterCode(n) == catchCenterKey || 
                                n.ExamCenterCode.ToString() == catchCenterKey || 
                                GetValidCenterCode(n) == catchCenterKey ||
                                GetEffectiveCollegeCode(n) == catchCenterKey ||
                                n.CollegeCode.ToString() == catchCenterKey
                            ).ToList();

                            if (matchedByCenter.Any()) return matchedByCenter;
                        }

                        return new List<NodalList>();
                    }

                    Func<TemporaryNrDatas, string> getTempKey = t => {
                        if (t.CollegeCode != 0) return t.CollegeCode.ToString();
                        var code = ExtractCodeNumber(t.CollegeName);
                        if (!string.IsNullOrEmpty(code) && code != "0") return code;
                        if (!string.IsNullOrEmpty(t.CenterCode) && t.CenterCode != "0") return t.CenterCode;
                        return "Unknown";
                    };

                    var unassignedTempItems = tempDatas
                        .Where(t => {
                            var nodals = FindMatchingNodals(t);

                            if (!nodals.Any())
                            {
                                return true;
                            }

                            var validNodals = nodals
                                .Where(n => GetValidCenterCode(n) != null || GetValidNodalCode(n) != null || n.ExamCenterCode != 0 || n.NodalCode != 0)
                                .ToList();

                            if (!validNodals.Any()) return true;

                            return false;
                        })
                        .ToList();

                    rule2CollegeCodes = unassignedTempItems
                        .Select(t => {
                            var k = getTempKey(t);
                            if (!string.IsNullOrEmpty(k) && k != "Unknown") return k;
                            return !string.IsNullOrEmpty(t.CenterCode) && t.CenterCode != "0" ? t.CenterCode : k;
                        })
                        .Where(k => !string.IsNullOrEmpty(k) && k != "Unknown")
                        .Distinct()
                        .ToList();

                    unassignedDetails = unassignedTempItems
                        .GroupBy(t => getTempKey(t))
                        .Select(g => {
                            var firstItem = g.First();
                            var keyVal = g.Key ?? "";
                            var rawKey = (keyVal == "0" || string.IsNullOrEmpty(keyVal)) 
                                ? (!string.IsNullOrEmpty(firstItem.CenterCode) && firstItem.CenterCode != "0" ? firstItem.CenterCode : (ExtractCodeNumber(firstItem.CollegeName) ?? "")) 
                                : keyVal;
                            var cKey = string.IsNullOrEmpty(rawKey) || rawKey == "0" ? "Unknown" : rawKey;
                            var cName = !string.IsNullOrEmpty(firstItem.CollegeName) ? firstItem.CollegeName : (!string.IsNullOrEmpty(firstItem.CenterCode) && firstItem.CenterCode != "0" ? $"Center {firstItem.CenterCode}" : "Unassigned Catch Record");
                            
                            string displayCollege;
                            if (cKey == "Unknown" || string.IsNullOrEmpty(cKey))
                            {
                                displayCollege = "Unassigned Catch Record (Missing College/Center Code)";
                            }
                            else if (string.IsNullOrEmpty(cName))
                            {
                                displayCollege = $"Center {cKey}";
                            }
                            else if (cName.ToLowerInvariant().Contains(cKey.ToLowerInvariant()))
                            {
                                displayCollege = cName;
                            }
                            else if (cKey.ToLowerInvariant().Contains(cName.ToLowerInvariant()))
                            {
                                displayCollege = cKey;
                            }
                            else
                            {
                                displayCollege = $"{cKey} - {cName}";
                            }
                                bool isCenterItem = displayCollege.StartsWith("Center ", StringComparison.OrdinalIgnoreCase) || firstItem.CollegeCode == 0;
                                string targetCenterType = isCenterItem ? "Nodal Center" : "Exam Center";
                                return new
                                {
                                    collegeCode = cKey,
                                    collegeName = cName,
                                    courses = g.Select(x => x.CourseName).Where(x => !string.IsNullOrEmpty(x)).Distinct().ToList(),
                                    subjects = g.Select(x => x.SubjectName).Where(x => !string.IsNullOrEmpty(x)).Distinct().ToList(),
                                    centerCodes = g.Select(x => x.CenterCode).Where(x => !string.IsNullOrEmpty(x) && x != "0").Distinct().ToList(),
                                    centerNames = g.Select(x => {
                                        var cCode = x.CenterCode;
                                        var nMatch = nodalLists.FirstOrDefault(n => GetEffectiveCenterCode(n) == cCode || n.ExamCenterCode.ToString() == cCode);
                                        return nMatch != null && !string.IsNullOrEmpty(nMatch.ExamCenterName) ? nMatch.ExamCenterName : $"Center {cCode}";
                                    }).Distinct().ToList(),
                                    totalQuantity = g.Sum(x => x.NRQuantity),
                                    description = $"{displayCollege} is not assigned to any {targetCenterType} in Nodal List."
                                };
                        })
                        .Cast<object>()
                        .ToList();

                    rule2Errors = unassignedTempItems.Any()
                        ? new List<string> { $"Rule 2 Failed: {rule2CollegeCodes.Count} college(s) are not assigned to an exam center: {string.Join(", ", rule2CollegeCodes)}" }
                        : new List<string>();

                    rule2Passed = !unassignedTempItems.Any();
                }

                // Sync all calculated Rule 1 & Rule 2 conflicts to ConflictingFields table in DB (soft resolving fixed ones by setting Status = 0)
                await SyncConflictsToDbAsync(projectId, multiCenterDetails, multiNodalDetails, unassignedDetails);


                bool rule1Passed = !rule1Errors.Any() && !dynamicConflicts.Any();

                return Ok(new
                {
                    rule1Passed,
                    rule1Errors,
                    multiCenterDetails,
                    multiNodalDetails,
                    rule1CollegeCodes = multiCenterColleges,
                    rule1CenterCodes = multiNodalCenters,

                    rule2Passed,
                    rule2Errors,
                    rule2CollegeCodes,
                    unassignedCount = unassignedDetails.Count,
                    unassignedDetails,

                    dynamicConflicts
                });
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
        }

        private async Task SyncConflictsToDbAsync(
            int projectId,
            List<object> multiCenterDetails,
            List<object> multiNodalDetails,
            List<object> unassignedDetails)
        {
            try
            {
                var existingInDb = await _context.ConflictingFields
                    .Where(c => c.ProjectId == projectId)
                    .ToListAsync();

                var activeUniqueKeys = new HashSet<string>();

                // 1. Rule 1: Multi-Center Colleges
                foreach (var item in multiCenterDetails)
                {
                    var json = JsonSerializer.Serialize(item);
                    using var doc = JsonDocument.Parse(json);
                    var root = doc.RootElement;
                    string cCode = root.TryGetProperty("collegeCode", out var cc) ? cc.GetString() ?? "" : "";
                    string gender = root.TryGetProperty("gender", out var g) ? g.GetString() ?? "ALL" : "ALL";

                    string uniqueKey = $"Rule1_MultiCenter_{cCode}_{gender}";
                    activeUniqueKeys.Add(uniqueKey);

                    var match = existingInDb.FirstOrDefault(e => e.UniqueField == uniqueKey);
                    if (match != null)
                    {
                        match.ConflictingField = json;
                        match.Status = 1;
                    }
                    else
                    {
                        _context.ConflictingFields.Add(new ConflictingFields
                        {
                            ProjectId = projectId,
                            UniqueField = uniqueKey,
                            ConflictingField = json,
                            Status = 1
                        });
                    }
                }

                // 2. Rule 1: Multi-Nodal Centers
                foreach (var item in multiNodalDetails)
                {
                    var json = JsonSerializer.Serialize(item);
                    using var doc = JsonDocument.Parse(json);
                    var root = doc.RootElement;
                    string cCode = root.TryGetProperty("centerCode", out var cc) ? cc.GetString() ?? "" : "";

                    string uniqueKey = $"Rule1_MultiNodal_{cCode}";
                    activeUniqueKeys.Add(uniqueKey);

                    var match = existingInDb.FirstOrDefault(e => e.UniqueField == uniqueKey);
                    if (match != null)
                    {
                        match.ConflictingField = json;
                        match.Status = 1;
                    }
                    else
                    {
                        _context.ConflictingFields.Add(new ConflictingFields
                        {
                            ProjectId = projectId,
                            UniqueField = uniqueKey,
                            ConflictingField = json,
                            Status = 1
                        });
                    }
                }

                // 3. Rule 2: Unassigned Colleges / Centers
                foreach (var item in unassignedDetails)
                {
                    var json = JsonSerializer.Serialize(item);
                    using var doc = JsonDocument.Parse(json);
                    var root = doc.RootElement;
                    string cCode = root.TryGetProperty("collegeCode", out var cc) ? cc.GetString() ?? "" : "";
                    if (string.IsNullOrEmpty(cCode) || cCode == "0") cCode = "Unknown";

                    string uniqueKey = $"Rule2_Unassigned_{cCode}";
                    activeUniqueKeys.Add(uniqueKey);

                    var match = existingInDb.FirstOrDefault(e => e.UniqueField == uniqueKey);
                    if (match != null)
                    {
                        match.ConflictingField = json;
                        match.Status = 1;
                    }
                    else
                    {
                        _context.ConflictingFields.Add(new ConflictingFields
                        {
                            ProjectId = projectId,
                            UniqueField = uniqueKey,
                            ConflictingField = json,
                            Status = 1
                        });
                    }
                }

                // 4. Soft-resolve (Status = 0) any existing Rule 1 / Rule 2 conflicts in DB that are no longer failing. NEVER DELETE!
                foreach (var ec in existingInDb)
                {
                    if (ec.UniqueField != null && (ec.UniqueField.StartsWith("Rule1_") || ec.UniqueField.StartsWith("Rule2_")))
                    {
                        if (!activeUniqueKeys.Contains(ec.UniqueField) && ec.Status == 1)
                        {
                            ec.Status = 0;
                        }
                    }
                }

                await _context.SaveChangesAsync();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"SyncConflictsToDbAsync error: {ex.Message}");
            }
        }

        [HttpGet("CheckAllConflictsResolved/{projectId}")]
        public async Task<IActionResult> CheckAllConflictsResolved(int projectId)
        {
            try
            {
                var pendingCount = await _context.ConflictingFields
                    .CountAsync(c => c.ProjectId == projectId && c.Status == 1);

                return Ok(new { allResolved = pendingCount == 0 });
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
        }

        [HttpPost("ResolveConflict/{conflictId}")]
        public async Task<IActionResult> ResolveConflict(int conflictId)
        {
            try
            {
                var conflict = await _context.ConflictingFields.FirstOrDefaultAsync(c => c.Id == conflictId);
                if (conflict != null)
                {
                    conflict.Status = 0;
                    await _context.SaveChangesAsync();
                }
                return Ok(new { message = "Conflict resolved successfully" });
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
        }

        [HttpPost("ResolveDynamicConflict")]
        public async Task<IActionResult> ResolveDynamicConflict([FromBody] ResolveDynamicConflictDto dto)
        {
            try
            {
                var conflict = await _context.ConflictingFields.FirstOrDefaultAsync(c => c.Id == dto.ConflictId);
                if (conflict != null)
                {
                    conflict.Status = 0;
                }

                var nodalLists = await _context.NodalList.Where(n => n.ProjectId == dto.ProjectId && n.Status).ToListAsync();

                var targetF = (dto.TargetField ?? "").Trim().ToLowerInvariant().Replace(" ", "").Replace("_", "");
                var matchF = (dto.MatchField ?? "").Trim().ToLowerInvariant().Replace(" ", "").Replace("_", "");
                var matchVal = (dto.MatchValue ?? "").Trim();
                int parsedTargetInt = int.TryParse(dto.TargetValue, out int ti) ? ti : 0;
                int parsedMatchInt = int.TryParse(matchVal, out int mi) ? mi : 0;

                if (parsedMatchInt == 0 && !string.IsNullOrEmpty(matchVal))
                {
                    var matchReg = System.Text.RegularExpressions.Regex.Match(matchVal, @"^(\d+)");
                    if (matchReg.Success) parsedMatchInt = int.Parse(matchReg.Groups[1].Value);
                }

                bool updatedAny = false;
                
                // Calculate names outside the loop so we look at the original un-modified list
                string resolvedNodalName = dto.TargetName;
                if (string.IsNullOrWhiteSpace(resolvedNodalName) && targetF.Contains("nodal") && parsedTargetInt > 0)
                {
                    var matching = nodalLists.FirstOrDefault(x => x.NodalCode == parsedTargetInt && !string.IsNullOrWhiteSpace(x.NodalName));
                    if (matching != null) resolvedNodalName = matching.NodalName;
                }

                string resolvedCenterName = dto.TargetName;
                if (string.IsNullOrWhiteSpace(resolvedCenterName) && (targetF.Contains("center") || targetF.Contains("examcenter")) && parsedTargetInt > 0)
                {
                    var matching = nodalLists.FirstOrDefault(x => x.ExamCenterCode == parsedTargetInt && !string.IsNullOrWhiteSpace(x.ExamCenterName));
                    if (matching != null) resolvedCenterName = matching.ExamCenterName;
                }

                string resolvedCollegeName = dto.TargetName;
                if (string.IsNullOrWhiteSpace(resolvedCollegeName) && targetF.Contains("college") && parsedTargetInt > 0)
                {
                    var matching = nodalLists.FirstOrDefault(x => x.CollegeCode == parsedTargetInt && !string.IsNullOrWhiteSpace(x.CollegeName));
                    if (matching != null) resolvedCollegeName = matching.CollegeName;
                }

                foreach (var n in nodalLists)
                {
                    bool isMatch = false;
                    if (matchF.Contains("college"))
                    {
                        isMatch = (parsedMatchInt > 0 && n.CollegeCode == parsedMatchInt) ||
                                  (!string.IsNullOrEmpty(n.CollegeName) && (n.CollegeName.Contains(matchVal) || (parsedMatchInt > 0 && n.CollegeName.StartsWith(parsedMatchInt.ToString()))));
                    }
                    else if (matchF.Contains("center") || matchF.Contains("examcenter"))
                    {
                        isMatch = (parsedMatchInt > 0 && n.ExamCenterCode == parsedMatchInt) ||
                                  (!string.IsNullOrEmpty(n.ExamCenterName) && (n.ExamCenterName.Contains(matchVal) || (parsedMatchInt > 0 && n.ExamCenterName.StartsWith(parsedMatchInt.ToString()))));
                    }
                    else if (matchF.Contains("nodal"))
                    {
                        isMatch = (parsedMatchInt > 0 && n.NodalCode == parsedMatchInt) ||
                                  (!string.IsNullOrEmpty(n.NodalName) && (n.NodalName.Contains(matchVal) || (parsedMatchInt > 0 && n.NodalName.StartsWith(parsedMatchInt.ToString()))));
                    }
                    else
                    {
                        isMatch = (parsedMatchInt > 0 && (n.CollegeCode == parsedMatchInt || n.ExamCenterCode == parsedMatchInt || n.NodalCode == parsedMatchInt));
                    }

                    if (isMatch)
                    {
                        if (targetF.Contains("nodal"))
                        {
                            n.NodalCode = parsedTargetInt > 0 ? parsedTargetInt : n.NodalCode;
                            if (!string.IsNullOrWhiteSpace(resolvedNodalName))
                            {
                                n.NodalName = resolvedNodalName;
                            }
                            updatedAny = true;
                        }
                        else if (targetF.Contains("center") || targetF.Contains("examcenter"))
                        {
                            n.ExamCenterCode = parsedTargetInt > 0 ? parsedTargetInt : n.ExamCenterCode;
                            if (!string.IsNullOrWhiteSpace(resolvedCenterName))
                            {
                                n.ExamCenterName = resolvedCenterName;
                            }
                            updatedAny = true;
                        }
                        else if (targetF.Contains("college"))
                        {
                            n.CollegeCode = parsedTargetInt > 0 ? parsedTargetInt : n.CollegeCode;
                            if (!string.IsNullOrWhiteSpace(resolvedCollegeName))
                            {
                                n.CollegeName = resolvedCollegeName;
                            }
                            updatedAny = true;
                        }
                    }
                }

                await _context.SaveChangesAsync();
                return Ok(new { message = "Dynamic conflict resolved successfully", updatedAny });
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
                var tempDatas = await _context.TemporaryNrDatas
                    .Where(x => x.ProjectId == projectId)
                    .ToListAsync();

                if (!tempDatas.Any())
                {
                    return BadRequest("No temporary data found to push. Please generate preview first.");
                }

                // Re-sync TemporaryNrDatas with NodalList assignments in case NodalList records were updated/resolved after preview generation
                var nodalLists = await _context.NodalList
                    .AsNoTracking()
                    .Where(x => x.ProjectId == projectId && x.Status)
                    .ToListAsync();

                static string GetEffectiveCollegeCode(NodalList n)
                {
                    if (n.CollegeCode != 0) return n.CollegeCode.ToString();
                    if (!string.IsNullOrEmpty(n.CollegeName))
                    {
                        var m = System.Text.RegularExpressions.Regex.Match(n.CollegeName, @"^(\d+)");
                        if (m.Success) return m.Groups[1].Value;
                        return n.CollegeName.Trim().ToLowerInvariant();
                    }
                    return "";
                }

                static string? GetValidCenterCode(NodalList n)
                {
                    if (n.ExamCenterCode != 0) return n.ExamCenterCode.ToString();
                    if (!string.IsNullOrWhiteSpace(n.ExamCenterName))
                    {
                        var trimmed = n.ExamCenterName.Trim();
                        if (trimmed != "0")
                        {
                            var m = System.Text.RegularExpressions.Regex.Match(trimmed, @"^(\d+)");
                            if (m.Success) return m.Groups[1].Value;
                            return trimmed.ToLowerInvariant();
                        }
                    }
                    return null;
                }

                static string? GetValidNodalCode(NodalList n)
                {
                    if (n.NodalCode != 0) return n.NodalCode.ToString();
                    if (!string.IsNullOrWhiteSpace(n.NodalName))
                    {
                        var trimmed = n.NodalName.Trim();
                        if (trimmed != "0")
                        {
                            var m = System.Text.RegularExpressions.Regex.Match(trimmed, @"^(\d+)");
                            if (m.Success) return m.Groups[1].Value;
                            return trimmed.ToLowerInvariant();
                        }
                    }
                    return null;
                }

                var nodalLookup = nodalLists
                    .Where(n => !string.IsNullOrEmpty(GetEffectiveCollegeCode(n)))
                    .ToLookup(n => GetEffectiveCollegeCode(n));

                bool updatedTemp = false;
                foreach (var temp in tempDatas)
                {
                    var collegeKey = temp.CollegeCode != 0 ? temp.CollegeCode.ToString() : (temp.CollegeName ?? "").Trim().ToLowerInvariant();
                    if (string.IsNullOrEmpty(collegeKey))
                    {
                        var m = System.Text.RegularExpressions.Regex.Match(temp.CollegeName ?? "", @"^(\d+)");
                        if (m.Success) collegeKey = m.Groups[1].Value;
                    }

                    if (!string.IsNullOrEmpty(collegeKey))
                    {
                        var matchNodals = nodalLookup[collegeKey].ToList();
                        if (matchNodals.Any())
                        {
                            var validNodal = matchNodals.FirstOrDefault(n => GetValidCenterCode(n) != null && GetValidNodalCode(n) != null) ?? matchNodals.First();
                            var centerCodeStr = GetValidCenterCode(validNodal) ?? (validNodal.ExamCenterCode != 0 ? validNodal.ExamCenterCode.ToString() : "");
                            var nodalCodeStr = GetValidNodalCode(validNodal) ?? (validNodal.NodalCode != 0 ? validNodal.NodalCode.ToString() : "");

                            if (!string.IsNullOrEmpty(centerCodeStr) && (temp.CenterCode != centerCodeStr || temp.NodalCode != nodalCodeStr))
                            {
                                temp.CenterCode = centerCodeStr;
                                if (!string.IsNullOrEmpty(nodalCodeStr))
                                {
                                    temp.NodalCode = nodalCodeStr;
                                }
                                updatedTemp = true;
                            }
                        }
                    }
                }

                if (updatedTemp)
                {
                    await _context.SaveChangesAsync();
                }

                // Rule 2 Validation Check: Ensure every college has an assigned exam center AND every exam center is assigned to a nodal code
                var unassignedTempRecords = tempDatas
                    .Where(x => string.IsNullOrWhiteSpace(x.CenterCode) || x.CenterCode == "0" || string.IsNullOrWhiteSpace(x.NodalCode) || x.NodalCode == "0")
                    .ToList();

                if (unassignedTempRecords.Any())
                {
                    var unassignedColleges = unassignedTempRecords
                        .Select(x => x.CollegeCode != 0 ? $"College Code {x.CollegeCode}" : (string.IsNullOrWhiteSpace(x.CollegeName) ? "Unknown College" : x.CollegeName))
                        .Distinct()
                        .ToList();

                    return BadRequest(new
                    {
                        message = "Rule 2 Validation Failed: Every college must be assigned an exam center before pushing to main NR Data.",
                        rule2Failed = true,
                        rule2Passed = false,
                        unassignedColleges = unassignedColleges,
                        count = unassignedTempRecords.Count,
                        errors = new List<string>
                        {
                            $"Rule 2 Failed: {unassignedColleges.Count} college(s) do not have an assigned exam center ({string.Join(", ", unassignedColleges)}). Pushing to main NR Data is strictly blocked."
                        }
                    });
                }

                using var transaction = await _context.Database.BeginTransactionAsync();
                try
                {
                    // Clear existing NRDatas for this project
                    await _context.NRDatas.Where(x => x.ProjectId == projectId).ExecuteDeleteAsync();

                    // Direct set-based SQL copy from TemporaryNrDatas into NRDatas
                    await _context.Database.ExecuteSqlInterpolatedAsync($@"
                        INSERT INTO `NRDatas` (
                            `ProjectId`, `CourseName`, `SubjectName`, `CenterCode`, `Quantity`, `NRQuantity`, 
                            `CatchNo`, `ExamDate`, `ExamTime`, `Day`, `NRDatas`, `NodalCode`, `Pages`, `Route`, 
                            `RouteSort`, `CenterSort`, `NodalSort`, `Symbol`, `Status`, `District`, `DistrictSort`, `Batch`,
                            `NRDataId`, `LotNo`, `Steps`, `UploadList`, `EnvLotNo`, `VerificationStatus`, `VerifiedBy`
                        )
                        SELECT 
                            `ProjectId`, `CourseName`, `SubjectName`, `CenterCode`, `NRQuantity`, `NRQuantity`, 
                            `CatchNo`, `ExamDate`, `ExamTime`, `Day`, `NRDatas`, `NodalCode`, `Pages`, `Route`, 
                            `RouteSort`, `CenterSort`, `NodalSort`, `Symbol`, `Status`, `District`, `DistrictSort`, 1,
                            0, 0, 0, '[]', 0, 0, 0
                        FROM `TemporaryNrDatas`
                        WHERE `ProjectId` = {projectId}");

                    // Clean up temp table
                    await _context.TemporaryNrDatas.Where(x => x.ProjectId == projectId).ExecuteDeleteAsync();

                    await transaction.CommitAsync();

                    return Ok(new { message = "Data pushed to main table successfully." });
                }
                catch
                {
                    await transaction.RollbackAsync();
                    throw;
                }
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
        }

        private async Task BulkInsertTemporaryNrDatasAsync(List<TemporaryNrDatas> tempDatas)
        {
            if (tempDatas == null || !tempDatas.Any()) return;

            const int chunkSize = 1000;
            var connection = _context.Database.GetDbConnection();
            bool wasOpen = connection.State == System.Data.ConnectionState.Open;
            if (!wasOpen) await connection.OpenAsync();

            try
            {
                for (int i = 0; i < tempDatas.Count; i += chunkSize)
                {
                    var chunk = tempDatas.Skip(i).Take(chunkSize).ToList();
                    var sb = new StringBuilder();
                    sb.Append(@"INSERT INTO `TemporaryNrDatas` 
                        (`ProjectId`, `CatchNo`, `NRQuantity`, `CollegeCode`, `CollegeName`, `CenterCode`, `NodalCode`, `CourseName`, `SubjectName`, `ExamDate`, `ExamTime`, `Day`, `NRDatas`, `Pages`, `Route`, `RouteSort`, `CenterSort`, `NodalSort`, `Symbol`, `Status`, `District`, `DistrictSort`) 
                        VALUES ");

                    for (int j = 0; j < chunk.Count; j++)
                    {
                        var item = chunk[j];
                        if (j > 0) sb.Append(",");

                        sb.Append($"({item.ProjectId}," +
                                  $"{EscapeSql(item.CatchNo)}," +
                                  $"{item.NRQuantity}," +
                                  $"{item.CollegeCode}," +
                                  $"{EscapeSql(item.CollegeName)}," +
                                  $"{EscapeSql(item.CenterCode)}," +
                                  $"{EscapeSql(item.NodalCode)}," +
                                  $"{EscapeSql(item.CourseName)}," +
                                  $"{EscapeSql(item.SubjectName)}," +
                                  $"{EscapeSql(item.ExamDate)}," +
                                  $"{EscapeSql(item.ExamTime)}," +
                                  $"{EscapeSql(item.Day)}," +
                                  $"{EscapeSql(item.NRDatas)}," +
                                  $"{item.Pages}," +
                                  $"{EscapeSql(item.Route)}," +
                                  $"{item.RouteSort}," +
                                  $"{item.CenterSort.ToString(CultureInfo.InvariantCulture)}," +
                                  $"{item.NodalSort.ToString(CultureInfo.InvariantCulture)}," +
                                  $"{EscapeSql(item.Symbol)}," +
                                  $"{(item.Status ? 1 : 0)}," +
                                  $"{EscapeSql(item.District)}," +
                                  $"{item.DistrictSort})");
                    }

                    using var cmd = connection.CreateCommand();
                    cmd.CommandText = sb.ToString();
                    cmd.CommandTimeout = 300;
                    await cmd.ExecuteNonQueryAsync();
                }
            }
            finally
            {
                if (!wasOpen) await connection.CloseAsync();
            }
        }

        private static string EscapeSql(string? str)
        {
            if (str == null) return "NULL";
            return "'" + str.Replace(@"\", @"\\").Replace("'", "''") + "'";
        }
    }
}
