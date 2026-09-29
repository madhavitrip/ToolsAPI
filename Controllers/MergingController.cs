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
                // Check if any unresolved active Rule 1 conflicts exist in database
                var activeRule1Conflicts = await _context.ConflictingFields
                    .AsNoTracking()
                    .Where(c => c.ProjectId == projectId && c.Status == true && c.Rule == 1)
                    .AnyAsync();

                if (activeRule1Conflicts)
                {
                    return BadRequest(new
                    {
                        message = "Validation failed: Unresolved Rule 1 conflicts exist in Conflict Report. Please resolve them first.",
                        errors = new List<string> { "Please resolve all pending Rule 1 conflicts in the Conflict Report tab before generating preview." }
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
                    if ((!string.IsNullOrWhiteSpace(n.ExamCenterCode))) return n.ExamCenterCode;
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
                        getCatchKey = c => ((!string.IsNullOrWhiteSpace(c.CenterCode)) ? c.CenterCode : (ExtractCodeNumber(c.CenterName) ?? c.CollegeCode.ToString())).Trim();
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
                    if ((!string.IsNullOrWhiteSpace(n.ExamCenterCode))) return n.ExamCenterCode;
                    if (!string.IsNullOrWhiteSpace(n.ExamCenterName))
                    {
                        var code = ExtractCodeNumber(n.ExamCenterName);
                        if (code != null && !code.StartsWith("center_")) return code;
                    }
                    return null;
                }

                static string? GetValidNodalCode(NodalList n)
                {
                    if ((!string.IsNullOrWhiteSpace(n.NodalCode))) return n.NodalCode;
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
                        var catchCenterKey = (!string.IsNullOrWhiteSpace(catchItem.CenterCode)) ? catchItem.CenterCode : ExtractCodeNumber(catchItem.CenterName);
                        if (!string.IsNullOrEmpty(catchCenterKey))
                        {
                            var centerNodals = nodalLists.Where(n => 
                                GetEffectiveCenterCode(n) == catchCenterKey || 
                                n.ExamCenterCode == catchCenterKey || 
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
                        var centerCodeStr = (!string.IsNullOrWhiteSpace(catchItem.CenterCode)) ? catchItem.CenterCode : (ExtractCodeNumber(catchItem.CenterName) ?? "");
                        tempDatas.Add(new TemporaryNrDatas
                        {
                            ProjectId = projectId,
                            CatchNo = catchItem.CatchNo,
                            NRQuantity = catchItem.NRQuantity,
                            CollegeCode = catchItem.CollegeCode,
                            CollegeName = catchItem.CollegeName,
                            CenterCode = centerCodeStr,
                            NodalCode = "", // Not assigned to any nodal center!
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
                            var centerCodeStr = GetValidCenterCode(nodalItem) ?? ((!string.IsNullOrWhiteSpace(nodalItem.ExamCenterCode)) ? nodalItem.ExamCenterCode : ((!string.IsNullOrWhiteSpace(catchItem.CenterCode)) ? catchItem.CenterCode : ""));
                            var nodalCodeStr = GetValidNodalCode(nodalItem) ?? ((!string.IsNullOrWhiteSpace(nodalItem.NodalCode)) ? nodalItem.NodalCode : centerCodeStr);

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

                // Dynamic Rule 2 Validation
                var (levels1, levels2, levels3) = await GetConfiguredLevelsAsync(projectId, null, null, null);
                var (rule2Passed, rule2Errors, rule2CollegeCodes, unassignedDetails) = EvaluateDynamicRule2(
                    levels1,
                    levels2,
                    levels3,
                    tempDatas,
                    catchLists,
                    nodalLists);

                await SyncConflictsToDbAsync(projectId, new List<object>(), new List<object>(), unassignedDetails);

                if (!rule2Passed)
                {
                    return Ok(new
                    {
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

        private static string GetValDynamic(CatchList? c, NodalList? n, TemporaryNrDatas? t, string field)
        {
            if (string.IsNullOrWhiteSpace(field)) return "";
            var f = field.Trim().ToLowerInvariant().Replace(" ", "").Replace("_", "");

            switch (f)
            {
                case "collegecode":
                    if (n != null && n.CollegeCode != 0) return n.CollegeCode.ToString();
                    if (c != null && c.CollegeCode != 0) return c.CollegeCode.ToString();
                    if (t != null && t.CollegeCode != 0) return t.CollegeCode.ToString();
                    if (n != null && !string.IsNullOrEmpty(n.CollegeName))
                    {
                        var match = System.Text.RegularExpressions.Regex.Match(n.CollegeName, @"^(\d+)");
                        if (match.Success) return match.Groups[1].Value;
                    }
                    if (c != null && !string.IsNullOrEmpty(c.CollegeName))
                    {
                        var match = System.Text.RegularExpressions.Regex.Match(c.CollegeName, @"^(\d+)");
                        if (match.Success) return match.Groups[1].Value;
                    }
                    if (t != null && !string.IsNullOrEmpty(t.CollegeName))
                    {
                        var match = System.Text.RegularExpressions.Regex.Match(t.CollegeName, @"^(\d+)");
                        if (match.Success) return match.Groups[1].Value;
                    }
                    return "";

                case "collegename":
                    if (n != null && !string.IsNullOrEmpty(n.CollegeName)) return n.CollegeName;
                    if (c != null && !string.IsNullOrEmpty(c.CollegeName)) return c.CollegeName;
                    if (t != null && !string.IsNullOrEmpty(t.CollegeName)) return t.CollegeName;
                    return "";

                case "examcentercode":
                case "centercode":
                    if (n != null && !string.IsNullOrWhiteSpace(n.ExamCenterCode)) return n.ExamCenterCode.Trim();
                    if (c != null && !string.IsNullOrWhiteSpace(c.CenterCode)) return c.CenterCode.Trim();
                    if (t != null && !string.IsNullOrWhiteSpace(t.CenterCode)) return t.CenterCode.Trim();
                    if (n != null && !string.IsNullOrEmpty(n.ExamCenterName))
                    {
                        var match = System.Text.RegularExpressions.Regex.Match(n.ExamCenterName, @"^(\d+)");
                        if (match.Success) return match.Groups[1].Value;
                    }
                    if (c != null && !string.IsNullOrEmpty(c.CenterName))
                    {
                        var match = System.Text.RegularExpressions.Regex.Match(c.CenterName, @"^(\d+)");
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
                    if (t != null && !string.IsNullOrWhiteSpace(t.NodalCode)) return t.NodalCode.Trim();
                    if (n != null && !string.IsNullOrWhiteSpace(n.NodalCode)) return n.NodalCode.Trim();
                    if (n != null && !string.IsNullOrEmpty(n.NodalName))
                    {
                        var match = System.Text.RegularExpressions.Regex.Match(n.NodalName, @"^(\d+)");
                        if (match.Success) return match.Groups[1].Value;
                    }
                    return "";

                case "nodalname":
                    if (t != null && !string.IsNullOrWhiteSpace(t.NodalCode)) return t.NodalCode.Trim();
                    if (n != null && !string.IsNullOrEmpty(n.NodalName)) return n.NodalName;
                    return "";

                case "catchno":
                    if (c != null && !string.IsNullOrEmpty(c.CatchNo)) return c.CatchNo;
                    if (t != null && !string.IsNullOrEmpty(t.CatchNo)) return t.CatchNo;
                    return "";

                case "nrquantity":
                case "quantity":
                    if (c != null) return c.NRQuantity.ToString();
                    if (t != null) return t.NRQuantity.ToString();
                    return "";

                case "coursename":
                case "course":
                    if (c != null && !string.IsNullOrEmpty(c.CourseName)) return c.CourseName;
                    if (t != null && !string.IsNullOrEmpty(t.CourseName)) return t.CourseName;
                    return "";

                case "subjectname":
                case "subject":
                    if (c != null && !string.IsNullOrEmpty(c.SubjectName)) return c.SubjectName;
                    if (t != null && !string.IsNullOrEmpty(t.SubjectName)) return t.SubjectName;
                    return "";

                case "examdate":
                    if (c != null && !string.IsNullOrEmpty(c.ExamDate)) return c.ExamDate;
                    if (t != null && !string.IsNullOrEmpty(t.ExamDate)) return t.ExamDate;
                    return "";

                case "examtime":
                    if (c != null && !string.IsNullOrEmpty(c.ExamTime)) return c.ExamTime;
                    if (t != null && !string.IsNullOrEmpty(t.ExamTime)) return t.ExamTime;
                    return "";

                default:
                    if (t != null && !string.IsNullOrEmpty(t.NRDatas))
                    {
                        try
                        {
                            using var doc = JsonDocument.Parse(t.NRDatas);
                            foreach (var prop in doc.RootElement.EnumerateObject())
                            {
                                if (prop.Name.Trim().ToLowerInvariant().Replace(" ", "").Replace("_", "") == f)
                                {
                                    return prop.Value.ToString();
                                }
                            }
                        }
                        catch { }
                    }
                    if (c != null && !string.IsNullOrEmpty(c.NRDatas))
                    {
                        try
                        {
                            using var doc = JsonDocument.Parse(c.NRDatas);
                            foreach (var prop in doc.RootElement.EnumerateObject())
                            {
                                if (prop.Name.Trim().ToLowerInvariant().Replace(" ", "").Replace("_", "") == f)
                                {
                                    return prop.Value.ToString();
                                }
                            }
                        }
                        catch { }
                    }
                    if (n != null && !string.IsNullOrEmpty(n.OtherFields))
                    {
                        try
                        {
                            using var doc = JsonDocument.Parse(n.OtherFields);
                            foreach (var prop in doc.RootElement.EnumerateObject())
                            {
                                if (prop.Name.Trim().ToLowerInvariant().Replace(" ", "").Replace("_", "") == f)
                                {
                                    return prop.Value.ToString();
                                }
                            }
                        }
                        catch { }
                    }

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
                    if (t != null)
                    {
                        var p = typeof(TemporaryNrDatas).GetProperty(field, System.Reflection.BindingFlags.IgnoreCase | System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance);
                        if (p != null) return p.GetValue(t)?.ToString() ?? "";
                    }
                    return "";
            }
        }

        private static string GetKeyDynamic(CatchList? c, NodalList? n, TemporaryNrDatas? t, List<string> fields)
        {
            if (fields == null || !fields.Any()) return "";
            var vals = fields.Select(f => GetValDynamic(c, n, t, f).Trim()).ToList();
            if (vals.All(v => string.IsNullOrEmpty(v))) return "";
            return string.Join(" | ", vals);
        }

        private async Task<(List<string> Level1, List<string> Level2, List<string> Level3)> GetConfiguredLevelsAsync(
            int projectId,
            string? queryL1,
            string? queryL2,
            string? queryL3)
        {
            List<string> l1 = !string.IsNullOrWhiteSpace(queryL1)
                ? queryL1.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).ToList()
                : new List<string>();

            List<string> l2 = !string.IsNullOrWhiteSpace(queryL2)
                ? queryL2.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).ToList()
                : new List<string>();

            List<string> l3 = queryL3 != null
                ? queryL3.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).ToList()
                : new List<string>();

            bool hasQuery = !string.IsNullOrWhiteSpace(queryL1) || !string.IsNullOrWhiteSpace(queryL2) || queryL3 != null;

            if (!hasQuery)
            {
                var configRec = await _context.ConflictingFields
                    .AsNoTracking()
                    .FirstOrDefaultAsync(c => c.ProjectId == projectId && c.UniqueField == "Rule_Level_Config");

                if (configRec != null && !string.IsNullOrEmpty(configRec.ConflictingField))
                {
                    try
                    {
                        using var doc = JsonDocument.Parse(configRec.ConflictingField);
                        var root = doc.RootElement;
                        if (root.TryGetProperty("level1", out var p1) && p1.ValueKind == JsonValueKind.Array)
                        {
                            l1 = p1.EnumerateArray().Select(x => x.GetString() ?? "").Where(x => !string.IsNullOrEmpty(x)).ToList();
                        }
                        if (root.TryGetProperty("level2", out var p2) && p2.ValueKind == JsonValueKind.Array)
                        {
                            l2 = p2.EnumerateArray().Select(x => x.GetString() ?? "").Where(x => !string.IsNullOrEmpty(x)).ToList();
                        }
                        if (root.TryGetProperty("level3", out var p3) && p3.ValueKind == JsonValueKind.Array)
                        {
                            l3 = p3.EnumerateArray().Select(x => x.GetString() ?? "").Where(x => !string.IsNullOrEmpty(x)).ToList();
                        }
                    }
                    catch { }
                }
            }

            if (!l1.Any() && !hasQuery) l1 = new List<string> { "CollegeCode" };
            if (!l2.Any() && !hasQuery) l2 = new List<string> { "ExamCenterCode" };

            return (l1, l2, l3);
        }

        private static (bool passed, List<string> errors, List<string> codes, List<object> details) EvaluateDynamicRule2(
            List<string> level1,
            List<string> level2,
            List<string> level3,
            List<TemporaryNrDatas> tempDatas,
            List<CatchList> catchLists,
            List<NodalList> nodalLists)
        {
            var unassignedDetails = new List<object>();
            var rule2Codes = new List<string>();

            if (level1 == null || !level1.Any() || level2 == null || !level2.Any())
            {
                return (true, new List<string>(), rule2Codes, unassignedDetails);
            }

            bool hasLevel3Config = level3 != null && level3.Any();
            bool hasTemp = tempDatas != null && tempDatas.Any();

            if (hasTemp)
            {
                var groupedL1 = tempDatas
                    .GroupBy(t => GetKeyDynamic(null, null, t, level1))
                    .Where(g => !string.IsNullOrEmpty(g.Key))
                    .ToList();

                foreach (var g in groupedL1)
                {
                    string l1Key = g.Key;
                    var firstItem = g.First();

                    var l2FromTemp = g.Select(x => GetKeyDynamic(null, null, x, level2)).Where(x => !string.IsNullOrEmpty(x)).Distinct().ToList();
                    bool hasLevel2 = l2FromTemp.Any();

                    bool hasLevel3 = true;
                    if (hasLevel3Config)
                    {
                        var l3FromTemp = g.Select(x => GetKeyDynamic(null, null, x, level3)).Where(x => !string.IsNullOrEmpty(x)).Distinct().ToList();
                        hasLevel3 = l3FromTemp.Any();
                    }

                    if (!hasLevel2 || !hasLevel3)
                    {
                        rule2Codes.Add(l1Key);
                        string missingTarget = !hasLevel2 ? string.Join("+", level2) : string.Join("+", level3);
                        string cName = GetValDynamic(null, null, firstItem, "CollegeName");
                        if (string.IsNullOrEmpty(cName)) cName = GetValDynamic(null, null, firstItem, "CollegeCode");
                        if (string.IsNullOrEmpty(cName)) cName = l1Key;

                        unassignedDetails.Add(new
                        {
                            collegeCode = l1Key,
                            collegeName = cName,
                            courses = g.Select(x => GetValDynamic(null, null, x, "CourseName")).Where(x => !string.IsNullOrEmpty(x)).Distinct().ToList(),
                            subjects = g.Select(x => GetValDynamic(null, null, x, "SubjectName")).Where(x => !string.IsNullOrEmpty(x)).Distinct().ToList(),
                            centerCodes = l2FromTemp,
                            centerNames = new List<string>(),
                            totalQuantity = g.Sum(x => x.NRQuantity),
                            description = $"{string.Join("+", level1)} ({l1Key}) is missing assignment for {missingTarget} in Temporary NR Data."
                        });
                    }
                }
            }

            rule2Codes = rule2Codes.Distinct().ToList();
            bool passed = !unassignedDetails.Any();
            var errors = passed
                ? new List<string>()
                : new List<string> { $"Rule 2 Failed: {rule2Codes.Count} record(s) missing assignment: {string.Join(", ", rule2Codes)}" };

            return (passed, errors, rule2Codes, unassignedDetails);
        }

        private static (bool passed, List<string> errors, List<string> codes, List<object> details) EvaluateRule3GenderMismatch(
            List<string> level1,
            List<string> level2,
            List<string> level3,
            List<CatchList> catchLists)
        {
            var rule3Details = new List<object>();
            var rule3Codes = new List<string>();

            if (catchLists == null || !catchLists.Any())
            {
                return (true, new List<string>(), rule3Codes, rule3Details);
            }

            bool hasGenderFieldInConfig = (level1 != null && level1.Any(f => f.Trim().Equals("Gender", StringComparison.OrdinalIgnoreCase))) ||
                                          (level2 != null && level2.Any(f => f.Trim().Equals("Gender", StringComparison.OrdinalIgnoreCase))) ||
                                          (level3 != null && level3.Any(f => f.Trim().Equals("Gender", StringComparison.OrdinalIgnoreCase)));

            bool hasGenderDataInCatch = catchLists.Any(c => c.Male > 0 || c.Female > 0 || c.Transgender > 0);

            if (hasGenderFieldInConfig || hasGenderDataInCatch)
            {
                var genderMismatches = catchLists
                    .Where(c => c.NRQuantity != (c.Male + c.Female))
                    .ToList();

                foreach (var item in genderMismatches)
                {
                    int sumGender = item.Male + item.Female;
                    string err = $"Gender Quantity Mismatch: Catch No '{item.CatchNo}' for College '{item.CollegeCode}' ({item.CollegeName}) has total NR Quantity {item.NRQuantity}, but Male ({item.Male}) + Female ({item.Female}) = {sumGender}.";
                    string key = $"Rule3_GenderMismatch_{item.CatchNo}_{item.CollegeCode}";

                    rule3Codes.Add(key);
                    rule3Details.Add(new
                    {
                        catchNo = item.CatchNo ?? "",
                        collegeCode = item.CollegeCode,
                        collegeName = item.CollegeName ?? "",
                        nrQuantity = item.NRQuantity,
                        male = item.Male,
                        female = item.Female,
                        transgender = item.Transgender,
                        genderSum = sumGender,
                        description = err
                    });
                }
            }

            rule3Codes = rule3Codes.Distinct().ToList();
            bool passed = !rule3Details.Any();
            var errors = passed
                ? new List<string>()
                : new List<string> { $"Rule 3 Failed: {rule3Codes.Count} record(s) gender quantity mismatch." };

            return (passed, errors, rule3Codes, rule3Details);
        }

        [HttpPost("CheckDynamicRule1/{projectId}")]
        public async Task<IActionResult> CheckDynamicRule1(int projectId, [FromBody] DynamicRule1Request req)
        {
            try
            {
                var nodalLists = await _context.NodalList.AsNoTracking().Where(x => x.ProjectId == projectId && x.Status).ToListAsync();
                var catchLists = await _context.CatchList.AsNoTracking().Where(x => x.ProjectId == projectId && x.Status).ToListAsync();
                var tempDatas = await _context.TemporaryNrDatas.AsNoTracking().Where(x => x.ProjectId == projectId).ToListAsync();

                // Save dynamic level configuration to DB for this project
                string configKey = "Rule_Level_Config";
                var levelConfigJson = JsonSerializer.Serialize(new
                {
                    level1 = req.Level1 ?? new List<string>(),
                    level2 = req.Level2 ?? new List<string>(),
                    level3 = req.Level3 ?? new List<string>()
                });

                var existingConfig = await _context.ConflictingFields
                    .FirstOrDefaultAsync(c => c.ProjectId == projectId && c.UniqueField == configKey);
                if (existingConfig != null)
                {
                    existingConfig.ConflictingField = levelConfigJson;
                    existingConfig.Status = true;
                }
                else
                {
                    _context.ConflictingFields.Add(new ConflictingFields
                    {
                        ProjectId = projectId,
                        UniqueField = configKey,
                        ConflictingField = levelConfigJson,
                        Status = true,
                        Rule = 0
                    });
                }

                // Build pair list for Rule 1 checking
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

                var newConflicts = new List<ConflictingFields>();

                // Check Level 1 -> Level 2
                if (req.Level1 != null && req.Level1.Any() && req.Level2 != null && req.Level2.Any())
                {
                    var g1 = unifiedList.GroupBy(u => GetKeyDynamic(u.Catch, u.Nodal, null, req.Level1)).ToList();
                    foreach (var g in g1)
                    {
                        if (string.IsNullOrEmpty(g.Key)) continue;
                        var l2Keys = g.Select(u => GetKeyDynamic(u.Catch, u.Nodal, null, req.Level2)).Where(k => !string.IsNullOrEmpty(k)).Distinct().ToList();
                        if (l2Keys.Count > 1)
                        {
                            newConflicts.Add(new ConflictingFields
                            {
                                ProjectId = projectId,
                                UniqueField = JsonSerializer.Serialize(new { fields = req.Level1, value = g.Key }),
                                ConflictingField = JsonSerializer.Serialize(new { fields = req.Level2, values = l2Keys }),
                                Status = true,
                                Rule = 1
                            });
                        }
                    }
                }

                // Check Level 2 -> Level 3
                if (req.Level2 != null && req.Level2.Any() && req.Level3 != null && req.Level3.Any())
                {
                    var g2 = unifiedList.GroupBy(u => GetKeyDynamic(u.Catch, u.Nodal, null, req.Level2)).ToList();
                    foreach (var g in g2)
                    {
                        if (string.IsNullOrEmpty(g.Key)) continue;
                        var l3Keys = g.Select(u => GetKeyDynamic(u.Catch, u.Nodal, null, req.Level3)).Where(k => !string.IsNullOrEmpty(k)).Distinct().ToList();
                        if (l3Keys.Count > 1)
                        {
                            newConflicts.Add(new ConflictingFields
                            {
                                ProjectId = projectId,
                                UniqueField = JsonSerializer.Serialize(new { fields = req.Level2, value = g.Key }),
                                ConflictingField = JsonSerializer.Serialize(new { fields = req.Level3, values = l3Keys }),
                                Status = true,
                                Rule = 1
                            });
                        }
                    }
                }

                var existingInDb = await _context.ConflictingFields
                    .Where(c => c.ProjectId == projectId && c.Rule == 1)
                    .ToListAsync();
                var activeUniqueKeys = new HashSet<string>();

                foreach (var nc in newConflicts)
                {
                    activeUniqueKeys.Add(nc.UniqueField);
                    var match = existingInDb.FirstOrDefault(e => e.UniqueField == nc.UniqueField);
                    if (match != null)
                    {
                        match.ConflictingField = nc.ConflictingField;
                        match.Status = true;
                    }
                    else
                    {
                        _context.ConflictingFields.Add(nc);
                    }
                }

                foreach (var c in existingInDb)
                {
                    if (!activeUniqueKeys.Contains(c.UniqueField))
                    {
                        c.Status = false;
                    }
                }

                // Evaluate Dynamic Rule 2
                var (r2Passed, r2Errors, r2Codes, r2Details) = EvaluateDynamicRule2(
                    req.Level1 ?? new List<string>(),
                    req.Level2 ?? new List<string>(),
                    req.Level3 ?? new List<string>(),
                    tempDatas,
                    catchLists,
                    nodalLists);

                await SyncConflictsToDbAsync(projectId, new List<object>(), new List<object>(), r2Details);

                await _context.SaveChangesAsync();

                return Ok(new
                {
                    message = "Dynamic Rule 1 & Rule 2 validated",
                    conflicts = newConflicts.Count,
                    rule1Passed = !newConflicts.Any(),
                    rule2Passed = r2Passed,
                    rule2Errors = r2Errors,
                    rule2CollegeCodes = r2Codes,
                    unassignedDetails = r2Details
                });
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
                await _context.ConflictingFields
                    .Where(c => c.ProjectId == projectId && (c.Rule == 1 || c.Rule == 2))
                    .ExecuteUpdateAsync(s => s.SetProperty(c => c.Status, false));
                return Ok(new { message = "Dynamic conflicts cleared successfully" });
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
        }

        [HttpGet("Reports/{projectId}")]
        public async Task<IActionResult> GetReports(
            int projectId,
            [FromQuery] string? mergeBy = "CollegeCode",
            [FromQuery] string? level1 = null,
            [FromQuery] string? level2 = null,
            [FromQuery] string? level3 = null)
        {
            try
            {
                var dbRule1Conflicts = await _context.ConflictingFields
                    .AsNoTracking()
                    .Where(c => c.ProjectId == projectId && c.Status == true && c.Rule == 1)
                    .ToListAsync();

                var dbRule2Conflicts = await _context.ConflictingFields
                    .AsNoTracking()
                    .Where(c => c.ProjectId == projectId && c.Status == true && c.Rule == 2)
                    .ToListAsync();

                var catchLists = await _context.CatchList
                    .AsNoTracking()
                    .Where(x => x.ProjectId == projectId && x.Status)
                    .ToListAsync();

                var nodalLists = await _context.NodalList
                    .AsNoTracking()
                    .Where(x => x.ProjectId == projectId && x.Status)
                    .ToListAsync();

                var tempDatas = await _context.TemporaryNrDatas
                    .AsNoTracking()
                    .Where(x => x.ProjectId == projectId)
                    .ToListAsync();

                var (levels1, levels2, levels3) = await GetConfiguredLevelsAsync(projectId, level1, level2, level3);

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
                    if (!string.IsNullOrWhiteSpace(n.ExamCenterCode)) return n.ExamCenterCode;
                    if (!string.IsNullOrEmpty(n.ExamCenterName))
                    {
                        var code = ExtractCodeNumber(n.ExamCenterName);
                        if (code != null) return code;
                        return n.ExamCenterName.Trim().ToLowerInvariant();
                    }
                    return $"center_{n.Id}";
                }

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

                var dbRule3Conflicts = await _context.ConflictingFields
                    .AsNoTracking()
                    .Where(c => c.ProjectId == projectId && c.Status == true && c.Rule == 3)
                    .ToListAsync();

                // Dynamic Rule 2 Validation
                var (rule2Passed, rule2Errors, rule2CollegeCodes, unassignedDetails) = EvaluateDynamicRule2(
                    levels1,
                    levels2,
                    levels3,
                    tempDatas,
                    catchLists,
                    nodalLists);

                // Rule 3 Gender Quantity Mismatch Validation
                var (rule3Passed, rule3Errors, rule3CollegeCodes, rule3Details) = EvaluateRule3GenderMismatch(
                    levels1,
                    levels2,
                    levels3,
                    catchLists);

                await SyncConflictsToDbAsync(projectId, multiCenterDetails, multiNodalDetails, unassignedDetails, rule3Details);

                bool rule1Passed = !rule1Errors.Any() && !dbRule1Conflicts.Any();

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

                    rule3Passed,
                    rule3Errors,
                    rule3CollegeCodes,
                    rule3Count = rule3Details.Count,
                    rule3Details,

                    dbRule1Conflicts,
                    dbRule2Conflicts,
                    dbRule3Conflicts,

                    configuredLevels = new
                    {
                        level1 = levels1,
                        level2 = levels2,
                        level3 = levels3
                    }
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
            List<object> unassignedDetails,
            List<object>? rule3Details = null)
        {
            try
            {
                var existingInDb = await _context.ConflictingFields
                    .Where(c => c.ProjectId == projectId)
                    .ToListAsync();

                var activeUniqueKeys = new HashSet<string>();

                // Rule 2: Unassigned Colleges / Centers
                foreach (var item in unassignedDetails)
                {
                    var json = JsonSerializer.Serialize(item);
                    using var doc = JsonDocument.Parse(json);
                    var root = doc.RootElement;
                    string cCode = root.TryGetProperty("collegeCode", out var cc) ? cc.GetString() ?? "" : "";
                    if (string.IsNullOrEmpty(cCode) || cCode == "0" || cCode == "Unknown") continue;

                    string uniqueKey = $"Rule2_Unassigned_{cCode}";
                    activeUniqueKeys.Add(uniqueKey);

                    var match = existingInDb.FirstOrDefault(e => e.UniqueField == uniqueKey);
                    if (match != null)
                    {
                        match.ConflictingField = json;
                        match.Status = true;
                    }
                    else
                    {
                        _context.ConflictingFields.Add(new ConflictingFields
                        {
                            ProjectId = projectId,
                            UniqueField = uniqueKey,
                            ConflictingField = json,
                            Status = true,
                            Rule = 2
                        });
                    }
                }

                // Rule 3: Gender Quantity Mismatch
                if (rule3Details != null)
                {
                    foreach (var item in rule3Details)
                    {
                        var json = JsonSerializer.Serialize(item);
                        using var doc = JsonDocument.Parse(json);
                        var root = doc.RootElement;
                        string cNo = root.TryGetProperty("catchNo", out var cn) ? cn.GetString() ?? "" : "";
                        int cCode = root.TryGetProperty("collegeCode", out var cc) ? cc.GetInt32() : 0;
                        string uniqueKey = $"Rule3_GenderMismatch_{cNo}_{cCode}";
                        activeUniqueKeys.Add(uniqueKey);

                        var match = existingInDb.FirstOrDefault(e => e.UniqueField == uniqueKey);
                        if (match != null)
                        {
                            match.ConflictingField = json;
                            match.Status = true;
                        }
                        else
                        {
                            _context.ConflictingFields.Add(new ConflictingFields
                            {
                                ProjectId = projectId,
                                UniqueField = uniqueKey,
                                ConflictingField = json,
                                Status = true,
                                Rule = 3
                            });
                        }
                    }
                }

                // Soft-resolve (Status = 0) active Rule 1, Rule 2, Rule 3 conflicts no longer failing
                foreach (var ec in existingInDb)
                {
                    if (ec.UniqueField != null && (ec.UniqueField.StartsWith("Rule1_") || ec.UniqueField.StartsWith("Rule2_") || ec.UniqueField.StartsWith("Rule3_")))
                    {
                        if (!activeUniqueKeys.Contains(ec.UniqueField) && ec.Status == true)
                        {
                            ec.Status = false;
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
                    .CountAsync(c => c.ProjectId == projectId && c.Status == true && (c.Rule == 1 || c.Rule == 2 || c.Rule == 3) && c.UniqueField != "Rule_Level_Config");

                return Ok(new { allResolved = pendingCount == 0 });
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
        }

        public class ResolveGenderMismatchDto
        {
            public int ProjectId { get; set; }
            public object? ConflictId { get; set; }
            public object? CatchNo { get; set; }
            public object? CollegeCode { get; set; }
            public object? NRQuantity { get; set; }
            public object? Male { get; set; }
            public object? Female { get; set; }
            public object? Transgender { get; set; }

            public int GetConflictIdInt() => ConvertToInt(ConflictId);
            public string GetCatchNoString() => CatchNo?.ToString() ?? "";
            public int GetCollegeCodeInt() => ConvertToInt(CollegeCode);
            public int GetNRQuantityInt() => ConvertToInt(NRQuantity);
            public int GetMaleInt() => ConvertToInt(Male);
            public int GetFemaleInt() => ConvertToInt(Female);
            public int GetTransgenderInt() => ConvertToInt(Transgender);

            private static int ConvertToInt(object? val)
            {
                if (val == null) return 0;
                if (val is JsonElement elem)
                {
                    if (elem.ValueKind == JsonValueKind.Number && elem.TryGetInt32(out int iVal)) return iVal;
                    if (elem.ValueKind == JsonValueKind.String && int.TryParse(elem.GetString(), out int sVal)) return sVal;
                }
                if (int.TryParse(val.ToString(), out int parsed)) return parsed;
                return 0;
            }
        }

        [HttpPost("ResolveGenderMismatch")]
        public async Task<IActionResult> ResolveGenderMismatch([FromBody] ResolveGenderMismatchDto dto)
        {
            try
            {
                int conflictId = dto.GetConflictIdInt();
                string catchNo = dto.GetCatchNoString();
                int collegeCode = dto.GetCollegeCodeInt();
                int nrQty = dto.GetNRQuantityInt();
                int male = dto.GetMaleInt();
                int female = dto.GetFemaleInt();
                int transgender = dto.GetTransgenderInt();

                if (conflictId > 0)
                {
                    var conflict = await _context.ConflictingFields.FirstOrDefaultAsync(c => c.Id == conflictId);
                    if (conflict != null)
                    {
                        conflict.Status = false;
                    }
                }
                else if (!string.IsNullOrEmpty(catchNo))
                {
                    var conflicts = await _context.ConflictingFields
                        .Where(c => c.ProjectId == dto.ProjectId && c.Rule == 3 && c.Status == true && c.UniqueField.Contains(catchNo))
                        .ToListAsync();
                    foreach (var c in conflicts) c.Status = false;
                }

                var catchItems = await _context.CatchList
                    .Where(c => c.ProjectId == dto.ProjectId && c.Status && c.CatchNo == catchNo && (collegeCode == 0 || c.CollegeCode == collegeCode))
                    .ToListAsync();

                foreach (var c in catchItems)
                {
                    c.NRQuantity = nrQty;
                    c.Male = male;
                    c.Female = female;
                    c.Transgender = transgender;
                }

                var tempDatas = await _context.TemporaryNrDatas
                    .Where(t => t.ProjectId == dto.ProjectId && t.CatchNo == catchNo && (collegeCode == 0 || t.CollegeCode == collegeCode))
                    .ToListAsync();

                foreach (var t in tempDatas)
                {
                    t.NRQuantity = nrQty;
                }

                await _context.SaveChangesAsync();
                return Ok(new { message = "Gender quantity updated and conflict resolved successfully" });
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
                    conflict.Status = false;
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
                    conflict.Status = false;
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
                    var matching = nodalLists.FirstOrDefault(x => x.NodalCode == parsedTargetInt.ToString() && !string.IsNullOrWhiteSpace(x.NodalName));
                    if (matching != null) resolvedNodalName = matching.NodalName;
                }

                string resolvedCenterName = dto.TargetName;
                if (string.IsNullOrWhiteSpace(resolvedCenterName) && (targetF.Contains("center") || targetF.Contains("examcenter")) && parsedTargetInt > 0)
                {
                    var matching = nodalLists.FirstOrDefault(x => x.ExamCenterCode == parsedTargetInt.ToString() && !string.IsNullOrWhiteSpace(x.ExamCenterName));
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
                        isMatch = (parsedMatchInt > 0 && n.ExamCenterCode == parsedMatchInt.ToString().ToString()) ||
                                  (!string.IsNullOrEmpty(n.ExamCenterName) && (n.ExamCenterName.Contains(matchVal) || (parsedMatchInt > 0 && n.ExamCenterName.StartsWith(parsedMatchInt.ToString()))));
                    }
                    else if (matchF.Contains("nodal"))
                    {
                        isMatch = (parsedMatchInt > 0 && n.NodalCode == parsedMatchInt.ToString()) ||
                                  (!string.IsNullOrEmpty(n.NodalName) && (n.NodalName.Contains(matchVal) || (parsedMatchInt > 0 && n.NodalName.StartsWith(parsedMatchInt.ToString()))));
                    }
                    else
                    {
                        isMatch = (parsedMatchInt > 0 && (n.CollegeCode == parsedMatchInt || n.ExamCenterCode == parsedMatchInt.ToString().ToString() || n.NodalCode == parsedMatchInt.ToString()));
                    }

                    if (isMatch)
                    {
                        if (targetF.Contains("nodal"))
                        {
                            n.NodalCode = parsedTargetInt > 0 ? parsedTargetInt.ToString() : n.NodalCode;
                            if (!string.IsNullOrWhiteSpace(resolvedNodalName))
                            {
                                n.NodalName = resolvedNodalName;
                            }
                            updatedAny = true;
                        }
                        else if (targetF.Contains("center") || targetF.Contains("examcenter"))
                        {
                            n.ExamCenterCode = parsedTargetInt > 0 ? parsedTargetInt.ToString() : n.ExamCenterCode;
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
                    if ((!string.IsNullOrWhiteSpace(n.ExamCenterCode))) return n.ExamCenterCode;
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
                    if ((!string.IsNullOrWhiteSpace(n.NodalCode))) return n.NodalCode;
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
                            var centerCodeStr = GetValidCenterCode(validNodal) ?? ((!string.IsNullOrWhiteSpace(validNodal.ExamCenterCode)) ? validNodal.ExamCenterCode : "");
                            var nodalCodeStr = GetValidNodalCode(validNodal) ?? ((!string.IsNullOrWhiteSpace(validNodal.NodalCode)) ? validNodal.NodalCode : "");

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
                    .Where(x => {
                        bool hasCollegeCode = x.CollegeCode != 0 || ExtractCodeNumber(x.CollegeName) != null;
                        bool hasCenterCode = !string.IsNullOrWhiteSpace(x.CenterCode) && x.CenterCode != "0";
                        // Ignore records with missing required college and center codes
                        if (!hasCollegeCode && !hasCenterCode) return false;

                        return string.IsNullOrWhiteSpace(x.CenterCode) || x.CenterCode == "0" || string.IsNullOrWhiteSpace(x.NodalCode) || x.NodalCode == "0";
                    })
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