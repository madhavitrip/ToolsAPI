
using ERPToolsAPI.Data;
using ERPToolsAPI.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Net.Http.Json;
using System.Text.RegularExpressions;

[ApiController]
[Route("api/[controller]")]
public class CorrectionController : ControllerBase
{
    private readonly IHttpClientFactory _httpClientFactory;
    private readonly ERPToolsDbContext _context;
    private readonly string _baseUrl;

    public CorrectionController(IHttpClientFactory httpClientFactory, ERPToolsDbContext context, IConfiguration configuration)
    {
        _httpClientFactory = httpClientFactory;
        _context = context;
        _baseUrl = configuration["ApiSettings:BaseUrl"];
    }

    /* [HttpPost("correct")]
     public async Task<IActionResult> CorrectCourseOnly([FromBody] CorrectionPayload payload)
     {

         var client = _httpClientFactory.CreateClient();

         var rows = await _context.ExcelUploads
             .Where(x => x.GroupId == payload.GroupId).
             ToListAsync();

         var courseList = await client.GetFromJsonAsync<List<CourseMasterItem>>($"{_baseUrl}/Course");
         var subjectList = await client.GetFromJsonAsync<List<SubjectMasterItem>>($"{_baseUrl}/Subject");
         var languageList = await client.GetFromJsonAsync<List<LanguageMasterItem>>($"{_baseUrl}/Language");
         var examTypeList = await client.GetFromJsonAsync<List<ExamTypeMasterItem>>($"{_baseUrl}/ExamType");
         var paperList = await client.GetFromJsonAsync<List<PaperTypeMasterItem>>($"{_baseUrl}/PaperTypes");

         var changesSummary = new List<object>();
         int correctedCount = 0;

         foreach (var row in rows)
         {
             var changes = new Dictionary<string, (string OldValue, string NewValue)>();

             if (!string.IsNullOrWhiteSpace(row.Course))
             {
                 row.Course = MatchCourse(row.Course, courseList);
             }
             if (!string.IsNullOrWhiteSpace(row.Subject))
             {
                 row.Subject = MatchSubject(row.Subject, subjectList);
             }
             if (!string.IsNullOrWhiteSpace(row.Language))
             {
                 row.Language = MatchLanguage(row.Language, languageList);
             }
             if (!string.IsNullOrWhiteSpace(row.ExamType))
             {
                 row.ExamType = MatchExamType(row.ExamType, examTypeList);
             }
             if (!string.IsNullOrWhiteSpace(row.Type))
             {
                 row.Type = MatchPaperType(row.Type, paperList);
             }
         }

         await _context.SaveChangesAsync();
         return Ok("Course corrections applied successfully.");
     }
 */
    /*
        [HttpPost("correct")]
        public async Task<IActionResult> CorrectCourseOnly([FromBody] CorrectionPayload payload)
        {
            var client = _httpClientFactory.CreateClient();

            var rows = await _context.ExcelUploads
                .Where(x => x.GroupId == payload.GroupId)
                .ToListAsync();

            var courseList = await client.GetFromJsonAsync<List<CourseMasterItem>>($"{_baseUrl}/Course");
            var subjectList = await client.GetFromJsonAsync<List<SubjectMasterItem>>($"{_baseUrl}/Subject");
            var languageList = await client.GetFromJsonAsync<List<LanguageMasterItem>>($"{_baseUrl}/Language");
            var examTypeList = await client.GetFromJsonAsync<List<ExamTypeMasterItem>>($"{_baseUrl}/ExamType");
            var paperList = await client.GetFromJsonAsync<List<PaperTypeMasterItem>>($"{_baseUrl}/PaperTypes");

            var changesSummary = new List<object>();
            int correctedCount = 0;

            foreach (var row in rows)
            {
                var changes = new Dictionary<string, (string OldValue, string NewValue)>();

                // Always Correct Everything
                var oldCourse = row.Course ?? "";
                var newCourse = MatchCourse(row.Course, courseList);
                if (!string.Equals(oldCourse, newCourse, StringComparison.OrdinalIgnoreCase))
                {
                    row.Course = newCourse;
                }
                if (payload.FieldsToCorrect.Contains("course", StringComparer.OrdinalIgnoreCase) &&
                    !string.Equals(oldCourse, newCourse, StringComparison.OrdinalIgnoreCase))
                    changes["Course"] = (oldCourse, newCourse);

                var oldSubject = row.Subject ?? "";
                var newSubject = MatchSubject(row.Subject, subjectList);
                if (!string.Equals(oldSubject, newSubject, StringComparison.OrdinalIgnoreCase))
                    row.Subject = newSubject;
                if (payload.FieldsToCorrect.Contains("subject", StringComparer.OrdinalIgnoreCase) &&
                    !string.Equals(oldSubject, newSubject, StringComparison.OrdinalIgnoreCase))
                    changes["Subject"] = (oldSubject, newSubject);

                var oldLanguage = row.Language ?? "";
                var newLanguage = MatchLanguage(row.Language, languageList);
                if (!string.Equals(oldLanguage, newLanguage, StringComparison.OrdinalIgnoreCase))
                    row.Language = newLanguage;
                if (payload.FieldsToCorrect.Contains("language", StringComparer.OrdinalIgnoreCase) &&
                    !string.Equals(oldLanguage, newLanguage, StringComparison.OrdinalIgnoreCase))
                    changes["Language"] = (oldLanguage, newLanguage);

                var oldExamType = row.ExamType ?? "";
                var newExamType = MatchExamType(row.ExamType, examTypeList);
                if (!string.Equals(oldExamType, newExamType, StringComparison.OrdinalIgnoreCase))
                    row.ExamType = newExamType;
                if (payload.FieldsToCorrect.Contains("examtype", StringComparer.OrdinalIgnoreCase) &&
                    !string.Equals(oldExamType, newExamType, StringComparison.OrdinalIgnoreCase))
                    changes["ExamType"] = (oldExamType, newExamType);

                var oldType = row.Type ?? "";
                var newType = MatchPaperType(row.Type, paperList);
                if (!string.Equals(oldType, newType, StringComparison.OrdinalIgnoreCase))
                    row.Type = newType;
                if (payload.FieldsToCorrect.Contains("type", StringComparer.OrdinalIgnoreCase) &&
                    !string.Equals(oldType, newType, StringComparison.OrdinalIgnoreCase))
                    changes["Type"] = (oldType, newType);

                if (changes.Any())
                {
                    correctedCount++;
                    changesSummary.Add(new
                    {
                        row.Id,
                        row.Catch,
                        ChangedFields = changes.ToDictionary(c => c.Key, c => new { Old = c.Value.OldValue, New = c.Value.NewValue })
                    });
                }
            }

            await _context.SaveChangesAsync();

            return Ok(new
            {
                CorrectedCount = correctedCount,
                CorrectedRows = changesSummary
            });
        }
    */



    [HttpPost("correct")]
    public async Task<IActionResult> CorrectCourseOnly([FromBody] CorrectionPayload payload)
    {
        var client = _httpClientFactory.CreateClient();

        var rows = await _context.ExcelUploads
            .Where(x => x.GroupId == payload.GroupId)
            .ToListAsync();

        var courseList = await client.GetFromJsonAsync<List<CourseMasterItem>>($"{_baseUrl}/Course");
        var subjectList = await client.GetFromJsonAsync<List<SubjectMasterItem>>($"{_baseUrl}/Subject");
        var languageList = await client.GetFromJsonAsync<List<LanguageMasterItem>>($"{_baseUrl}/Language");
        var examTypeList = await client.GetFromJsonAsync<List<ExamTypeMasterItem>>($"{_baseUrl}/ExamType");
        var paperList = await client.GetFromJsonAsync<List<PaperTypeMasterItem>>($"{_baseUrl}/PaperTypes");

        var changesSummary = new List<object>();
        int correctedCount = 0;

        foreach (var row in rows)
        {
            var changes = new Dictionary<string, (string OldValue, string NewValue)>();

            // Course Correction
            var oldCourse = row.Course ?? "";
            if (!string.IsNullOrWhiteSpace(oldCourse))
            {
                var newCourse = MatchCourse(oldCourse, courseList);
                row.Course = newCourse;
                if (payload.FieldsToCorrect.Contains("course", StringComparer.OrdinalIgnoreCase) &&
                    !string.Equals(oldCourse, newCourse, StringComparison.Ordinal)) // case-sensitive difference
                {
                    changes["Course"] = (oldCourse, newCourse);
                }
            }

            // Subject Correction
            var oldSubject = row.Subject ?? "";
            if (!string.IsNullOrWhiteSpace(oldSubject))
            {
                var newSubject = MatchSubject(oldSubject, subjectList);
                row.Subject = newSubject;
                if (payload.FieldsToCorrect.Contains("subject", StringComparer.OrdinalIgnoreCase) &&
                    !string.Equals(oldSubject, newSubject, StringComparison.Ordinal))
                {
                    changes["Subject"] = (oldSubject, newSubject);
                }
            }

            // Language Correction
            var oldLanguage = row.Language ?? "";
            if (!string.IsNullOrWhiteSpace(oldLanguage))
            {
                var newLanguage = MatchLanguage(oldLanguage, languageList);
                row.Language = newLanguage;
                if (payload.FieldsToCorrect.Contains("language", StringComparer.OrdinalIgnoreCase) &&
                    !string.Equals(oldLanguage, newLanguage, StringComparison.Ordinal))
                {
                    changes["Language"] = (oldLanguage, newLanguage);
                }
            }

            // ExamType Correction
            var oldExamType = row.ExamType ?? "";
            if (!string.IsNullOrWhiteSpace(oldExamType))
            {
                var newExamType = MatchExamType(oldExamType, examTypeList);
                row.ExamType = newExamType;
                if (payload.FieldsToCorrect.Contains("examtype", StringComparer.OrdinalIgnoreCase) &&
                    !string.Equals(oldExamType, newExamType, StringComparison.Ordinal))
                {
                    changes["ExamType"] = (oldExamType, newExamType);
                }
            }

            // Type Correction (Paper Type)
            var oldType = row.Type ?? "";
            if (!string.IsNullOrWhiteSpace(oldType))
            {
                var newType = MatchPaperType(oldType, paperList);
                row.Type = newType;
                if (payload.FieldsToCorrect.Contains("type", StringComparer.OrdinalIgnoreCase) &&
                    !string.Equals(oldType, newType, StringComparison.Ordinal))
                {
                    changes["Type"] = (oldType, newType);
                }
            }

            if (changes.Any())
            {
                correctedCount++;
                changesSummary.Add(new
                {
                    row.Id,
                    row.Catch,
                    ChangedFields = changes.ToDictionary(c => c.Key, c => new { Old = c.Value.OldValue, New = c.Value.NewValue })
                });
            }
        }

        await _context.SaveChangesAsync();

        return Ok(new
        {
            CorrectedCount = correctedCount,
            CorrectedRows = changesSummary
        });
    }


    private string MatchCourse(string value, List<CourseMasterItem>? masterList)
    {
        if (string.IsNullOrWhiteSpace(value) || masterList == null)
            return value;

        string normalizedValue = Normalize(value);

        foreach (var item in masterList)
        {
            if (string.IsNullOrWhiteSpace(item.CourseName)) continue;

            string normalizedMaster = Normalize(item.CourseName);

            if (normalizedValue == normalizedMaster)
                return item.CourseName; // Return proper casing and punctuation from master
        }

        return value; // No match, return original
    }


    [HttpPost("updateRows")]
    public async Task<IActionResult> UpdateRows([FromBody] List<ExcelUpload> updatedRows)
    {
        foreach (var updatedRow in updatedRows)
        {
            var row = await _context.ExcelUploads.FirstOrDefaultAsync(x => x.Id == updatedRow.Id);
            if (row != null)
            {
                row.Course = updatedRow.Course;
                row.Subject = updatedRow.Subject;
                row.Type = updatedRow.Type;
                row.Language = updatedRow.Language;
                row.PaperNumber = updatedRow.PaperNumber;
                row.ExamType = updatedRow.ExamType;
                row.Catch = updatedRow.Catch;
            }
        }
        await _context.SaveChangesAsync();
        return Ok("Rows updated");
    }

    private string MatchSubject(string value, List<SubjectMasterItem>? masterList)
    {
        if (string.IsNullOrWhiteSpace(value) || masterList == null)
            return value;

        string normalizedValue = Normalize(value);

        foreach (var item in masterList)
        {
            if (string.IsNullOrWhiteSpace(item.SubjectName)) continue;

            string normalizedMaster = Normalize(item.SubjectName);

            if (normalizedValue == normalizedMaster)
                return item.SubjectName; // Return proper casing and punctuation from master
        }

        return value; // No match, return original
    }

    private string MatchLanguage(string value, List<LanguageMasterItem>? masterList)
    {
        if (string.IsNullOrWhiteSpace(value) || masterList == null)
            return value;

        string normalizedValue = Normalize(value);

        foreach (var item in masterList)
        {
            if (string.IsNullOrWhiteSpace(item.Languages)) continue;

            string normalizedMaster = Normalize(item.Languages);

            if (normalizedValue == normalizedMaster)
                return item.Languages; // Return proper casing and punctuation from master
        }

        return value; // No match, return original
    }
    private string MatchExamType(string value, List<ExamTypeMasterItem>? masterList)
    {
        if (string.IsNullOrWhiteSpace(value) || masterList == null)
            return value;

        string cleanedInput = PreprocessExamType(value);
        Console.WriteLine($"Cleaned Input: {cleanedInput}");
        // Try exact match first
        foreach (var item in masterList)
        {
            string cleanedMaster = PreprocessExamType(item.TypeName);
            Console.WriteLine($"Cleaned Master: {cleanedMaster}");
            if (cleanedInput == cleanedMaster)
                return item.TypeName;
        }

        // Fuzzy match fallback (>= 90% similarity)
        foreach (var item in masterList)
        {
            string cleanedMaster = PreprocessExamType(item.TypeName);
            double similarity = CalculateSimilarity(cleanedInput, cleanedMaster);
            if (similarity >= 0.90)
                return item.TypeName;
        }

        return value; // return original if no match
    }


    private string PreprocessExamType(string input)
    {
        if (string.IsNullOrWhiteSpace(input)) return "";

        string cleaned = input.ToLowerInvariant();

        // Standardization dictionary
        var replacements = new Dictionary<string, string>
    {
        { "sem\\.", "semester" },
        { "\\bsem\\b", "semester" }, // Match "sem" as a word
        { "\\bfirst\\b", "i" },
        { "\\bsecond\\b", "ii" },
        { "\\bthird\\b", "iii" },
        { "\\bfourth\\b", "iv" },
        { "\\bfifth\\b", "v" },
        { "\\bsixth\\b", "vi" },
        { "\\bseventh\\b", "vii" },
        { "\\beighth\\b", "viii" },
            {"\\bninth\\b", "ix" },
            { "\\btenth\\b", "x" },
    };

        foreach (var pair in replacements)
        {
            cleaned = System.Text.RegularExpressions.Regex.Replace(cleaned, pair.Key, pair.Value, System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        }

        // Replace non-alphanumeric characters (except spaces)
        cleaned = cleaned.Replace("-", " ")
            .Replace("Part-", "")
            .Replace("Previous", "")
            .Replace("Final", "")
                         .Replace("(", "")
                         .Replace(")", "");

        // Collapse multiple spaces
        cleaned = System.Text.RegularExpressions.Regex.Replace(cleaned, @"\s+", " ").Trim();

        return cleaned;
    }

    private string MatchPaperType(string value, List<PaperTypeMasterItem>? masterList)
    {
        if (string.IsNullOrWhiteSpace(value) || masterList == null)
            return value;

        string normalizedValue = TypeNormalize(value);
        string bestMatch = value;
        double highestScore = 0;

        foreach (var item in masterList)
        {
            if (string.IsNullOrWhiteSpace(item.Types)) continue;

            string normalizedMaster = TypeNormalize(item.Types);
            double similarity = CalculateSimilarity(normalizedValue, normalizedMaster);

            if (similarity >= 0.90 && similarity > highestScore)
            {
                highestScore = similarity;
                bestMatch = item.Types;
            }
        }

        return bestMatch;
    }

    private double CalculateSimilarity(string source, string target)
    {
        if (string.IsNullOrEmpty(source) || string.IsNullOrEmpty(target))
            return 0;

        int distance = LevenshteinDistance(source, target);
        int maxLen = Math.Max(source.Length, target.Length);
        return maxLen == 0 ? 1.0 : 1.0 - (double)distance / maxLen;
    }


    private int LevenshteinDistance(string s, string t)
    {
        int[,] d = new int[s.Length + 1, t.Length + 1];

        for (int i = 0; i <= s.Length; i++) d[i, 0] = i;
        for (int j = 0; j <= t.Length; j++) d[0, j] = j;

        for (int i = 1; i <= s.Length; i++)
        {
            for (int j = 1; j <= t.Length; j++)
            {
                int cost = s[i - 1] == t[j - 1] ? 0 : 1;
                d[i, j] = Math.Min(
                    Math.Min(d[i - 1, j] + 1, d[i, j - 1] + 1),
                    d[i - 1, j - 1] + cost
                );
            }
        }

        return d[s.Length, t.Length];
    }

    [HttpGet("audit")]
    public async Task<IActionResult> AuditGroupData([FromQuery] int groupId)
    {
        var client = _httpClientFactory.CreateClient();

        var rows = await _context.ExcelUploads
            .Where(x => x.GroupId == groupId)
            .ToListAsync();

        var courseList = await client.GetFromJsonAsync<List<CourseMasterItem>>($"{_baseUrl}/Course");
        var subjectList = await client.GetFromJsonAsync<List<SubjectMasterItem>>($"{_baseUrl}/Subject");
        var languageList = await client.GetFromJsonAsync<List<LanguageMasterItem>>($"{_baseUrl}/Language");
        var examTypeList = await client.GetFromJsonAsync<List<ExamTypeMasterItem>>($"{_baseUrl}/ExamType");
        var paperList = await client.GetFromJsonAsync<List<PaperTypeMasterItem>>($"{_baseUrl}/PaperTypes");

        var auditResults = new List<object>();
        int mismatchCount = 0;

        foreach (var row in rows)
        {
            var mismatches = new Dictionary<string, string>();

            if (!string.IsNullOrWhiteSpace(row.Course) &&
                !courseList.Any(c => Normalize(c.CourseName) == Normalize(row.Course)))
                mismatches.Add("Course", row.Course);

            if (!string.IsNullOrWhiteSpace(row.Subject) &&
                !subjectList.Any(s => Normalize(s.SubjectName) == Normalize(row.Subject)))
                mismatches.Add("Subject", row.Subject);

            if (!string.IsNullOrWhiteSpace(row.Language) &&
                !languageList.Any(l => Normalize(l.Languages) == Normalize(row.Language)))
                mismatches.Add("Language", row.Language);

            if (!string.IsNullOrWhiteSpace(row.ExamType) &&
                !examTypeList.Any(e => Normalize(e.TypeName) == Normalize(row.ExamType)))
                mismatches.Add("ExamType", row.ExamType);

            if (!string.IsNullOrWhiteSpace(row.Type) &&
                !paperList.Any(p => Normalize(p.Types) == Normalize(row.Type)))
                mismatches.Add("Type", row.Type);

            if (mismatches.Any())
            {
                mismatchCount++;
                auditResults.Add(new
                {
                    row.Id,
                    row.Catch,
                    row.Course,
                    row.Subject,
                    row.Type,
                    row.Language,
                    row.PaperNumber,
                    row.ExamType,
                    row.GroupId,
                    Mismatches = mismatches
                });
            }
        }

        return Ok(new
        {
            MismatchCount = mismatchCount,
            MismatchedRows = auditResults
        });

    }



    private string Normalize(string value)
    {
        return value?.ToLower().Replace(".", "").Replace(" ", "").Trim() ?? "";
    }


    private string TypeNormalize(string value)
    {
        if (string.IsNullOrWhiteSpace(value)) return "";

        value = value.ToLower().Replace(".", "").Replace(" ", "").Trim();

        // Collapse repeated characters (e.g., "boookkklet" => "boklet")
        return Regex.Replace(value, @"(\w)\1+", "$1");
    }


   

    public class CorrectionPayload
    {
        public int GroupId { get; set; }
        public List<string> FieldsToCorrect { get; set; } = new();
    }

    public class CourseMasterItem
    {
        public int Id { get; set; }
        public string CourseName { get; set; } = "";
    }

    public class SubjectMasterItem
    {
        public int Id { get; set; }
        public string SubjectName { get; set; } = "";
    }

    public class LanguageMasterItem
    {
        public int Id { get; set; }
        public string Languages { get; set; } = "";
    }

    public class ExamTypeMasterItem
    {
        public int Id { get; set; }
        public string TypeName { get; set; } = "";
    }

    public class PaperTypeMasterItem
    {
        public int Id { get; set; }
        public string Types { get; set; } = "";
    }

}
