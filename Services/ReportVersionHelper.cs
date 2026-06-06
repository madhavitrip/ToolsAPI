using System;
using System.IO;
using System.Text.RegularExpressions;

namespace Tools.Services
{
    public static class ReportVersionHelper
    {
        public static string GetNextVersionFileName(string directory, string baseFileName)
        {
            if (string.IsNullOrWhiteSpace(baseFileName))
                return baseFileName;

            string extension = Path.GetExtension(baseFileName);
            string nameWithoutExt = Path.GetFileNameWithoutExtension(baseFileName);

            int maxVersion = 0;

            if (Directory.Exists(directory))
            {
                // Pattern matches nameWithoutExt_v{number}.extension (e.g. DuplicateTool_v2.xlsx)
                var pattern = new Regex(@"^" + Regex.Escape(nameWithoutExt) + @"_v(\d+)" + Regex.Escape(extension) + "$", RegexOptions.IgnoreCase);
                foreach (var file in Directory.GetFiles(directory))
                {
                    var match = pattern.Match(Path.GetFileName(file));
                    if (match.Success)
                    {
                        if (int.TryParse(match.Groups[1].Value, out int v))
                        {
                            if (v > maxVersion)
                            {
                                maxVersion = v;
                            }
                        }
                    }
                }
            }

            // If maxVersion is 0, check if the unversioned file exists
            if (maxVersion == 0)
            {
                var unversionedPath = Path.Combine(directory, baseFileName);
                if (File.Exists(unversionedPath))
                {
                    try
                    {
                        var v1Path = Path.Combine(directory, $"{nameWithoutExt}_v1{extension}");
                        if (!File.Exists(v1Path))
                        {
                            File.Move(unversionedPath, v1Path);
                        }
                        maxVersion = 1;
                    }
                    catch
                    {
                        maxVersion = 0;
                    }
                }
            }

            int nextVersion = maxVersion + 1;
            return $"{nameWithoutExt}_v{nextVersion}{extension}";
        }
    }
}
