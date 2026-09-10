using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;
using Microsoft.AspNetCore.Http;

namespace Tools.Services
{
    public static class AuditJsonHelper
    {
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            ReferenceHandler = ReferenceHandler.IgnoreCycles,
            DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
            WriteIndented = false,
            Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
        };

        private static readonly HashSet<string> SensitiveKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "password", "passcode", "secret", "secretkey", "token", "accesstoken", "authorization"
        };

        public static string? SerializeSafe(object? value, int maxCharacters = 32768)
        {
            if (value == null)
            {
                return null;
            }

            if (value is string strValue)
            {
                if (string.IsNullOrWhiteSpace(strValue)) return null;
                return TruncateIfNeeded(strValue, maxCharacters);
            }

            try
            {
                // Handle IFormFile
                if (value is IFormFile file)
                {
                    return JsonSerializer.Serialize(new
                    {
                        FileType = "IFormFile",
                        FileName = file.FileName,
                        Length = file.Length,
                        ContentType = file.ContentType
                    }, JsonOptions);
                }

                // Handle IFormFileCollection or IEnumerable<IFormFile>
                if (value is IEnumerable<IFormFile> files)
                {
                    var fileList = files.Select(f => new
                    {
                        FileType = "IFormFile",
                        FileName = f.FileName,
                        Length = f.Length,
                        ContentType = f.ContentType
                    }).ToList();

                    return JsonSerializer.Serialize(fileList, JsonOptions);
                }

                // Handle streams
                if (value is Stream stream)
                {
                    return JsonSerializer.Serialize(new
                    {
                        StreamType = value.GetType().Name,
                        CanRead = stream.CanRead,
                        Length = stream.CanSeek ? stream.Length : -1
                    }, JsonOptions);
                }

                var json = JsonSerializer.Serialize(value, JsonOptions);
                return TruncateIfNeeded(json, maxCharacters);
            }
            catch (Exception ex)
            {
                return $"{{\"serialization_error\": \"{ex.Message}\"}}";
            }
        }

        public static int ExtractProjectId(object? model)
        {
            if (model == null) return 0;

            try
            {
                var type = model.GetType();

                // Check ProjectId
                var projProp = type.GetProperty("ProjectId", BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);
                if (projProp != null)
                {
                    var val = projProp.GetValue(model);
                    if (val is int intVal && intVal > 0) return intVal;
                    if (val != null && int.TryParse(val.ToString(), out var parsedInt) && parsedInt > 0) return parsedInt;
                }

                // Check GroupId as fallback if ProjectId is not present
                var groupProp = type.GetProperty("GroupId", BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);
                if (groupProp != null)
                {
                    var val = groupProp.GetValue(model);
                    if (val is int intVal && intVal > 0) return intVal;
                    if (val != null && int.TryParse(val.ToString(), out var parsedInt) && parsedInt > 0) return parsedInt;
                }
            }
            catch
            {
                // Suppress reflection errors
            }

            return 0;
        }

        private static string TruncateIfNeeded(string text, int maxCharacters)
        {
            if (string.IsNullOrEmpty(text) || text.Length <= maxCharacters)
            {
                return text;
            }

            return text.Substring(0, maxCharacters) + $"... [Truncated {text.Length - maxCharacters} chars]";
        }
    }
}
