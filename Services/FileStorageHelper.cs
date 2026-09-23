using System;
using System.IO;
using Microsoft.Extensions.Configuration;

namespace Tools.Services
{
    public static class FileStorageHelper
    {
        private static string? _storageBasePath;
        private static readonly object _lock = new object();

        public static void Initialize(IConfiguration configuration)
        {
            var configuredPath = configuration["FileStorage:BasePath"]
                ?? configuration["StorageSettings:BasePath"]
                ?? @"E:\Files";

            _storageBasePath = ResolveAndEnsurePath(configuredPath);
        }

        public static string GetStorageBasePath(IConfiguration? configuration = null)
        {
            if (!string.IsNullOrEmpty(_storageBasePath))
                return _storageBasePath;

            lock (_lock)
            {
                if (!string.IsNullOrEmpty(_storageBasePath))
                    return _storageBasePath;

                var configuredPath = configuration?["FileStorage:BasePath"]
                    ?? configuration?["StorageSettings:BasePath"]
                    ?? @"E:\Files";

                _storageBasePath = ResolveAndEnsurePath(configuredPath);
                return _storageBasePath;
            }
        }

        private static string ResolveAndEnsurePath(string targetPath)
        {
            try
            {
                var root = Path.GetPathRoot(targetPath);
                if (!string.IsNullOrEmpty(root) && Directory.Exists(root))
                {
                    if (!Directory.Exists(targetPath))
                    {
                        Directory.CreateDirectory(targetPath);
                    }
                    return Path.GetFullPath(targetPath);
                }
            }
            catch
            {
                // Ignore and fall back
            }

            // Fallback to wwwroot in current directory if the target drive is unavailable
            var fallback = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot");
            try
            {
                if (!Directory.Exists(fallback))
                {
                    Directory.CreateDirectory(fallback);
                }
            }
            catch
            {
            }
            return fallback;
        }

        public static string GetProjectFolder(int projectId)
        {
            var basePath = GetStorageBasePath();
            var folder = Path.Combine(basePath, projectId.ToString());
            if (!Directory.Exists(folder))
            {
                Directory.CreateDirectory(folder);
            }
            return folder;
        }

        public static string GetGroupFolder(string folderName)
        {
            var basePath = GetStorageBasePath();
            var folder = Path.Combine(basePath, folderName);
            if (!Directory.Exists(folder))
            {
                Directory.CreateDirectory(folder);
            }
            return folder;
        }

        public static string GetRptTemplatesFolder()
        {
            var basePath = GetStorageBasePath();
            var folder = Path.Combine(basePath, "rpt-templates");
            if (!Directory.Exists(folder))
            {
                Directory.CreateDirectory(folder);
            }
            return folder;
        }

        public static string GetUploadsFolder(string? subFolder = null)
        {
            var basePath = GetStorageBasePath();
            var folder = string.IsNullOrWhiteSpace(subFolder)
                ? Path.Combine(basePath, "uploads")
                : Path.Combine(basePath, "uploads", subFolder);

            if (!Directory.Exists(folder))
            {
                Directory.CreateDirectory(folder);
            }
            return folder;
        }

        /// <summary>
        /// Looks for an existing file in the storage directory first, falling back to legacy wwwroot.
        /// </summary>
        public static string? FindExistingFilePath(string relativePath)
        {
            var basePath = GetStorageBasePath();
            var primaryPath = Path.Combine(basePath, relativePath);
            if (File.Exists(primaryPath))
                return primaryPath;

            var legacyPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", relativePath);
            if (File.Exists(legacyPath))
                return legacyPath;

            return null;
        }

        /// <summary>
        /// Converts an absolute path to a relative path by stripping the storage base directory.
        /// If the path is already relative, it is returned as-is (normalized).
        /// </summary>
        public static string GetRelativePath(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                return filePath;

            // Normalize separators
            filePath = filePath
                .Replace('/', Path.DirectorySeparatorChar)
                .Replace('\\', Path.DirectorySeparatorChar);

            if (!Path.IsPathRooted(filePath))
                return filePath.TrimStart(Path.DirectorySeparatorChar);

            var basePath = GetStorageBasePath();
            if (!basePath.EndsWith(Path.DirectorySeparatorChar.ToString()))
                basePath += Path.DirectorySeparatorChar;

            if (filePath.StartsWith(basePath, StringComparison.OrdinalIgnoreCase))
                return filePath.Substring(basePath.Length).TrimStart(Path.DirectorySeparatorChar);

            // Also strip legacy wwwroot path if present
            var wwwRoot = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot") + Path.DirectorySeparatorChar;
            if (filePath.StartsWith(wwwRoot, StringComparison.OrdinalIgnoreCase))
                return filePath.Substring(wwwRoot.Length).TrimStart(Path.DirectorySeparatorChar);

            // Path is absolute but doesn't match known bases — return filename + parent folder only
            return filePath.TrimStart(Path.DirectorySeparatorChar);
        }

        /// <summary>
        /// Ensures the given path is resolved to a fully qualified absolute path.
        /// Relative/unrooted paths are resolved against the storage base directory.
        /// </summary>
        public static string? GetAbsolutePath(string? filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                return filePath;

            if (Path.IsPathRooted(filePath))
            {
                return Path.GetFullPath(filePath);
            }

            var basePath = GetStorageBasePath();
            return Path.GetFullPath(Path.Combine(basePath, filePath));
        }
    }
}
