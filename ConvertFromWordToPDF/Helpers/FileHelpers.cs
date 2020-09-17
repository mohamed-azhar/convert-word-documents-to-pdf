using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ConvertFromWordToPDF.Helpers
{
    public class FileHelpers
    {
        public static bool IsValidWordFile(string extension)
        {
            if (string.IsNullOrWhiteSpace(extension))
            {
                throw new ArgumentNullException(nameof(extension));
            }

            return extension == ".doc" || extension == ".docx";
        }

        public static string GetFileName(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                throw new ArgumentNullException(nameof(path));
            }

            return path.Split("\\").LastOrDefault().Split(".").FirstOrDefault();
        }

        public static string GetDirectoryFromFilePath(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                throw new ArgumentNullException(nameof(filePath));
            }

            var splitted = filePath.Split("\\").Reverse().Skip(1).ToArray();
            return string.Join("\\", splitted.Reverse());
        }

        public static string[] GetFilesInDirectory(string directory)
        {
            if (string.IsNullOrWhiteSpace(directory))
            {
                throw new ArgumentNullException(nameof(directory));
            }

            var directoryInformation = new DirectoryInfo(directory);
            return directoryInformation.EnumerateFiles(string.Empty, SearchOption.TopDirectoryOnly).Select(x => x.FullName).ToArray();
        }

        public static string[] ExtractValidWordFiles(string[] paths)
        {
            if (paths?.Any() ?? false)
            {
                var filtered = new List<string>();

                foreach (var path in paths)
                {
                    if (IsValidWordFile($".{path.Split(".").LastOrDefault()}"))
                    {
                        filtered.Add(path);
                    }
                }

                return filtered.ToArray();
            }
            return null;
        }
    }
}
