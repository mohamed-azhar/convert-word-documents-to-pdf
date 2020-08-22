using System;
using System.IO;
using System.Linq;

namespace ConvertFromWordToPDF.Helpers
{
    public class FileHelpers
    {
        public static bool IsValidWordFile(string extension) => extension == ".doc" || extension == ".docx";

        public static string GetFileName(string path) => path.Split("\\").LastOrDefault().Split(".").FirstOrDefault();

        public static string GetDirectoryFromFilePath(string filePath)
        {
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
            return directoryInformation.EnumerateFiles(null, SearchOption.TopDirectoryOnly).Select(x => x.FullName).ToArray();
        }
    }
}
