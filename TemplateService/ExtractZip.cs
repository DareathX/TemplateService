using System.IO.Compression;

namespace TemplateService
{
    public class ExtractZip
    {
        /// <summary>
        /// Extracts .zip to desired path.
        /// </summary>
        /// <param name="zipPath">Path of the .zip</param>
        /// <param name="extractPath">Path to extract</param>
        public ExtractZip(string zipPath, string extractPath)
        {
            if (File.Exists(zipPath))
            {
                ZipFile.ExtractToDirectory(zipPath, extractPath);
            }
        }
    }
}
