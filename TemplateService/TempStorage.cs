namespace TemplateService
{
    public class TempStorage
    {
        /// <summary>
        /// Creates new folder which will be used as temporary storage.
        /// </summary>
        /// <param name="tempFolder">Path to desired folder location.</param>
        /// <returns>Returns path to temporary folder.</returns>
        public string CreateTempFolder(string tempFolder)
        {
            tempFolder = tempFolder == string.Empty ? Directory.GetCurrentDirectory() + @"\temp" : tempFolder;
            if (!Directory.Exists(tempFolder))
            {
                tempFolder = Directory.CreateDirectory(tempFolder).FullName;
            }

            return tempFolder;
        }

        /// <summary>
        /// Get all file name from folder path.
        /// </summary>
        /// <param name="path">Path to folder to get all file names.</param>
        /// <returns>Returns array with all file names.</returns>
        public string[] GetFolderFileNames(string path)
        {
            return Directory.GetFiles(path).Select(Path.GetFileNameWithoutExtension).ToArray();
        }

        /// <summary>
        /// Removes file.
        /// </summary>
        /// <param name="filePath">Path to the file to remove</param>
        public void DeleteFile(string filePath)
        {
            File.Delete(filePath);
        }
    }
}
