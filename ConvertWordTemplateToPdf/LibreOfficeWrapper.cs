using TemplateService;
using System.Diagnostics;
using System.Linq;

namespace ConvertWordTemplateToPdf
{
    internal class LibreOfficeWrapper
    {
        private TempStorage tempStorage = new TempStorage();

        /// <summary>
        /// Get LibreOffice Portable path. If folder path is not given, it will be set to the program output path.
        /// </summary>
        /// <param name="libreOfficeFolderPath">Path to the folder of LibreOffice.</param>
        /// <param name="pluginFolderPath">Folder where unzipped LibreOffice will be saved.</param>
        /// <returns>Returns path with the LibreOffice executable.</returns>
        private string GetLibreOfficePath(string libreOfficeFolderPath, string pluginFolderPath)
        {
            string libreOfficePath;
            if (string.IsNullOrEmpty(libreOfficeFolderPath))
            {
                libreOfficeFolderPath = Directory.GetCurrentDirectory() + @"\LibreOfficePortable";
                libreOfficePath = Directory.GetCurrentDirectory() + @"\LibreOfficePortable\App\libreoffice\program\soffice.exe";

                if (!Directory.Exists(libreOfficeFolderPath))
                {
                    ExtractLibreOfficeZip(libreOfficeFolderPath, pluginFolderPath);
                }
            }
            else
            {
                libreOfficePath = libreOfficeFolderPath + @"\App\libreoffice\program\soffice.exe";
            }

            return libreOfficePath;
        }

        /// <summary>
        /// Converts Word documents to Pdf by using LibreOffice. Uses a temporary storage to edit the files.
        /// </summary>
        /// <param name="inputStream">Template in stream format.</param>
        /// <param name="libreOfficeFolderPath">Path to the folder of LibreOffice</param>
        /// <param name="tempFolderPath">Path to temporary folder.</param>
        /// <param name="pluginFolderPath">Folder where unzipped LibreOffice will be saved, if it's not set.</param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public byte[] Convert(MemoryStream inputStream, string libreOfficeFolderPath, string tempFolderPath, string pluginFolderPath)
        {
            string libreOfficePath = GetLibreOfficePath(libreOfficeFolderPath, pluginFolderPath);

            string tempFolder = tempStorage.CreateTempFolder(tempFolderPath);
            string[] folderContents = tempStorage.GetFolderFileNames(tempFolder);

            int number = 0;
            for (int i = 0; i < folderContents.Length; i++)//Check names so that the documents don't get mixed up while multiple edits are happening.
            {
                if (!folderContents[i].Contains("#")) continue;

                int newNumber = int.Parse(folderContents[i].Split('#').Last());
                number = newNumber >= number ? newNumber + 1: number;
            }

            string inputFile = tempFolder + $@"\Document#{number}.docx";
            string outputFile = tempFolder + $@"\Document#{number}.pdf";
            
            File.WriteAllBytes(inputFile, inputStream.ToArray());

            List<string> commandArgs = new List<string>()
            {
                "--convert-to",
                "pdf:writer_pdf_Export",
                inputFile,
                "--norestore", 
                "--writer", 
                "--headless", 
                "--outdir", 
                tempFolder
            };

            ProcessStartInfo procStartInfo = new ProcessStartInfo(libreOfficePath);
            foreach (string arg in commandArgs) 
            { 
                procStartInfo.ArgumentList.Add(arg); 
            }

            Process process = new Process() 
            { 
                StartInfo = procStartInfo 
            };

            process.Start();
            process.WaitForExit();

            if (process.ExitCode != 0)
            {
                throw new Exception("LibreOffice returned the following error code: " + process.ExitCode);
            }

            byte[] output = File.ReadAllBytes(outputFile);

            //Cleans the temporary storage.
            tempStorage.DeleteFile(inputFile);
            tempStorage.DeleteFile(outputFile);
            return output;
        }

        /// <summary>
        /// Extracts LibreOffice out of zip. This happens only when LibreOffice is not found.
        /// </summary>
        /// <param name="libreOfficeFolderPath">Path to the LibreOffice folder. </param>
        /// <param name="pluginFolderPath">Folder where unzipped LibreOffice will be saved.</param>
        private void ExtractLibreOfficeZip(string libreOfficeFolderPath, string pluginFolderPath)
        {
            new ExtractZip(pluginFolderPath + @"\LibreOfficePortable.zip", libreOfficeFolderPath);
        }
    }
}
