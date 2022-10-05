using TemplateService;

namespace ConvertWordTemplateToPdf
{
    public class ConvertWordTemplateToPdf : IConvertTemplateTo
    {
        public string[] supportedFileExtensions => new string[] { ".pdf" };
        LibreOfficeWrapper LOWrapper = new LibreOfficeWrapper();

        /// <summary>
        /// Converts Word template to pdf by using a LibbreOffice wrapper.
        /// </summary>
        /// <param name="stream">Word template as MemoryStream</param>
        /// <param name="additionalData">Dictionary with folder locations needed for the conversion</param>
        /// <returns>The converted document</returns>
        public byte[]? Convert(MemoryStream stream, Dictionary<string, string> additionalData)
        {
            return LOWrapper.Convert(stream, additionalData["pdfLibraryFolderPath"], additionalData["tempFolderPath"], additionalData["pluginFolderPath"]);
        }

        /// <summary>
        /// Converts filled in template to requested document.
        /// </summary>
        /// <param name="stream">Filled in template as a MemoryStream.</param>
        /// <returns>The converted document</returns>
        public byte[]? Convert(MemoryStream stream)//Not needed for pdf conversion.
        {
            return stream.ToArray();
        }
    }
}