namespace TemplateService
{
    public interface IConvertTemplateTo
    {
        public string[] supportedFileExtensions { get; }

        /// <summary>
        /// Converts filled in template to requested document.
        /// </summary>
        /// <param name="stream">Filled in template as a MemoryStream.</param>
        /// <param name="additionalData">Dictionary with folder locations if they are needed for conversion</param>
        /// <returns>The converted document</returns>
        public byte[]? Convert(MemoryStream stream, Dictionary<string, string> additionalData);

        /// <summary>
        /// Converts filled in template to requested document.
        /// </summary>
        /// <param name="stream">Filled in template as a MemoryStream.</param>
        /// <returns>The converted document</returns>
        public byte[]? Convert(MemoryStream stream);
    }
}
