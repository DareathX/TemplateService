namespace TemplateService
{
    public interface IFillInTemplates
    {

        public string[] supportedFileExtensions { get; }

        /// <summary>
        /// Get MemoryStream from filled in template.
        /// </summary>
        /// <param name="fileInByteArray">Template in a byte array.</param>
        /// <param name="placeholders">Dictionary of all the placeholders.</param>
        /// <param name="emptyOnNotFound">Throws error when placeholder is not found in given placeholders or removes the placeholder from docuemnt.</param>
        /// <param name="regexExpression">Set own regex expression</param>
        /// <param name="creatorName">Set name of creator of file</param>
        /// <returns>Returns a filled in template in a MemoryStream.</returns>
        public MemoryStream GetMemoryStream(byte[] fileInByteArray, Dictionary<string, object> placeholders, bool emptyOnNotFound, string regexExpression, string creatorName);
    }
}
