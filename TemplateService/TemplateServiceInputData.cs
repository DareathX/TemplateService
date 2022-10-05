namespace TemplateService
{
    public struct TemplateServiceInputData
    {
        /// <summary>
        /// Sets to only convert document and skips filling in document. Non-mandatory.
        /// </summary>
        public bool ConvertOnly { get; set; }

        /// <summary>
        /// Throws error if there is a placeholder in the template but not in the given Dictinoary. Non-mandatory.
        /// </summary>
        public bool EmptyOnNotFound { get; set; }

        /// <summary>
        /// Dictionary with all the placeholders and placeholder values. Mandatory.
        /// </summary>
        public Dictionary<string, object> InputDataDict { get; set; }

        /// <summary>
        /// Sets the conversion from a document to a document: (.docx, .pdf). 
        /// If conversion is not needed then set both as the template extension: (.docx, .docx). Mandatory.
        /// </summary>
        public (string, string)? SetConversionFromTo { get; set; }

        /// <summary>
        /// Sets the creator and editor in Word documents. Non-mandatory.
        /// </summary>
        public string SetDocumentCreatorName { get; set; }

        /// <summary>
        /// Set path for the output document. Recommanded if you don't want a byte array or MemoryStream in the return. Non-mandatory.
        /// </summary>
        public string SetOutputDocumentPath { get; set; }

        /// <summary>
        /// Overrides the Regex expression used to find placeholders in the templates. Recommanded to use braces: { } Non-mandatory.
        /// </summary>
        public string SetOwnRegexSearchExpression { get; set; }

        /// <summary>
        /// Path to the LibreOffice folder. Non-mandatory.
        /// </summary>
        public string SetPdfLibraryFolderPath { get; set; }

        /// <summary>
        /// Path to the plugin folder. Non-mandatory.
        /// </summary>
        public string SetPluginFolderPath { get; set; }

        /// <summary>
        /// Path to the temporary folder which is need to use .docx to pdf conversion. Non-mandatory.
        /// </summary>
        public string SetTempFolderPath { get; set; }

        /// <summary>
        /// Template in the form of string path, MemoryStream or byte array. Mandatory.
        /// </summary>
        public dynamic SetTemplatePathOrBytes { get; set; }
    }
}
