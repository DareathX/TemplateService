using System.Text.RegularExpressions;
using TemplateService;

namespace FillInHtmlTemplate
{
    public class HtmlTemplate : IFillInTemplates
    {
        public string[] supportedFileExtensions => new string[] { ".txt", ".html" };
        private Regex _regex = new Regex(@"\{\{[a-zA-Z0-9_&^. -]+\}\}", RegexOptions.IgnoreCase);

        /// <summary>
        /// Get MemoryStream from filled in template.
        /// </summary>
        /// <param name="fileInByteArray">Template in a byte array.</param>
        /// <param name="placeholders">Dictionary of all the placeholders.</param>
        /// <param name="emptyOnNotFound">Throws error when placeholder is not found in given placeholders or removes the placeholder from docuemnt.</param>
        /// <param name="regexExpression">Set own regex expression</param>
        /// <param name="creatorName">Set name of creator of file. Cannot be set for Html.</param>
        /// <returns>Returns a filled in template in a MemoryStream.</returns>
        public MemoryStream GetMemoryStream(byte[] fileInByteArray, Dictionary<string, object> placeholders, bool emptyOnNotFound, string regexExpression, string creatorName)
        {
            MemoryStream memStr = new MemoryStream();

            if (regexExpression != string.Empty)
            {
                _regex = new Regex(regexExpression);
            }

            using (MemoryStream stream = new MemoryStream(fileInByteArray.Length))//Reads template in memory.
            {
                stream.Write(fileInByteArray, 0, fileInByteArray.Length);
                stream.Seek(0, SeekOrigin.Begin);

                using (StreamWriter write = new StreamWriter(stream))
                {
                    using (StreamReader read = new StreamReader(stream))
                    {
                        string content = read.ReadToEnd();
                        stream.Seek(0, SeekOrigin.Begin);
                        write.Write(FillInTemplate(content, placeholders, emptyOnNotFound));
                        write.Flush();

                        stream.Seek(0, SeekOrigin.Begin);
                        stream.CopyTo(memStr);//Writes filled in template bytes to a new MemoryStream so it won't get disposed.
                    }
                }
            }

            return memStr;
        }

        /// <summary>
        /// Fills in template.
        /// </summary>
        /// <param name="content">The whole text found in a Html template</param>
        /// <param name="placeholders">Dictionary of all placeholders.</param>
        /// <param name="emptyOnNotFound">Throws error when placeholder is not found in given placeholders or removes the placeholder from docuemnt.</param>
        /// <returns></returns>
        private string FillInTemplate(string content, Dictionary<string, object> placeholders, bool emptyOnNotFound)
        {
            char[] charsToTrim = { '{', '}' };

            foreach (string match in getPlaceholders(content))
            {
                string trimmedMatch = match.Trim(charsToTrim);
                object? placeholderData = placeholders.
                    Where(f => trimmedMatch.
                    Contains(f.Key, StringComparison.OrdinalIgnoreCase)).
                    Select(f => f.Value).
                    FirstOrDefault();//Checks placeholders with placeholder out of dictionary.

                IsStringEmpty.SetStringEmptyOrThrowException(ref placeholderData, emptyOnNotFound, $"Dictionary does not contain a value for {trimmedMatch}");//Checks if placeholder is not null or empty.

                if (placeholderData.GetType() == typeof(string) && !File.Exists(placeholderData as string))//Replace placeholder if it is not a path.
                {
                    placeholderData = Regex.Escape(placeholderData as string).Contains("\\") ? "<div style=\"white-space: pre-wrap\">" + placeholderData + "</div>" : placeholderData;
                    content = content.Replace(match, Regex.Unescape(placeholderData as string));
                    continue;
                }

                if (IsImage.HasImageExtension(placeholderData as string) || IsImage.HasImageExtension(placeholderData as MemoryStream))//Replace placeholder with an inline image.
                {
                    content = content.Replace(match, '"' + "data:image/png;base64," + Convert.ToBase64String((placeholderData as MemoryStream).ToArray()) + '"');
                    continue;
                }

                content = content.Replace(match, TemplateInTemplate(placeholderData, placeholders, emptyOnNotFound));//Replace placeholder with another template.

            }
            return content;
        }

        /// <summary>
        /// Get placeholders found in template.
        /// </summary>
        /// <param name="content">Whole template text.</param>
        /// <returns>Returns a HashSet with all placeholders found.</returns>
        private HashSet<string> getPlaceholders(string content)
        {
            HashSet<string> matches = _regex.Matches(content)
                .OfType<Match>()
                .Select(m => m.Groups[0].Value)
                .Distinct()
                .ToHashSet();

            return matches;
        }

        /// <summary>
        /// Merges multiple template with each other. 
        /// </summary>
        /// <param name="template">Template in MemoryStream or by path.</param>
        /// <param name="placeholders">Dictionary of all the placeholders.</param>
        /// <param name="emptyOnNotFound">Throws error when placeholder is not found in given placeholders or removes the placeholder from docuemnt.</param>
        /// <returns>Returns the whole template with filled in template.</returns>
        private string TemplateInTemplate(object template, Dictionary<string, object> placeholders, bool emptyOnNotFound)
        {
            switch (template)
            {
                case MemoryStream stream:
                    return MergeTemplatesWithStream(stream, placeholders, emptyOnNotFound);
                case string path:
                    return MergeTemplatesWithPath(path, placeholders, emptyOnNotFound);
                default:
                    return "";
                    break;
            }
        }

        /// <summary>
        /// Insert template in another template. It creates a <div></div> with a style so that the format of the template remains.
        /// </summary>
        /// <param name="templateStream">MemoryStream of the template which will be inserted</param>
        /// <param name="placeholders">Dictionary of all the placeholders.</param>
        /// <param name="emptyOnNotFound">Throws error when placeholder is not found in given placeholders or removes the placeholder from docuemnt.</param>
        /// <returns></returns>
        private string MergeTemplatesWithStream(MemoryStream templateStream, Dictionary<string, object> placeholders, bool emptyOnNotFound)
        {
            string content = "";
            using (StreamReader reader = new StreamReader(templateStream))
            {
                content = "<div style=\"white-space: pre-wrap\">" + Regex.Unescape(reader.ReadToEnd()) + "</div>";
            }

            return FillInTemplate(content, placeholders, emptyOnNotFound);
        }

        /// <summary>
        /// Inserts template in another template.
        /// </summary>
        /// <param name="templatePath">Path to the template which will be inserted.</param>
        /// <param name="placeholders">Dictionary of all the placeholders.</param>
        /// <param name="emptyOnNotFound">Throws error when placeholder is not found in given placeholders or removes the placeholder from docuemnt.</param>
        /// <returns></returns>
        private string MergeTemplatesWithPath(string templatePath, Dictionary<string, object> placeholders, bool emptyOnNotFound)
        {
            return FillInTemplate(File.ReadAllText(templatePath), placeholders, emptyOnNotFound);
        }
    }
}