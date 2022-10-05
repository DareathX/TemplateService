using DocumentFormat.OpenXml.Packaging;
using System.Text.RegularExpressions;
using TemplateService;
using OpenXmlPowerTools;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Xml.Linq;

namespace FillInWordTemplate
{
    public class WordTemplate : IFillInTemplates
    {
        public string[] supportedFileExtensions => new string[] { ".docx", ".docm" };
        private Regex _regex = new Regex(@"\{\{[a-zA-Z0-9_&^. -]+\}\}", RegexOptions.IgnoreCase);
        private List<(object, object)> mergeTemplates = new List<(object, object)>();

        /// <summary>
        /// Compares found placeholders with placeholder dictionary and replacees them in the template.
        /// </summary>
        /// <param name="wordDoc">The provided Word template.</param>
        /// <param name="placeholders">The provided dictionary with data for the placeholders.</param>
        private void FillInTemplate(WordprocessingDocument wordDoc, Dictionary<string, object> placeholders, bool emptyOnNotFound, MemoryStream stream)
        {
            char[] charsToTrim = { '{', '}' };
            List<(HashSet<string>, object)> placeholderLocation;

            foreach ((HashSet<string>, object) match in getPlaceholders(wordDoc))
            {
                foreach (string foundPlaceholder in match.Item1)
                {
                    string trimmedMatch = foundPlaceholder.Trim(charsToTrim);
                    object? placeholderData = placeholders.
                        Where(f => trimmedMatch.
                        Equals(f.Key, StringComparison.OrdinalIgnoreCase)).
                        Select(f => f.Value).
                        FirstOrDefault();//Checks placeholders with placeholder out of dictionary.

                    IsStringEmpty.SetStringEmptyOrThrowException(ref placeholderData, emptyOnNotFound, $"Dictionary does not contain a value for {trimmedMatch}");//Checks if placeholder is not null or empty.

                    if (match.Item2.GetType() == typeof(HeaderPart))//Needs to be unloaded or else TextReplacer.SearchAndReplace won't execute properly.
                    {
                        (match.Item2 as HeaderPart).UnloadRootElement();
                    } 
                    else if (match.Item2 is FooterPart)
                    {
                        (match.Item2 as FooterPart).UnloadRootElement();
                    }

                    if (placeholderData.GetType() == typeof(string) && !File.Exists(placeholderData as string))//Replace placeholder if it is not a path.
                    {
                        TextReplacer.SearchAndReplace(wordDoc, foundPlaceholder, placeholderData as string, true);
                        continue;
                    }

                    if (IsImage.HasImageExtension(placeholderData as string) || IsImage.HasImageExtension(placeholderData as MemoryStream))//Replace placeholder with an inline image.
                    {
                        InsertImage.InsertAnImage(wordDoc, placeholderData, foundPlaceholder);
                        continue;
                    }
                    //Replace placeholder with another template if the match location is footer or header.
                    if (match.Item2 != "body")
                    {
                        mergeTemplates.Add((placeholderData, match.Item2));
                        continue;
                    }

                    //Inserts new attribute to the location where the template should be inserted.
                    Paragraph paragraph = wordDoc.MainDocumentPart.Document.Descendants<Paragraph>().FirstOrDefault(item => item.InnerText.Contains(foundPlaceholder));
                    int idOfParagraph = wordDoc.MainDocumentPart.Document.Descendants<Paragraph>().ToList().IndexOf(paragraph);
                    wordDoc.MainDocumentPart.UnloadRootElement();

                    var xDoc = wordDoc.MainDocumentPart.GetXDocument();
                    var paragraphToEdit = xDoc.Root.Descendants(W.p).ElementAt(idOfParagraph);
                    paragraphToEdit.ReplaceWith(new XElement(PtOpenXml.Insert, new XAttribute("Id", "Placeholder")));
                    wordDoc.MainDocumentPart.PutXDocument();

                    mergeTemplates.Add((placeholderData, match.Item2));
                }
            }
        }

        /// <summary>
        /// Gets placeholders by looking in the headers, footers and in the main body of the document with Regex.
        /// </summary>
        /// <param name="wordDoc">The provided template.</param>
        /// <returns>Returns List of placeholders found seperated by location where they were found.</returns>
        private List<(HashSet<string>, object)> getPlaceholders(WordprocessingDocument wordDoc)
        {
            List<(HashSet<string>, object)> matches = new List<(HashSet<string>, object)>()
            {
                (_regex.Matches(PtOpenXmlExtensions.GetXDocument(wordDoc.MainDocumentPart).Root.Value)
                .OfType<Match>()
                .Select(m => m.Groups[0].Value)
                .Distinct()
                .ToHashSet(), "body")
            };

            HashSet<string> headerMatches = new HashSet<string>();
            HashSet<string> footerMatches = new HashSet<string>();

            foreach (HeaderPart headerPart in wordDoc.MainDocumentPart!.HeaderParts)
            {
                HashSet<string> headerPlaceholders = _regex.Matches(headerPart.RootElement.InnerText)
                .OfType<Match>()
                .Select(m => m.Groups[0].Value)
                .Distinct()
                .ToHashSet();

                matches.Add((headerPlaceholders, headerPart));
            }
            foreach (FooterPart footerPart in wordDoc.MainDocumentPart!.FooterParts)
            {
                HashSet<string> footerPlaceholders = _regex.Matches(footerPart.RootElement.InnerText)
                .OfType<Match>()
                .Select(m => m.Groups[0].Value)
                .Distinct()
                .ToHashSet();

                matches.Add((footerPlaceholders, footerPart));
            }
            
            return matches;
        }

        /// <summary>
        /// Get MemoryStream from filled in template. Removes macro's from Word template to make it compatible.
        /// </summary>
        /// <param name="fileInByteArray">Template in a byte array.</param>
        /// <param name="placeholders">Dictionary of all the placeholders.</param>
        /// <param name="emptyOnNotFound">Throws error when placeholder is not found in given placeholders or removes the placeholder from docuemnt.</param>
        /// <param name="regexExpression">Set own regex expression</param>
        /// <param name="creatorName">Set name of creator of file</param>
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

                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(stream, true))
                {
                    DeleteVbaProjectPart(wordDoc);
                    FillInTemplate(wordDoc, placeholders, emptyOnNotFound, stream);
                    SetCreatorName(wordDoc, creatorName);
                    
                    wordDoc.Save();
                }

                (object, object) templateToBeInserted = mergeTemplates.FirstOrDefault();
                mergeTemplates.Remove(templateToBeInserted);

                byte[] mergedTemplates = MergeTemplates(stream, templateToBeInserted, placeholders, emptyOnNotFound);
                
                if(mergedTemplates != null)
                {
                    new MemoryStream(mergedTemplates).CopyTo(memStr);
                }
                else
                {
                    stream.Seek(0, SeekOrigin.Begin);
                    stream.CopyTo(memStr);//Writes filled in template bytes to a new MemoryStream so it won't be disposed.
                }
            }

            return memStr;
        }
        /// <summary>
        /// Removes macro's from the document to be able to convert it from .docm to a .docx document.
        /// </summary>
        /// <param name="doc"></param>
        private void DeleteVbaProjectPart(WordprocessingDocument doc)
        {
            VbaProjectPart vbaPart = doc.MainDocumentPart.VbaProjectPart;
            if (vbaPart != null)
            {
                doc.MainDocumentPart.DeletePart(vbaPart);
                
                doc.ChangeDocumentType(
                    WordprocessingDocumentType.Document);
            }
        }

        /// <summary>
        /// Set the name of the creator/editor of the document.
        /// </summary>
        /// <param name="wordDoc">Word template.</param>
        /// <param name="creatorName">Name of creator.</param>
        private void SetCreatorName(WordprocessingDocument wordDoc, string creatorName)
        {
            if (creatorName != string.Empty)
            {
                wordDoc.PackageProperties.Creator = creatorName;
                wordDoc.PackageProperties.LastModifiedBy = creatorName;
            }
        }

        /// <summary>
        /// Inserts template and fills it in.  
        /// </summary>
        /// <param name="mainWordDoc">Template where other template will be inserted into.</param>
        /// <param name="template">Template to be inserted.</param>
        /// <param name="placeholders">The provided dictionary with data for the placeholders.</param>
        /// <param name="emptyOnNotFound">Throws error when placeholder is not found in given placeholders or removes the placeholder from docuemnt.</param>
        /// <returns>Returns byte array with a Word document of the merged templates</returns>
        private byte[] MergeTemplates(MemoryStream mainWordDoc, (object, object) template, Dictionary<string, object> placeholders, bool emptyOnNotFound)
        {
            byte[] templateArray = null;
            switch (template.Item1)//Template stream or byte[]
            {
                case MemoryStream stream:
                    templateArray = stream.ToArray();
                    break;
                case string path:
                    templateArray = File.ReadAllBytes(path);
                    break;
                default:
                    return null;
            }
            templateArray = GetMemoryStream(templateArray.ToArray(), placeholders, emptyOnNotFound, "", "").ToArray();

            switch (template.Item2)//Placeholder part location
            {
                case "body":
                    return MergeBody(mainWordDoc.ToArray(), templateArray);
                case HeaderPart:
                case FooterPart:
                    return MergeHeadersAndFooters(mainWordDoc.ToArray(), templateArray);
                default:
                    return null;
            }
        }

        /// <summary>
        /// Inserts template into the body of the template.
        /// </summary>
        /// <param name="mainWordDoc">Template where other template will be inserted into.</param>
        /// <param name="newBody">Body of the to be inserted template.</param>
        /// <returns>Returns byte array with a Word document of the merged templates</returns>
        private byte[] MergeBody(byte[] mainWordDoc, byte[] newBody)
        {
            List<Source> sources = new List<Source>()
            {
                new Source(new WmlDocument(new OpenXmlPowerToolsDocument(mainWordDoc)), true),
                new Source(new WmlDocument(new OpenXmlPowerToolsDocument(newBody)), "Placeholder")
            };

            return DocumentBuilder.BuildDocument(sources).DocumentByteArray;
        }

        /// <summary>
        /// Inserts template headers and footers into the headers and footers of the template.
        /// </summary>
        /// <param name="mainWordDoc">Template where other template will be inserted into.</param>
        /// <param name="newHeaderOrFooterDoc">Header or Footer of the to be inserted template.</param>
        /// <returns>Returns byte array with a Word document of the merged templates</returns>
        private byte[] MergeHeadersAndFooters(byte[] mainWordDoc, byte[] newHeaderOrFooterDoc)
        {
            List<Source> sources = new List<Source>()
            {
                new Source(new WmlDocument(new OpenXmlPowerToolsDocument(mainWordDoc)), false),
                new Source(new WmlDocument(new OpenXmlPowerToolsDocument(newHeaderOrFooterDoc)), true)
            };

            return DocumentBuilder.BuildDocument(sources).DocumentByteArray;
        }
    }
}