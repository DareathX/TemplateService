namespace TemplateService
{
    public class ProcessTemplateRequest
    {
        private Dictionary<string, object>? _inputDataDict;
        private dynamic? _setTemplatePathOrBytes;
        private (string, string) _setConversionFromTo;
        private string? _setOutputDocumentPath;
        private string? _setPluginFolderPath;
        private List<dynamic> _allPlugins = new List<dynamic>();
        private Dictionary<string, string> _additionalData = new Dictionary<string, string>();

        /// <summary>
        /// Inputted template will be filled in with the provided data and outputted as requested document.
        /// Use no overload if you only want to generate the document on a specific path.
        /// </summary>
        /// <param name="data">Struct with additional information needed to process template.</param>
        public ProcessTemplateRequest(TemplateServiceInputData data)
        {
            ProcessRequest(data);
        }
        /// <summary>
        /// Inputted template will be filled in with the provided data and outputted as requested document as a byte array.
        /// </summary>
        /// <param name="data">Struct with additional information needed to process template.</param>
        /// <param name="output">Puts the generated document in a byte array.</param>
        public ProcessTemplateRequest(TemplateServiceInputData data, ref byte[] output)
        {           
            output = ProcessRequest(data);
        }
        /// <summary>
        /// Inputted template will be filled in with the provided data and outputted as requested document as a MemoryStream.
        /// </summary>
        /// <param name="data">Struct with additional information needed to process template.</param>
        /// <param name="output">Puts the generated document in a MemoryStream.</param>
        public ProcessTemplateRequest(TemplateServiceInputData data, ref MemoryStream output)
        {
            output = new MemoryStream(ProcessRequest(data));
        }

        /// <summary>
        /// Fills in template if ConvertOnly is turned off and convert to a desired document format.
        /// </summary>
        /// <param name="data">Struct with additional data needed to process template.</param>
        /// <returns>Returns document in a byte array.</returns>
        /// <exception cref="NullReferenceException">Struct data is missing or invalid.</exception>
        private byte[] ProcessRequest(TemplateServiceInputData data)
        {
            ValidateInput(data);

            bool convert = false;
            LoadPlugins(ref convert);
            convert = data.ConvertOnly ? data.ConvertOnly : convert; 

            MemoryStream memStr = null;
            bool conversionSamePlugin = false;
            if (!convert)
            {
                memStr = FillInTemplate(CheckForSupportedInputPlugins(ref conversionSamePlugin)
                               ?? throw new NullReferenceException("No plugin found that supports this type of template."), data.EmptyOnNotFound, data.SetOwnRegexSearchExpression, data.SetDocumentCreatorName);
            }
            else
            {
                memStr = GetMemoryStreamFromFile();
            }

            byte[] bytes = new byte[0];

            if (!conversionSamePlugin && _setConversionFromTo.Item1 != _setConversionFromTo.Item2)//Checks if provided document format and requested document format are the same.
            {
                bytes = GenerateFile(CheckForSupportedOutputPlugins()
                ?? throw new NullReferenceException("No plugin found that supports this type of output file."), memStr);
            }
            else
            {
                bytes = memStr.ToArray();
            }
            
            WriteDocToSpecificPath(bytes);
            return bytes;
        }

        /// <summary>
        /// Looks in the plugin if the document is supported by the plugin.
        /// </summary>
        /// <returns>Returns a supported plugin or null.</returns>
        private object? CheckForSupportedInputPlugins(ref bool conversionSamePlugin)
        {
            foreach (var plugin in _allPlugins.OfType<IFillInTemplates>())
            {
                if (CheckIfSupported(plugin, GetFileExtension()))
                {
                    conversionSamePlugin = CheckIfSupported(plugin, _setConversionFromTo.Item2);
                    return plugin;
                }
            }
            return null;
        }

        /// <summary>
        /// Looks in the plugin if the document is supported by the plugin.
        /// </summary>
        /// <returns>Returns a supported plugin or null.</returns>
        private object? CheckForSupportedOutputPlugins()
        {
            foreach (var plugin in _allPlugins.OfType<IConvertTemplateTo>())
            {
                if (CheckIfSupported(plugin, _setConversionFromTo.Item2)) return plugin;
            }
            return null;
        }

        /// <summary>
        /// Checks if the extension is within the supported file extensions of the plugin.
        /// </summary>
        /// <param name="plugin">Plugin with the interface IFillInTemplates or IConvertTemplateTo.</param>
        /// <param name="extension">The extension of the template or the wanted outputted document.</param>
        /// <returns></returns>
        private bool CheckIfSupported(dynamic plugin, string extension)
        {
            return ((IEnumerable<string>)plugin.supportedFileExtensions).Contains(extension, StringComparer.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Validates the input data. Required data is placeholder dictionary, template, conversion from and to.
        /// </summary>
        /// <param name="data">Struct with the all the input data</param>
        /// <exception cref="NullReferenceException"></exception>
        private void ValidateInput(TemplateServiceInputData data)
        {
            try
            {
                if (!data.ConvertOnly)
                {
                    _inputDataDict = data.InputDataDict
                    ?? throw new Exception(nameof(data.InputDataDict));
                }

                _setTemplatePathOrBytes = data.SetTemplatePathOrBytes
                    ?? throw new Exception(nameof(data.SetTemplatePathOrBytes));
            }
            catch (Exception e)
            {

                throw new Exception("Missing data " + e.Message);
            }

            _setConversionFromTo = data.SetConversionFromTo
                    ?? (".docx", ".docx");
            _setOutputDocumentPath = data.SetOutputDocumentPath;
            _setPluginFolderPath = data.SetPluginFolderPath ?? Directory.GetCurrentDirectory() + "/TemplatePlugins";

            //Adds additional folder path data.
            _additionalData.Add("tempFolderPath", data.SetTempFolderPath != null ? data.SetTempFolderPath : "");
            _additionalData.Add("pdfLibraryFolderPath", data.SetPdfLibraryFolderPath != null ? data.SetPdfLibraryFolderPath : "");
            _additionalData.Add("pluginFolderPath", _setPluginFolderPath);
        }

        /// <summary>
        /// Loads in runtime all dll's in the plugin folder.
        /// </summary>
        /// <param name="convertOnly">Sets to true if FillInTemplate plugins are not found.</param>
        private void LoadPlugins(ref bool convertOnly)
        {
            PluginLoader pluginLoader = new PluginLoader();
            convertOnly = pluginLoader.LoadAllPlugins(ref _allPlugins, Directory.GetFiles(_setPluginFolderPath, "*.dll"));
        }

        /// <summary>
        /// Get extension from file.
        /// </summary>
        /// <returns>Returns found or expected extension.</returns>
        private string GetFileExtension()
        {
            if (_setTemplatePathOrBytes?.GetType() == typeof(string))
            {
                return Path.GetExtension(_setTemplatePathOrBytes);
            }

            return _setConversionFromTo.Item1;
        }

        /// <summary>
        /// Copies template and fills it in which will be returned as a MemoryStream.
        /// </summary>
        /// <param name="plugin"></param>
        /// <returns></returns>
        private MemoryStream FillInTemplate(dynamic plugin, bool emptyOnNotFound, string regexExpression, string creatorName)
        {
            return plugin.GetMemoryStream(GetByteArrayFromFile(), _inputDataDict, emptyOnNotFound, regexExpression ?? string.Empty, creatorName );
        }

        /// <summary>
        /// Converts the document to a byte array.
        /// </summary>
        /// <returns>A byte array of the document.</returns>
        /// <exception cref="Exception"></exception>
        private byte[] GetByteArrayFromFile()
        {
            switch (_setTemplatePathOrBytes)
            {
                case string:
                    if (!File.Exists(_setTemplatePathOrBytes as string)) throw new Exception("The file on provided template path cannot be found.");

                    using (FileStream strStream = new FileStream(_setTemplatePathOrBytes, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) 
                    {
                        using (MemoryStream ms = new MemoryStream())
                        {
                            strStream.CopyTo(ms);
                            return ms.ToArray();
                        }
                    }
                case byte[]:
                    return _setTemplatePathOrBytes as byte[];
                case MemoryStream stream:
                    return stream.ToArray();
                default:
                    break;
            }
            throw new Exception("The provided template has not an accepted type. Accepted types are byte array, Memorystream and path.");
        }

        /// <summary>
        /// Converts the document to a memorystream.
        /// </summary>
        /// <returns>A memorystream of the document.</returns>
        /// <exception cref="Exception"></exception>
        private MemoryStream GetMemoryStreamFromFile()
        {
            switch (_setTemplatePathOrBytes)
            {
                case string:
                    if (!File.Exists(_setTemplatePathOrBytes as string)) throw new Exception("The file on provided template path cannot be found.");

                    return new MemoryStream(File.ReadAllBytes(_setTemplatePathOrBytes));
                case byte[]:
                    return new MemoryStream(_setTemplatePathOrBytes);
                case MemoryStream stream:
                    return _setTemplatePathOrBytes as MemoryStream;
                default:
                    break;
            }
            throw new Exception("The provided template has not an accepted type. Accepted types are byte array, Memorystream and path.");
        }

        /// <summary>
        /// Converts filled in template to the desired document format.
        /// </summary>
        /// <param name="plugin">The plugin which supports the desired document format.</param>
        /// <param name="memStr">The filled in template as a MemoryStream.</param>
        /// <param name="path">Optional destination path, for when you want to generate the document to a specific path.</param>
        /// <returns></returns>
        private byte[] GenerateFile(dynamic plugin, MemoryStream memStr)
        {

            return plugin.Convert(memStr, _additionalData);
        }

        /// <summary>
        /// Writes document to specific path. Document extension is not needed but can be used in the path.
        /// </summary>
        /// <param name="bytes">Document in a byte array.</param>
        private void WriteDocToSpecificPath(byte[] bytes)
        {
            if (_setOutputDocumentPath != null)
            {
                if (Path.HasExtension(_setOutputDocumentPath))
                {
                    File.WriteAllBytes(_setOutputDocumentPath, bytes);
                }
                else
                {
                    File.WriteAllBytes(_setOutputDocumentPath + _setConversionFromTo.Item2, bytes);
                }
            }
        }
    }
}
