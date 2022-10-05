namespace TemplateService
{
    public static class IsImage
    {
        private static readonly List<byte> _jpeg = new List<byte> { 0xFF, 0xD8 };
        private static readonly List<byte> _png = new List<byte> { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };
        public static readonly List<(List<byte> imageFormat, string extension)> imageFormats = new List<(List<byte> imageFormat, string extension)>()
        {
            (_jpeg, ".jpeg"),
            (_jpeg, ".jpg"),
            (_png, ".png"),
        };

        /// <summary>
        /// Checks if file has an image extension.
        /// </summary>
        /// <param name="path">Path to the file to check.</param>
        /// <returns>Returns bool based on if the file is an image.</returns>
        public static bool HasImageExtension(string path)
        {
            if (path == null || path == string.Empty) return false;
            string foundExtension = Path.GetExtension(path);

            foreach (var file in imageFormats)
            {
               if (file.extension.Contains(foundExtension)) return true;
            }

            return false;
        }

        /// <summary>
        /// Checks if the file is an image.
        /// </summary>
        /// <param name="array">File in a MemoryStream.</param>
        /// <returns>Returns bool based on if the file is an image.</returns>
        public static bool HasImageExtension(MemoryStream array)
        {
            if (array == null) return false;
            string foundExtension;

            return CompareFirstFewbytes(array.ToArray(), out foundExtension);
        }

        /// <summary>
        /// Checks if the file is an image.
        /// </summary>
        /// <param name="array">File in a byte array</param>
        /// <param name="extension">Outputs the found extension.</param>
        /// <returns>Returns bool based on if the file is an image.</returns>
        public static bool HasImageExtension(byte[] array, out string extension)
        {
            if (array == null)
            {
                extension = "";
                return false;
            }

            return CompareFirstFewbytes(array, out extension);
        }

        /// <summary>
        /// Compares the first few bytes of the byte array to check if the file is an image.
        /// </summary>
        /// <param name="array">File in a byte array.</param>
        /// <param name="extension">Outputs the found extension.</param>
        /// <returns>Returns bool based on if the file is an image.</returns>
        private static bool CompareFirstFewbytes(byte[] array, out string extension)
        {
            bool found = false;
            string foundExtension = "";
            foreach (var comparer in imageFormats)
            {
                int arrayIndex = 0;
                foundExtension = comparer.extension;
                found = true;

                foreach (byte c in comparer.imageFormat)
                {
                    if (arrayIndex > array.Length - 1 || array[arrayIndex] != c)
                    {
                        found = false;
                        foundExtension = "";
                        break;
                    }
                    ++arrayIndex;
                }
                if (found)
                {
                    extension = foundExtension;
                    return found;
                }
            }

            extension = foundExtension;
            return found;
        }
    }
}
