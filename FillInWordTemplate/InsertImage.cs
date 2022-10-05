using TemplateService;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace FillInWordTemplate
{
    internal class InsertImage
    {
        /// <summary>
        /// Replace placeholder with an image.
        /// </summary>
        /// <param name="wordDoc">Word template.</param>
        /// <param name="image">Image to be inserted.</param>
        /// <param name="placeholder">Placeholder to be replaced by image.</param>
        public static void InsertAnImage(WordprocessingDocument wordDoc, object image, string placeholder)
        {
            switch (image)
            {
                case string:
                    InsertImagePart(wordDoc, new MemoryStream(File.ReadAllBytes(image as string)), placeholder);
                    break;
                case MemoryStream:
                    InsertImagePart(wordDoc, image as MemoryStream, placeholder);
                    break;
                default:
                    break;
            }
        }

        /// <summary>
        /// Adds ImagePart to the Word template and feeds the ImagePart with the image.
        /// </summary>
        /// <param name="wordDoc">Word template.</param>
        /// <param name="image">Image to be inserted.</param>
        /// <param name="placeholder">Placeholder to be replaced by image.</param>
        private static void InsertImagePart(WordprocessingDocument wordDoc, MemoryStream image, string placeholder)
        {
            string extension;
            if (IsImage.HasImageExtension(image.ToArray(), out extension))
            {
                MainDocumentPart mainPart = wordDoc.MainDocumentPart;
                ImagePart imagePart = null;

                switch (extension)
                {
                    case ".jpg":
                    case ".jpeg":
                        imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);
                        break;
                    case ".png":
                        imagePart = mainPart.AddImagePart(ImagePartType.Png);
                        break;
                    default:
                        return;
                }

                imagePart.FeedData(image);

                InsertImageToDoc(wordDoc, mainPart.GetIdOfPart(imagePart), placeholder);
            }
        }

        /// <summary>
        /// Insert image in the headers and footers if placeholder is found.
        /// </summary>
        /// <param name="wordDoc">Word template.</param>
        /// <param name="relationshipId">Id of ImagePart.</param>
        /// <param name="placeholder">Placeholder to be replaced by image.</param>
        private static void InsertImageToDoc(WordprocessingDocument wordDoc, string relationshipId, string placeholder)
        {
            foreach (var header in wordDoc.MainDocumentPart.HeaderParts)//Searches through headers.
            {
                InsertImageToLayout(wordDoc, header.Header.Descendants<Text>(), relationshipId, placeholder);
            }

            InsertImageToLayout(wordDoc, wordDoc.MainDocumentPart.Document.Body.Descendants<Text>(), relationshipId, placeholder);

            foreach (var footer in wordDoc.MainDocumentPart.FooterParts)//Searches through footers.
            {
                InsertImageToLayout(wordDoc, footer.Footer.Descendants<Text>(), relationshipId, placeholder);
            }
        }

        /// <summary>
        /// Find insert location by going through every character of the docment.
        /// </summary>
        /// <param name="wordDoc">Word template.</param>
        /// <param name="docText">All text found in Word tempalte.</param>
        /// <param name="relationshipId">Id of ImagePart.</param>
        /// <param name="placeholder">Placeholder to be replaced by image.</param>
        private static void InsertImageToLayout(WordprocessingDocument wordDoc, IEnumerable<Text> docText, string relationshipId, string placeholder)
        {
            List<Text> docTextList = docText.ToList();
            List<int> textIndexes = new List<int>();
            List<int> charIndexes = new List<int>();
            int index = 0;

            for (int i = 0; i < docTextList.Count; i++)
            {
                for (int j = 0; j < docTextList[i].Text.Length; j++)//Go through every character in doc.
                {
                    //If correct character found, save it and move to next. If wrong one is found after a correct one, process gets deleted and it moves on. 
                    if (docTextList[i].Text[j] == placeholder[index])
                    {
                        charIndexes.Add(j);
                        textIndexes.Add(docTextList.IndexOf(docTextList[i]));
                        index++;
                        if (textIndexes.Count == placeholder.Length)//Match has been found.
                        {
                            textIndexes = textIndexes.Distinct().ToList();
                            Insert(ref docTextList, docText, textIndexes, relationshipId, docTextList[i].Text.Substring(j).Remove(0, 1), charIndexes, placeholder);

                            charIndexes.Clear();
                            textIndexes.Clear();
                            index = 0;
                            i = 0;
                            j = 0;
                        }
                    }
                    else
                    {
                        charIndexes.Clear();
                        textIndexes.Clear();
                        index = 0;
                        continue;
                    }
                }
            }
        }

        /// <summary>
        /// Inserts image based on found indexes.
        /// </summary>
        /// <param name="docText">All text found in Word template.</param>
        /// <param name="indexes">List of all placeholder character indexes.</param>
        /// <param name="relationshipId">Id of ImagePart.</param>
        private static void Insert(ref List<Text> docTextList, IEnumerable<Text> docText, List<int> indexes, string relationshipId, string leftOver, List<int> charIndexes, string placeholder)
        {
            for (int i = 0; i < indexes.Count; i++)
            {
                var found = docText.ElementAt(indexes[i]);

                if (indexes.Count > 1 && found.Text.Length >= charIndexes[i] && i == 0)
                {
                    found.Text = found.Text.Remove(charIndexes.First());//Deletes all characters from first element if multiple Text elements are found
                    continue;
                }
                else if (i != 0 && i != indexes.Count - 1)
                {
                    found.Parent.RemoveChild(found);//Delete all characters between the first and last Text elements.
                    indexes.RemoveAt(indexes.Count - 1);
                    i = 0;
                    continue;
                }
                else if (indexes.Count > 1 && i == indexes.Count - 1)
                {
                    found.Text = found.Text.Remove(0, charIndexes.Last() + 1);//Removes first few characters of the last element if multple Text elements are found
                }
                else
                {
                    found.Text = found.Text.Replace(placeholder, ""); //Removes placeholder.
                }
                
                //Replaces placeholder location with image and the puts text behind the placeholder with a new element.
                found.Parent.Append(NewImageElement(relationshipId));

                if (leftOver.Length > 1)
                {
                    found.Text = found.Text.Replace(leftOver, "");

                    Text second = found.Clone() as Text;
                    second.Text = leftOver;

                    found.Parent.Append(second);
                    docTextList = docText.ToList();
                }
            }
        }

        /// <summary>
        /// Creates a new image element for Word template.
        /// </summary>
        /// <param name="relationshipId">Id of ImagePart.</param>
        /// <returns></returns>
        private static Drawing NewImageElement(string relationshipId)
        {
            return new Drawing(
             new DW.Inline(
                 new DW.Extent() { Cx = 990000L, Cy = 792000L },
                 new DW.EffectExtent()
                 {
                     LeftEdge = 0L,
                     TopEdge = 0L,
                     RightEdge = 0L,
                     BottomEdge = 0L
                 },
                 new DW.DocProperties()
                 {
                     Id = (UInt32Value)1U,
                     Name = "Picture 1"
                 },
                 new DW.NonVisualGraphicFrameDrawingProperties(
                     new A.GraphicFrameLocks() { NoChangeAspect = true }),
                 new A.Graphic(
                     new A.GraphicData(
                         new PIC.Picture(
                             new PIC.NonVisualPictureProperties(
                                 new PIC.NonVisualDrawingProperties()
                                 {
                                     Id = (UInt32Value)0U,
                                     Name = "New Bitmap Image.jpg"
                                 },
                                 new PIC.NonVisualPictureDrawingProperties()),
                             new PIC.BlipFill(
                                 new A.Blip(
                                     new A.BlipExtensionList(
                                         new A.BlipExtension()
                                         {
                                             Uri =
                                               "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                         })
                                 )
                                 {
                                     Embed = relationshipId,
                                     CompressionState =
                                     A.BlipCompressionValues.Print
                                 },
                                 new A.Stretch(
                                     new A.FillRectangle())),
                             new PIC.ShapeProperties(
                                 new A.Transform2D(
                                     new A.Offset() { X = 0L, Y = 0L },
                                     new A.Extents() { Cx = 990000L, Cy = 792000L }),
                                 new A.PresetGeometry(
                                     new A.AdjustValueList()
                                 )
                                 { Preset = A.ShapeTypeValues.Rectangle }))
                     )
                     { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
             )
             {
                 DistanceFromTop = (UInt32Value)0U,
                 DistanceFromBottom = (UInt32Value)0U,
                 DistanceFromLeft = (UInt32Value)0U,
                 DistanceFromRight = (UInt32Value)0U,
                 EditId = "50D07946"
             });
        }
    }
}
