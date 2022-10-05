using Microsoft.VisualStudio.TestTools.UnitTesting;
using FillInWordTemplate;
using System.IO;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;

namespace TemplateServiceTests
{
    [TestClass]
    public class FillInWordPluginTests
    {
        [TestMethod]
        public void CheckIfPlaceholderAreBeingReplacedInWord()
        {
            //Arrange
            WordTemplate template = new WordTemplate();
            string path = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent + "/FakeFiles/FakeWordFillIn.docx";
            Dictionary<string, object> dict = new Dictionary<string, object>() { { "Test_Word", "Filled" } };
            MemoryStream ms;
            StreamReader reader;
            string txt;

            //Act
            ms = template.GetMemoryStream(File.ReadAllBytes(path), dict, false, "", "");
            ms.Seek(0, SeekOrigin.Begin);
            
            using (ms)
            {
                using (WordprocessingDocument doc = WordprocessingDocument.Open(ms, true))
                {
                    txt = doc.MainDocumentPart.Document.InnerText;
                }
            }

            //Assert
            Assert.IsTrue(txt.Contains("Filled") && !txt.Contains("{{") && !txt.Contains("}}"));
        }
    }
}
