using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.IO;

namespace TemplateServiceTests
{
    [TestClass]
    public class ConvertToPdfPluginTests
    {
        [TestMethod]
        public void CheckIfByteArrayGetsChanged()//Cannot look into or convert pdf to something else
        {
            //Arrange
            ConvertWordTemplateToPdf.ConvertWordTemplateToPdf template = new ConvertWordTemplateToPdf.ConvertWordTemplateToPdf();
            string templatePath = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent + "/FakeFiles/FakeWordFillIn.docx";
            string libreOfficePath = "";
            string tempFolderPath = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent + "/FakeFiles/";
            Dictionary<string, string> dict = new Dictionary<string, string>() { { "pdfLibraryFolderPath", libreOfficePath },
                { "pluginFolderPath", tempFolderPath  },
                { "tempFolderPath", tempFolderPath + "pdf/" } };
            byte[] input = File.ReadAllBytes(templatePath);
            MemoryStream ms = new MemoryStream(input);
            
            //Act
            byte[] output = template.Convert(ms, dict);

            //Assert
            Assert.IsTrue(input.Length != output.Length && output.Length != 0);
        }
    }
}
