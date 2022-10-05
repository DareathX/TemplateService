using Microsoft.VisualStudio.TestTools.UnitTesting;
using FillInHtmlTemplate;
using System.IO;
using System.Collections.Generic;

namespace TemplateServiceTests
{
    [TestClass]
    public class FillInHtmlPluginTests
    {
        [TestMethod]
        public void CheckIfPlaceholderAreBeingReplacedInHtml()
        {
            //Arrange
            HtmlTemplate template = new HtmlTemplate();
            string path = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent + "/FakeFiles/FakeHtmlFillIn.txt";
            Dictionary<string, object> dict = new Dictionary<string, object>() { { "content", "Filled" } };
            MemoryStream ms;
            StreamReader reader;
            string txt;

            //Act
            ms = template.GetMemoryStream(File.ReadAllBytes(path), dict, false, "", "");
            ms.Seek(0, SeekOrigin.Begin);
            reader = new StreamReader(ms);
            txt = reader.ReadToEnd();


            //Assert
            Assert.IsTrue(txt.Contains("Filled") && !txt.Contains("{{") && !txt.Contains("}}"));
        }
    }
}
