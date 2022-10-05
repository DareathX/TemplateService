using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using TemplateService;

namespace TemplateServiceTests.MainComponentTests
{
    [TestClass]
    public class IsImageTests
    {
        [TestMethod]
        public void CheckIfItReturnsTrueIfJpegOrPngIsBeingGiven()
        {
            //Arrange
            string pngPath = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent + "/FakeFiles/PngTest.png";
            string jpegPath = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent + "/FakeFiles/JpegTest.jpg";

            //Act
            bool pathTest = IsImage.HasImageExtension(pngPath);
            string byteArrayTestExtension;
            bool byteArrayTest = IsImage.HasImageExtension(File.ReadAllBytes(jpegPath), out byteArrayTestExtension);
            bool memoryStreamTest = IsImage.HasImageExtension(new MemoryStream(File.ReadAllBytes(pngPath)));

            //Assert
            Assert.IsTrue(pathTest);
            Assert.IsTrue(byteArrayTest && (byteArrayTestExtension.Equals(".jpg") || byteArrayTestExtension.Equals(".jpeg")));
            Assert.IsTrue(memoryStreamTest);
        }

        [TestMethod]
        public void CheckIfItReturnsFalseIfNullsAreGiven()
        {
            //Act
            bool pathTest = IsImage.HasImageExtension("");
            string byteArrayTestExtension;
            bool byteArrayTest = IsImage.HasImageExtension(null, out byteArrayTestExtension);
            bool memoryStreamTest = IsImage.HasImageExtension((MemoryStream)null);

            //Assert
            Assert.IsFalse(pathTest);
            Assert.IsFalse(byteArrayTest);
            Assert.IsFalse(memoryStreamTest);
        }
    }
}
