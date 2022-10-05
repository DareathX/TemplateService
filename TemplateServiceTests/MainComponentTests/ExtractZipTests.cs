using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using TemplateService;

namespace TemplateServiceTests.MainComponentTests
{
    [TestClass]
    public class ExtractZipTests
    {
        [TestMethod]
        public void CheckIfExpectedFilesAreFoundAfterExtraction()
        {
            //Arrange
            string inputPath = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent + "/FakeFiles/TestZip.zip";
            string outputPath = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent + "/FakeFiles/Test";

            //Act
            if (File.Exists(outputPath + "/test.txt")) Directory.Delete(outputPath, true);

            new ExtractZip(inputPath, outputPath);

            //Assert
            Assert.IsTrue(File.Exists(outputPath + "/test.txt"));
        }
    }
}
