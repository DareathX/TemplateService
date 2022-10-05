using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using System.Linq;
using TemplateService;

namespace TemplateServiceTests
{
    [TestClass]
    public class TempStorageTests
    {
        [TestMethod]
        public void CheckIfFolderGetsCreated()
        {
            //Arrange
            TempStorage tempStorage = new TempStorage();
            string folderPath = Directory.GetCurrentDirectory() + "/temp/";

            //Act
            tempStorage.CreateTempFolder(folderPath);

            //Assert
            Assert.IsTrue(Directory.Exists(folderPath));
        }

        [TestMethod]
        public void CheckIfFileNamesAreReturned()
        {
            //Arrange
            TempStorage tempStorage = new TempStorage();
            string folderPath = Directory.GetCurrentDirectory() + "/temp/";
            string filePath = "";

            //Act
            tempStorage.CreateTempFolder(folderPath);
            filePath = folderPath + "/Test.txt";
            File.Create(filePath);

            //Assert
            Assert.IsTrue("Test" == tempStorage.GetFolderFileNames(folderPath)[0]);
        }

        [TestMethod]
        public void CheckIfFilesAreBeingRemoved()
        {
            //Arrange
            TempStorage tempStorage = new TempStorage();
            string folderPath = Directory.GetCurrentDirectory() + "/temp/";
            string filePath = "";

            //Act
            tempStorage.CreateTempFolder(folderPath);
            filePath = folderPath + "/Test2.txt";
            File.Create(filePath).Close();
            tempStorage.DeleteFile(filePath);

            //Assert
            Assert.IsFalse(tempStorage.GetFolderFileNames(folderPath).Contains("Test2.txt"));
        }
    }
}
