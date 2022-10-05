using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using TemplateService;

namespace TemplateServiceTests
{
    [TestClass]
    public class DictionaryHelperTests
    {
        [TestMethod]
        public void CreateDictionaryWithObjectReturnsFilledDictionary()
        {
            //Arrange
            var test = new { Name = "Tester", Age = 20 };

            //Act
            var testDict = DictionaryHelper.CreateDictionary(test);

            //Assert
            Assert.IsTrue(testDict.ContainsKey("Name") && testDict.ContainsValue("20"));
        }

        [TestMethod]
        public void CreateDictionaryWithObjectAndDatetimeFormatReturnsFilledDictionary()
        {
            //Arrange
            var test = new { Time = DateTime.Today };

            //Act
            var testDict = DictionaryHelper.CreateDictionary(test, "dd MMM yyyy");

            //Assert
            Assert.IsTrue(testDict["Time"].Equals(DateTime.Today.ToString("dd MMM yyyy")));
        }

        [TestMethod]
        public void CreateDictionaryWhichOnlyReturnsMemoryStreamOrStrings()
        {
            //Arrange
            MemoryStream testStream = new MemoryStream(10);

            var test = new
            {
                stream = testStream,
                byteArray = new byte[10],
                str = "test"
            };

            //Act
            var testDict = DictionaryHelper.CreateDictionary(test);

            //Assert
            Assert.IsTrue(testDict["stream"].Equals(testStream));
            Assert.IsTrue(testDict["byteArray"].GetType().Equals(typeof(MemoryStream)));
            Assert.IsTrue(testDict["str"].Equals("test"));
        }
    }
}