using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using TemplateService;

namespace TemplateServiceTests.MainComponentTests
{
    [TestClass]
    public class ProcessTemplateRequestTests
    {
        private Dictionary<string, object> _templateDict = new Dictionary<string, object>() { { "Test", "Testing" } };

        [TestMethod]
        public void CheckIfExceptonIsThrownWhenDictionaryIsNotGiven()
        {
            //Assign
            byte[] bytes = new byte[0];

            //Assert
            Exception exception = Assert.ThrowsException<Exception>(() => CreateNewRequest(new TemplateServiceInputData(), ref bytes));
            Assert.AreEqual("Missing data InputDataDict", exception.Message);
        }

        [TestMethod]
        public void CheckIfExceptonIsThrownWhenPathOrBytesAreNotGiven()
        {
            //Assign
            byte[] bytes = new byte[0];

            //Assert
            Exception exception = Assert.ThrowsException<Exception>(() => CreateNewRequest(new TemplateServiceInputData() with
            {
                InputDataDict = _templateDict
            }, ref bytes));
            Assert.AreEqual("Missing data SetTemplatePathOrBytes", exception.Message);
        }

        [TestMethod]
        public void CheckIfTheBytesAreCorrectlyReturnedWithoutConversion()
        {
            //Assign
            byte[] bytes = new byte[0];

            //Act
            ProcessTemplateRequest testProcess = new ProcessTemplateRequest(new TemplateServiceInputData() with
            {
                InputDataDict = _templateDict,
                SetTemplatePathOrBytes = new byte[0],
                SetConversionFromTo = ( ".test", ".test" ),
                SetPluginFolderPath = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent + "/TemplateTestPlugins"
        }, ref bytes);

            //Assert
            Assert.IsTrue(bytes.Length == 10 && bytes[9] == 1);
        }

        [TestMethod]
        public void CheckIfTheBytesAreCorrectlyReturnedWithConversion()
        {
            //Assign
            byte[] bytes = new byte[0];

            //Act
            CreateNewRequest(new TemplateServiceInputData() with
            {
                InputDataDict = _templateDict,
                SetTemplatePathOrBytes = new byte[0],
                SetConversionFromTo = (".test", ".test2"),
                SetPluginFolderPath = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent + "/TemplateTestPlugins"
            }, ref bytes);

            //Assert
            Assert.IsTrue(bytes.Length == 11 && bytes[10] == 1);
        }

        private void CreateNewRequest(TemplateServiceInputData data, ref byte[] bytes)
        {
            new ProcessTemplateRequest(data, ref bytes);
        }
    }
}
