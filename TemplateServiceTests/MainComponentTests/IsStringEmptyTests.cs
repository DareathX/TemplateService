using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using TemplateService;

namespace TemplateServiceTests
{
    [TestClass]
    public class IsStringEmptyTests
    {
        [TestMethod]
        public void IfStringIsNullThrowException()
        {
            //Arrange
            object testString = null;

            //Assert
            Assert.ThrowsException<Exception>(() => IsStringEmpty.SetStringEmptyOrThrowException(ref testString, false, "error"));
        }
    }
}
