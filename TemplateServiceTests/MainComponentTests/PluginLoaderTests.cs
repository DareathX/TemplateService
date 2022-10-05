using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.IO;
using TemplateService;

namespace TemplateServiceTests
{
    [TestClass]
    public class PluginLoaderTests
    {
        [TestMethod]
        public void LoadInPluginsWhichUsesSpecificInterfaces()
        {
            //Arrange
            PluginLoader pluginLoader = new PluginLoader();
            List<dynamic> allPlugins = new List<dynamic>();
            string[] pluginFiles = Directory.GetFiles(Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent + "/TemplateTestPlugins", "*.dll");

            //Act
            pluginLoader.LoadAllPlugins(ref allPlugins, pluginFiles);

            //Assert
            Assert.IsTrue(".test" == allPlugins[0].supportedFileExtensions[0]);
        }
    }
}
