using System.Reflection;
using System.Runtime.Loader;

namespace TemplateService
{
    public class PluginLoader
    {
        /// <summary>
        /// Checks if the plugin has a supported interface.
        /// </summary>
        /// <typeparam name="T">IFillInTemplates or IConvertTemplateTo.</typeparam>
        /// <param name="plugins">List to place supported plugins into.</param>
        /// <param name="pluginFiles">All dll to be checked.</param>
        /// <returns></returns>
        private List<T> LoadPlugins<T>(List<T> plugins, string[] pluginFiles)
        {
            plugins = new List<T>();

            foreach (string? dll in pluginFiles)
            {
                AssemblyLoadContext assemblyLoadContext = new AssemblyLoadContext(dll);
                Assembly assembly = assemblyLoadContext.LoadFromAssemblyPath(dll);
                Type[]? pluginFromAssembly = assembly
                    .GetTypes()
                    .Where(t => (typeof(T).IsAssignableFrom(t)) && !t.IsInterface)
                    .ToArray();

                CreatePluginInstance(ref plugins, pluginFromAssembly);
            }
            return plugins;
        }

        /// <summary>
        /// Creates instance of plugin.
        /// </summary>
        /// <typeparam name="T">IFillInTemplates or IConvertTemplateTo.</typeparam>
        /// <param name="plugins">List to place supported plugins into.</param>
        /// <param name="pluginFromAssembly"></param>
        private void CreatePluginInstance<T>(ref List<T> plugins, Type[]? pluginFromAssembly)
        {
            foreach (Type? pluginType in pluginFromAssembly)
            {
                plugins.Add((T)Activator.CreateInstance(pluginType));
            }
        }

        /// <summary>
        /// Loads all plugins with the interfaces IFillInTemplates and IConvertTemplateTo.
        /// </summary>
        /// <param name="plugins">Gives list of all found supported plugins.</param>
        /// <param name="pluginFiles">All dll to be checked.</param>
        public bool LoadAllPlugins(ref List<object> plugins, string[] pluginFiles)
        {
            List<dynamic> allPluginTypes = new List<dynamic>
            {
                new List<IFillInTemplates>(),
                new List<IConvertTemplateTo>()
            };
            
            foreach (var pluginTypes in allPluginTypes)
            {
                plugins.AddRange(LoadPlugins(pluginTypes, pluginFiles));
            }

            foreach (var plugin in plugins)
            {
                if (plugin.ToString().Contains("FillIn"))
                {
                    return false;
                }
            }

            return true;
        }
}
}
