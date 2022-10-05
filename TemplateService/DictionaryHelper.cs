using System.Reflection;

namespace TemplateService
{
    public class DictionaryHelper
    {
        /// <summary>
        /// Adds formating to every Dictionary entry.
        /// </summary>
        /// <param name="dict">Dictionary with all placeholders.</param>
        /// <param name="dateTimeFormat">Set disired DateTime format</param>
        /// <returns>Returns formattted dictionary with placeholders.</returns>
        private static Dictionary<string, object> AddFormatingToNewDictionary(Dictionary<string, object> dict, string dateTimeFormat)
        {
            Dictionary<string, object> newDict = new Dictionary<string, object>();
            foreach (var entry in dict)
            {
                newDict.Add(entry.Key, AddFormating(entry.Value, dateTimeFormat));
            }
            return newDict;
        }

        /// <summary>
        /// Adds formating to the property values.
        /// </summary>
        /// <param name="dict">Dictionary with all placeholders.</param>
        /// <returns>Returns formattted dictionary with placeholders.</returns>
        public static Dictionary<string, object> AddFormatingToDictionary(Dictionary<string, object> dict)
        {
            return AddFormatingToNewDictionary(dict, "");
        }

        /// <summary>
        /// Adds formating to the property values.
        /// </summary>
        /// <param name="dict">Dictionary with all placeholders.</param>
        /// <param name="dateTimeFormat">Set disired DateTime format</param>
        /// <returns>Returns formattted dictionary with placeholders.</returns>
        public static Dictionary<string, object> AddFormatingToDictionary(Dictionary<string, object> dict, string dateTimeFormat)
        {
            return AddFormatingToNewDictionary(dict, dateTimeFormat);
        }

        /// <summary>
        /// Receives a class object, which will be used to generate a diciotnary based on reflection. 
        /// </summary>
        /// <typeparam name="T">class</typeparam>
        /// <param name="newObject">Object where reflection will be used on.</param>
        /// <param name="dateTimeFormat">Format that will be used for DateTime values</param>
        /// <returns>Returns a dictionary that contains all propteries of given object.</returns>
        private static Dictionary<string, object> CreateNewDictionary<T>(T newObject, string dateTimeFormat) where T : class
        {
            Dictionary<string, object> dict = new Dictionary<string, object>();

            Type type = newObject.GetType();
            MemberInfo[] memberInfoProperties = type.GetProperties();
            MemberInfo[] allMemberInfo = new MemberInfo[memberInfoProperties.Length];
            memberInfoProperties.CopyTo(allMemberInfo, 0);

            foreach (MemberInfo memberInfo in allMemberInfo)
            {
                switch (memberInfo.MemberType)
                {
                    case MemberTypes.Property:
                        PropertyInfo propertyInfo = (PropertyInfo)memberInfo;
                        dict.Add(propertyInfo.Name.ToString(), AddFormating(propertyInfo.GetValue(newObject), dateTimeFormat));
                        break;
                    default:
                        break;
                }
            }

            return dict;
        }

        /// <summary>
        /// Receives a class object, which will be used to generate a Dictionary based on reflection. 
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="newObject">Object where reflection will be used on.</param>
        /// <returns>Returns a dictionary with the property Names as Key and property Values as value.</returns>
        public static Dictionary<string, object> CreateDictionary<T>(T newObject) where T : class
        {
            return CreateNewDictionary(newObject, "");
        }

        /// <summary>
        /// Receives a class object, which will be used to generate a diciotnary based on reflection. 
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="newObject">Object where reflection will be used on.</param>
        /// <param name="dateTimeFormat">Format that will be used for DateTime values</param>
        /// <returns>Returns a dictionary with the property Names as Key and property Values as value.</returns>
        public static Dictionary<string, object> CreateDictionary<T>(T newObject, string dateTimeFormat) where T : class
        {
            return CreateNewDictionary(newObject, dateTimeFormat);
        }

        /// <summary>
        /// Adds formating to the property value.
        /// </summary>
        /// <param name="type">Object type where formating is being applied on.</param>
        /// <param name="format">Format is being used for a DateTime.</param>
        /// <returns></returns>
        private static object AddFormating(object type, string format)
        {
            object value;
            switch (type)
            {
                case DateTime:
                    DateTime dateTime = (DateTime)type;
                    value = string.IsNullOrEmpty(format) ? dateTime.ToString("dd-MM-yyyy") : dateTime.ToString(format);
                    break;
                case byte[] bytes:
                    value = new MemoryStream(bytes);
                    break;
                case MemoryStream stream:
                    value = stream;
                    break;
                default:
                    value = type.ToString() ?? "";
                    break;
            }

            return value;
        }
    }
}
