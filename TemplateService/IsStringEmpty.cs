using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TemplateService
{
    public class IsStringEmpty
    {
        /// <summary>
        /// Based on de bool parameter, sets the string as a space or throws exception if the object is empty or null.
        /// </summary>
        /// <param name="text">Object that will be checked.</param>
        /// <param name="isEmtpyAllowed">True = string.Empty and False = exception.</param>
        /// <param name="exceptionMessage">Message that the exception will return.</param>
        /// <exception cref="Exception"></exception>
        public static void SetStringEmptyOrThrowException(ref object text, bool isEmtpyAllowed, string exceptionMessage)
        {
            if (text == null || (text as string) == string.Empty)
            {
                if (isEmtpyAllowed)
                {
                    text = " ";//Gets set to a space, because an empty string may give errors with documents.
                }
                else
                {
                    throw new Exception(exceptionMessage);
                }
            }
        }
    }
}
