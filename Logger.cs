using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using HPMSdk;

namespace Hansoft.SdkUtils
{
    class Logger
    {
        /// <summary>
        /// Private constructor, this is a pure utility class
        /// </summary>
        private Logger()
        {
        }

        public static void Information(string message)
        {
            DisplayMessage("Information: " + message);
        }

        public static void Warning(string message)
        {
            DisplayMessage("Warning: " + message);
        }

        public static void Error(string message)
        {
            DisplayMessage("Error: " + message);
        }

        public static void Exception(Exception e)
        {
            Exception("", e);
        }

        public static void Exception(string message, Exception e)
        {
            DisplayMessage("Exception: " + message);
            if (e is HPMSdkException)
            {
                HPMSdkException hpme = (HPMSdkException)e;
                Error(hpme.ErrorAsStr());
            }
            else
                Error(e.Message);
        }

        public static void DisplayMessage(string message)
        {
            Console.WriteLine(message);
        }

    }
}
