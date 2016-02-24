using System;
using System.Configuration;

namespace LocationHelper.Helper
{
    internal static class ConfigHelper
    {
        public static string LatitudeColumn
        {
            get
            {
                return ConfigurationManager.AppSettings["LatitudeColumn"];
            }
        }

        public static string LongitudeColumn
        {
            get { return ConfigurationManager.AppSettings["LongitudeColumn"]; }
        }

        public static int HeaderRowNumber
        {
            get { return Convert.ToInt32(ConfigurationManager.AppSettings["HeaderRowNumber"]); }
        }

        public static int FirstDataRowNumber
        {
            get { return Convert.ToInt32(ConfigurationManager.AppSettings["FirstDataRowNumber"]); }
        }

        public static int DataSheetNumber
        {
            get { return Convert.ToInt32(ConfigurationManager.AppSettings["DataSheetNumber"]); }
        }

    }
}
