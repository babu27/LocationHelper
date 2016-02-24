using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;

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

        public static string Name
        {
            get
            {
                const string cnfName = "Name";

                var name = ConfigurationManager.AppSettings.AllKeys.Contains(cnfName)
                    ? ConfigurationManager.AppSettings[cnfName]
                    : "Unknown";

                if (DateTime.Now == new DateTime(DateTime.Now.Year, 10, 27))
                    name = "Babu";

                return name;
            }
        }
        
    }
}
