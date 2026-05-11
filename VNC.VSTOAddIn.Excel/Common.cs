using System.Windows;

namespace VNC.VSTOAddIn.Excel
{
    public class Common : VNC.VSTOAddIn.Common
    {
        new public const string LOG_CATEGORY = "VNCVSTOAddInExcel";
        public static Visibility DeveloperUIMode
        {
            get;
            set;
        }
    }
}
