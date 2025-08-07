using System.Windows;

namespace VNC.Visio.VSTOAddIn
{
    public class Common : VNC.VSTOAddIn.Common
    {
        new public const string LOG_CATEGORY = "VNCVisioVSTOAddIn";

        public static Visibility DeveloperUIMode
        {
            get;
            set;
        }
    }
}
