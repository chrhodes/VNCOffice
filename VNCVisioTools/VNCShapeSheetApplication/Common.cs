using System.Windows;

using Prism.Events;

using MSVisio = Microsoft.Office.Interop.Visio;

namespace VNCShapeSheetApplication
{
    public class Common : VNC.VSTOAddIn.Common
    {
        new public const string LOG_CATEGORY = "VNCShapeSheetApplication";

        public const string cCONFIG_FILE = @"C:\temp\VNC.Visio.VSTOAddIn.xml";

        public static Events.VisioAppEvents AppEvents;
        public static Events.AddInApplicationEvents AddInApplicationEvents;

        public static IEventAggregator EventAggregator = new EventAggregator();
        public static Bootstrapper ApplicationBootstrapper;

        public static MSVisio.Application VisioApplication { get; set; }

        internal const int DEFAULT_WINDOW_WIDTH_LARGE = 1800;
        internal const int DEFAULT_WINDOW_HEIGHT_LARGE = 1200;

        internal const int DEFAULT_WINDOW_WIDTH = 900;
        internal const int DEFAULT_WINDOW_HEIGHT = 600;

        internal const int DEFAULT_WINDOW_WIDTH_SMALL = 450;
        internal const int DEFAULT_WINDOW_HEIGHT_SMALL = 300;

        internal const int WINDOW_HOSTING_USER_CONTROL_WIDTH_PAD = 30;
        internal const int WINDOW_HOSTING_USER_CONTROL_HEIGHT_PAD = 75;

    }
}