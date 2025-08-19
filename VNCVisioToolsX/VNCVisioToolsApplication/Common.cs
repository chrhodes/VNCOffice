using System.Windows;

using Prism.Events;
using Prism.Ioc;

using MSVisio = Microsoft.Office.Interop.Visio;

namespace VNCVisioToolsApplication
{
    public class Common : VNC.VSTOAddIn.Common
    {
        new public const string LOG_CATEGORY = "VNCVisioToolsApplication";

        public const string cCONFIG_FILE = @"C:\temp\VNCVisioToolsApplication.xml";

        // NOTE(crhodes)
        // Add new VNC.Core.Information? InformationXXX
        // for other Assemblies that are used as part of the application.
        //
        // Initialize GetAndSetInformation() in AddInApplication.cs 
        //
        // Extend Views\AppVersionInfo.xaml as needed
        //  Add new properties
        //  Update InitializeViewModel()

        //public static VNC.Core.Information? InformationVNCVisioToolsApplicationCore;

        public static VNC.Core.Information? InformationVNCVisioVSTOAddIn;
        public static VNC.Core.Information? InformationVNCVSTOAddIn;
        public static VNC.Core.Information? InformationVNCWpfPresentation;
        public static VNC.Core.Information? InformationVNCWpfPresentationDx;

        public static IContainerProvider Container;


        public static Events.VisioAppEvents AppEvents;
        public static Events.AddInApplicationEvents AddInApplicationEvents;

        public static IEventAggregator EventAggregator = new EventAggregator();
        public static VNCVisioToolsApplication.Bootstrapper ApplicationBootstrapper;

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
