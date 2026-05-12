using Prism.Events;
using Prism.Ioc;

using MSExcel = Microsoft.Office.Interop.Excel;

namespace VNCExcelToolsApplication
{
    public class Common : VNC.VSTOAddIn.Common
    {
        new public const string LOG_CATEGORY = "VNCExcelToolsApplication";

        public const string cCONFIG_FILE = @"C:\temp\VNCExcelTools.xml";

        public static Events.ExcelAppEvents AppEvents;
        public static Events.AddInApplicationEvents AddInApplicationEvents;

        public static MSExcel.Application ExcelApplication { get; set; }

        //// NOTE(crhodes)
        //// Add new VNC.Core.Information? InformationXXX
        //// for other Assemblies that are used as part of the application.
        ////
        //// Initialize GetAndSetInformation() in AddInApplication.cs
        ////
        //// Extend Views\AppVersionInfo.xaml as needed
        ////  Add new properties
        ////  Update InitializeViewModel()

        //// NOTE(crhodes)
        //// If we want this we have to set it from VNCExcelTools.ThisAddIn_Startup

        public static VNC.Core.Information? InformationVNCExcelTools;
        public static VNC.Core.Information? InformationVNCExcelToolsApplication;
        public static VNC.Core.Information? InformationVNCExcelToolsApplicationCore;

        public static VNC.Core.Information? InformationVNCVSTOAddInExcel;
        public static VNC.Core.Information? InformationVNCVSTOAddIn;

        public static VNC.Core.Information? InformationVNCWpfPresentation;
        public static VNC.Core.Information? InformationVNCWpfPresentationDx;

        public static IContainerProvider Container;

        public static IEventAggregator EventAggregator = new EventAggregator();
        public static VNCExcelToolsApplication.Bootstrapper ApplicationBootstrapper;

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
