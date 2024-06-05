using System;
using System.Security.Principal;
//using Prism.Events;
//using VNCExcelAddin.User_Interface;

namespace VNCExcelAddin
{
    ///<summary>
    ///Common items declared at the Class level.
    ///</summary>
    ///<remarks>
    ///Use this class for any thing you want globally available.
    ///Place only Static items in this class.  This Class cannot not be instantiated.
    ///</remarks>  
    ///
    public class Common // : VNC.AddinHelper.Common
    {
        //new public const string PROJECT_NAME = "VNCExcelAddin";
        public const string LOG_CATEGORY = "VNCExcelAddin";

        public static Boolean HasAppEvents = true;  // Custom Header and Footer need this enabled.
        public static Boolean DisplayEvents = false;

        public static Boolean DebugSQL
        {
            get;
            set;
        }
        public static Boolean DebugLevel1
        {
            get;
            set;
        }
        public static Boolean DebugLevel2
        {
            get;
            set;
        }
        public static Boolean DebugMode
        {
            get;
            set;
        }
        public static Boolean DeveloperMode
        {
            get;
            set;
        }

        public static Boolean DisplayXlLocationUpdates
        {
            get;
            set;
        }

        public static Boolean EnableLogging
        {
            get;
            set;
        }

        ////public const string cCONFIG_FILE = @"C:\temp\SupportTools_Config.xml";
        //public const string cCONFIG_FILE = @"C:\temp\VNCExcelAddin.xml";
        //public const string cEXPORT_TEMPLATE_PATH = @"C:\temp\AZDO-TFS";

        //public const string cGO_BACK_DAYS = "7";

        //public static System.Windows.Application XamlApplication;

        //public static VNC.AddinHelper.Excel ExcelHelper = new VNC.AddinHelper.Excel();
        //public static Events.ExcelAppEvents AppEvents;

        //public const int cMaxFileNameLength = 128;

        //public const string cDEFAULT_FONT = "Calibri";

        //public static IEventAggregator EventAggregator = new EventAggregator();
        //public static Application.Bootstrapper ApplicationBootstrapper;

        //// These values are added to the dimensions of a hosting window if the
        //// hosted User_Control specifies values for MinWidth/MinHeight.
        //// They have not been thought through but do seem to "work".

        //internal const int DEFAULT_WINDOW_WIDTH_LARGE = 1800;
        //internal const int DEFAULT_WINDOW_HEIGHT_LARGE = 1200;

        //internal const int DEFAULT_WINDOW_WIDTH = 900;
        //internal const int DEFAULT_WINDOW_HEIGHT = 600;

        //internal const int DEFAULT_WINDOW_WIDTH_SMALL = 450;
        //internal const int DEFAULT_WINDOW_HEIGHT_SMALL = 300;

        //internal const int WINDOW_HOSTING_USER_CONTROL_WIDTH_PAD = 30;
        //internal const int WINDOW_HOSTING_USER_CONTROL_HEIGHT_PAD = 75;

        //public static IPrincipal CurrentUser
        //{
        //    get;
        //    set;
        //}

        //internal static string PriorStatusBar;

        //public static event EventHandler AutoHideGroupSpeedChanged;

        //private static int _AutoHideGroupSpeed = 250;

        //public static int AutoHideGroupSpeed
        //{
        //    get { return _AutoHideGroupSpeed; }
        //    set
        //    {
        //        _AutoHideGroupSpeed = value;

        //        EventHandler evt = AutoHideGroupSpeedChanged;

        //        if (evt != null)
        //        {
        //            evt(null, EventArgs.Empty); ;
        //        }
        //    }
        //}

        //// This controls the behavior of the overall application.
        //// It is initialized from app.config and is updated when the user changes the mode.
        //// Changes are reflected in the app.config file.

        //public static ViewMode UserMode { get; set; }

        //public static bool IsAdministrator { get; set; }
        //public static bool IsBetaUser { get; set; }
        //public static bool IsDeveloper { get; set; }
        ////public static bool IsAdvancedUser { get; set; }

        //public static bool AllowEditing { get; set; }

        //public static string RowDetailMode { get; set; }

        //private static Data.ApplicationDS _ApplicationDS;

        //public static Data.ApplicationDS ApplicationDS
        //{
        //    get
        //    {
        //        if (_ApplicationDS == null)
        //        {
        //            _ApplicationDS = new Data.ApplicationDS();
        //        }
        //        return _ApplicationDS;
        //    }
        //    set
        //    {
        //        _ApplicationDS = value;
        //    }
        //}

        //// TaskPane specific stuff

        //public static Microsoft.Office.Tools.CustomTaskPane TaskPaneActiveDirectory;

        ////public static Microsoft.Office.Tools.CustomTaskPane TaskPaneConfig;

        //public static Microsoft.Office.Tools.CustomTaskPane TaskPaneDevelopment;


        //public static Microsoft.Office.Tools.CustomTaskPane TaskPaneAppUtilities;

        ////public static Microsoft.Office.Tools.CustomTaskPane TaskPaneHelp;

        //public static Microsoft.Office.Tools.CustomTaskPane TaskPaneLogParser;

        //public static Microsoft.Office.Tools.CustomTaskPane TaskPaneLTC;

        //public static Microsoft.Office.Tools.CustomTaskPane TaskPaneNetworkTrace;

        //public static Microsoft.Office.Tools.CustomTaskPane TaskPaneSharePoint;

        //public static Microsoft.Office.Tools.CustomTaskPane TaskPaneSMO;

        //public static Microsoft.Office.Tools.CustomTaskPane TaskPaneSQLSMO;

        //public static Microsoft.Office.Tools.CustomTaskPane TaskPaneTFS;

        //public static Microsoft.Office.Tools.CustomTaskPane TaskPaneUtilities;


    }
}
