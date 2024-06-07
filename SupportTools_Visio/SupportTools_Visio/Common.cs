using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Principal;
using System.Text;
using Prism.Events;

namespace SupportTools_Visio
{
    public class Common : VNC.Visio.VSTOAddIn.Common
    {

        #region SMARTS

        public const int EMPTY_INT = 0;
        public const string EMPTY_STR = "''";

        public const int TABLE_ABSTRACTION = 1;
        public const int TABLE_ATTRIBUTE = 2;
        public const int TABLE_SHAPETYPE = 3;
        public const int TABLE_SHAPE = 4;
        public const int TABLE_SHAPETYPEATTRIBUTE = 5;
        public const int TABLE_ATTRIBUTEVALUE = 6;
        public const int TABLE_RELATION = 7;
        public const int TABLE_SDTABLE = 8;
        public const int TABLE_LOG = 9;
        public const int TABLE_CONSTRAINEDVALUES = 10;
        public const int TABLE_ARTIFACT = 11;
        public const int TABLE_CONTAINER = 12;
        public const int TABLE_SDLOGFUNCTION = 13;
        public const int TABLE_LIFECYCLE = 14;

        public const int SHAPETYPE_APPLICATION = 1;
        public const int SHAPETYPE_DATA = 2;
        public const int SHAPETYPE_BATCHFILE = 3;
        public const int SHAPETYPE_ACTOR = 4;
        public const int SHAPETYPE_FILEREPORT = 5;
        public const int SHAPETYPE_QUEUE = 6;
        public const int SHAPETYPE_JOB = 7;
        public const int SHAPETYPE_OBJECTRELATION = 8;
        public const int SHAPETYPE_SYSTEM = 9;
        public const int SHAPETYPE_SYSTEMRELATION = 10;
        public const int SHAPETYPE_SERVER = 11;
        public const int SHAPETYPE_IP = 12;
        public const int SHAPETYPE_DEVICERELATION = 13;
        public const int SHAPETYPE_DESKTOP = 14;
        public const int SHAPETYPE_MAINFRAME = 15;
        public const int SHAPETYPE_LAPTOP = 16;
        public const int SHAPETYPE_PDA = 17;
        public const int SHAPETYPE_TABLET = 18;
        public const int SHAPETYPE_PRINTER = 19;
        public const int SHAPETYPE_SAN = 20;
        public const int SHAPETYPE_FIREWALL = 21;
        public const int SHAPETYPE_LOADBALANCER = 22;
        public const int SHAPETYPE_ROUTER = 23;
        public const int SHAPETYPE_PIECE = 24;
        public const int SHAPETYPE_BUSINESS = 25;
        public const int SHAPETYPE_STEP = 26;
        public const int SHAPETYPE_DECISION = 27;
        public const int SHAPETYPE_BUSINESSRELATION = 28;
        public const int SHAPETYPE_STEPRELATION = 29;
        public const int SHAPETYPE_CROSSRELATION = 30;
        public const int SHAPETYPE_TIME = 31;
        public const int SHAPETYPE_START = 32;
        public const int SHAPETYPE_END = 33;
        public const int SHAPETYPE_SUBSTEP = 34;
        public const int SHAPETYPE_BROWSER = 35;
        public const int SHAPETYPE_SQLINSTANCE = 36;
        public const int SHAPETYPE_FOLDER = 37;
        public const int SHAPETYPE_FILE = 38;
        public const int SHAPETYPE_REPORT = 39;

        public const int ABSTRACTION_SYSTEM = 1;
        public const int ABSTRACTION_OBJECT = 2;
        public const int ABSTRACTION_DEVICE = 3;
        public const int ABSTRACTION_SYSTEMOBJECT = 4;
        public const int ABSTRACTION_IP = 5;
        public const int ABSTRACTION_DEVICEOBJECT = 6;
        public const int ABSTRACTION_ARTIFACT = 7;
        public const int ABSTRACTION_PIECE = 8;
        public const int ABSTRACTION_OBJECTPIECE = 9;
        public const int ABSTRACTION_BUSINESS = 10;
        public const int ABSTRACTION_STEP = 11;
        public const int ABSTRACTION_BUSINESSSTEP = 12;
        public const int ABSTRACTION_RELATION = 13;
        public const int ABSTRACTION_STEPOBJECT = 14;
        public const int ABSTRACTION_SUBSTEP = 15;

        public const int VISIO_ENABLED = 0;
        public const int VISIO_DISABLED = 1;

        public const int ACTIONITEM_DRILLDOWN = 0;
        public const int ACTIONITEM_DRILLUP = 1;
        public const int ACTIONITEM_EDITVISIO = 2;
        public const int ACTIONITEM_PROPERTIES = 4;
        public const int ACTIONITEM_BUSINESSPROCESS = 5;
        public const int ACTIONITEM_INFRASTRUCTURE = 6;
        public const int ACTIONITEM_SYSTEM = 7;
        public const int ACTIONITEM_RETRIEVE = 8;
        public const int ACTIONITEM_VALIDATESHAPE = 9;
        public const int ACTIONITEM_SHAPEDROP = 10;
        public const int ACTIONITEM_RECONNECT = 11;
        public const int ACTIONITEM_SUMMARIZE = 12;
        public const int ACTIONITEM_VALIDATESHEET = 13;
        public const int ACTIONITEM_SHOWRELATIONS = 14;

        public const int RELATIONTYPE_CONTAINER = 1;
        public const int RELATIONTYPE_CONTAINED = 2;
        public const int RELATIONTYPE_DIRECT = 3;

        public const int LINEFORMAT_LINEPATTERN_SOLID = 1;
        public const int LINEFORMAT_LINEPATTERN_SLASH = 2;
        public const int LINEFORMAT_LINEPATTERN_DOT = 3;
        public const string LINEFORMAT_LINEWEIGHT = "1.5 pt";
        public const int LINEFORMAT_ARROW_NONE = 0;
        public const int LINEFORMAT_ARROW_NORMAL = 1;

        #endregion

        public const string LOG_CATEGORY = "SupportTools_Visio";

        public const string cCONFIG_FILE = @"C:\temp\SupportTools_Visio.xml";

        // HACK(crhodes)
        // What is this used for?

        //public static VNC.Visio.VSTOAddIn VNCVisioAddIn = new VNC.Visio.VSTOAddIn();
        public static Events.VisioAppEvents AppEvents;
        public static bool DisplayChattyEvents = false;

        public static IEventAggregator EventAggregator = new EventAggregator();
        public static Application.Bootstrapper ApplicationBootstrapper;

        //public static Microsoft.Office.Tools.CustomTaskPane TaskPaneConfig;

        //public static Microsoft.Office.Tools.CustomTaskPane TaskPaneHelp;

        public static Microsoft.Office.Tools.CustomTaskPane TaskPaneAppUtil;

        #region Core

        // TODO: Add as many DebugLevels as needed.
        // Add accompanying checkboxes on frmDebugWindow
        //
        // NB. AddInHelper provides DebugLevel1, DebugLevel2, DebugSQL

        //public static bool DebugLevel1
        //{
        //    get;
        //    set;
        //}
        //public static bool DebugLevel2
        //{
        //    get;
        //    set;
        //}
        //public static bool DebugSQL
        //{
        //    get;
        //    set;
        //}

        public static event EventHandler AutoHideGroupSpeedChanged;

        private static int _AutoHideGroupSpeed = 250;

        public static int AutoHideGroupSpeed
        {
            get { return _AutoHideGroupSpeed; }
            set
            {
                _AutoHideGroupSpeed = value;

                EventHandler evt = AutoHideGroupSpeedChanged;

                if (evt != null)
                {
                    evt(null, EventArgs.Empty); ;
                }
            }
        }

        //public static void RaiseAutoHideGroupSpeedChanged()
        //{
        //    EventHandler evt = AutoHideGroupSpeedChanged;

        //    if (evt != null)
        //    {
        //        evt(null, EventArgs.Empty); ;
        //    }
        //}

        // NB. These are in Common.AddInHelper

        ///// <summary>
        ///// Indicates whether the UI is running in DeveloperMode
        ///// </summary>
        //public static bool DeveloperMode
        //{
        //    get;
        //    set;
        //}
        ///// <summary>
        ///// Indicates whether the UI is running in DebugMode
        ///// </summary>
        //public static bool DebugMode
        //{
        //    get;
        //    set;
        //}

        //private static OutputWindow _outputWindow;

        //public static OutputWindow OutputWindow
        //{
        //    get
        //    {
        //        if(_outputWindow == null)
        //        {
        //            _outputWindow = new OutputWindow();
        //        }
        //        return _outputWindow;
        //    }
        //    set
        //    {
        //        _outputWindow = value;
        //    }
        //}

        //public static long WriteToDebugWindow(string message)
        //{
        //    return OutputWindow.Write(message);
        //}

        //public static long WriteToDebugWindow(string message, long startTime)
        //{
        //    return OutputWindow.Write(message, startTime);
        //}

        #endregion

        // These values are added to the dimensions of a hosting window if the
        // hosted User_Control specifies values for MinWidth/MinHeight.
        // They have not been thought through but do seem to "work".

        internal const int DEFAULT_WINDOW_WIDTH_LARGE = 1800;
        internal const int DEFAULT_WINDOW_HEIGHT_LARGE = 1200;

        internal const int DEFAULT_WINDOW_WIDTH = 900;
        internal const int DEFAULT_WINDOW_HEIGHT = 600;

        internal const int DEFAULT_WINDOW_WIDTH_SMALL = 450;
        internal const int DEFAULT_WINDOW_HEIGHT_SMALL = 300;

        internal const int WINDOW_HOSTING_USER_CONTROL_WIDTH_PAD = 30;
        internal const int WINDOW_HOSTING_USER_CONTROL_HEIGHT_PAD = 75;

        // This controls the behavior of the overall application.
        // It is initialized from app.config and is updated when the user changes the mode.
        // Changes are reflected in the app.config file.

        public static IPrincipal CurrentUser
        {
            get;
            set;
        }

        //public static User_Interface.ViewMode UserMode { get; set; }

        public static bool IsAdministrator { get; set; }
        public static bool IsBetaUser { get; set; }
        public static bool IsDeveloper { get; set; }
        //public static bool IsAdvancedUser { get; set; }

        public static bool AllowEditing { get; set; }

        public static string RowDetailMode { get; set; }

        private static bool _DataFullyLoaded = false;
        public static bool DataFullyLoaded
        {
            get { return _DataFullyLoaded; }
            set
            {
                _DataFullyLoaded = value;
            }
        }

        // TODO(crhodes): Get rid of this (I think) and use the one from SupportTools_Visio.  See if need anything else
        // in  a DataSet first.  May want a separate one for the App.
        private static Data.ApplicationDS _ApplicationDataSet;
        public static Data.ApplicationDS ApplicationDataSet
        {
            get
            {
                if (_ApplicationDataSet == null)
                {
                    //_ApplicationDataSet = new SupportTools_Visio.Data.ApplicationDataSet();
                    _ApplicationDataSet = new Data.ApplicationDS();

                    // TODO: Add any other initialization of things related to the ApplicationDS
                }

                return _ApplicationDataSet;
            }
            set
            {
                _ApplicationDataSet = value;
            }
        }
    }
}
