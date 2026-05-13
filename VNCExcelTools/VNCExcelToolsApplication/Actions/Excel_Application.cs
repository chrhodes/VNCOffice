using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;

using Prism.Events;

using VNC;
using VNC.Core.Events;
using VNC.Core.Presentation;
using VNC.WPF.Presentation.Views;

using VNCExcelToolsApplication.Presentation.Views;

using MSExcel = Microsoft.Office.Interop.Excel;
using VNCExcelAddIn = VNC.VSTOAddIn.Excel;

namespace VNCExcelToolsApplication.Actions
{
    public class Excel_Application
    {
        public static IEventAggregator eventAggregator;
        public static DeveloperModeEvent developerModeEvent;
        public static StatusMessageEvent statusMessageEvent;

        public static WindowHost _aboutHost = null;

        public static void DeveloperModeUI(Boolean developerUIMode)
        {
            if (developerUIMode)
            {
                Common.DeveloperUIMode = Visibility.Visible;
            }
            else
            {
                Common.DeveloperUIMode = Visibility.Collapsed;
            }

            PublishDeveloperMode(developerUIMode);
        }

        public static void DisplayAddInInfo()
        {
            if (_aboutHost is null) _aboutHost = new WindowHost(Common.EventAggregator);

            _aboutHost.InformationApplication = Common.InformationApplication;
            _aboutHost.InformationApplicationCore = Common.InformationApplicationCore;
            _aboutHost.InformationVNCCore = Common.InformationVNCCore;

            // NOTE(crhodes)
            // About has About() and About(ViewModel) constructors.
            // If No DI Registrations - About() is called - does not wire View to ViewModel
            // If About DI Registrations - About() is called - does not wire View to ViewModel
            // If AboutViewModel DI Registrations - About(viewModel) is called - does wire View to ViewModel
            // NB.  AutoWireViewModel=false

            // NB. If AutoWireViewModel=true, the About() is called but then magically it is wired to ViewModel!

            //UserControl userControl = (Views.About)Common.Container.Resolve(typeof(Views.About));
            UserControl userControl = new About();

            _aboutHost.DisplayUserControlInHost(
                "VNCExcelTools About",
                    Common.DEFAULT_WINDOW_WIDTH, Common.DEFAULT_WINDOW_HEIGHT,
                //(Int32)userControl.Width + Common.WINDOW_HOSTING_USER_CONTROL_WIDTH_PAD,
                //(Int32)userControl.Height + Common.WINDOW_HOSTING_USER_CONTROL_HEIGHT_PAD,
                ShowWindowMode.Modeless_Show,
                userControl
            );
        }

        public static void DisplayInfo()
        {
            Common.WriteToDebugWindow($"{System.Reflection.MethodInfo.GetCurrentMethod().Name}");

            MSExcel.Application app = Common.ExcelApplication;

            StringBuilder sb = new StringBuilder();

            Common.WriteToDebugWindow($"App.Name - {app.Name}");

            // try
            // {
            // Common.WriteToDebugWindow($"App.ActiveDocument.Name - {app.ActiveDocument.Name}");
            // }
            // catch (Exception ex)
            // {
            // Common.WriteToDebugWindow("App.ActiveDocument.Name - <none>");
            // }

            // try
            // {
            // Common.WriteToDebugWindow($"App.ActivePage.Name - {app.ActivePage.Name}");
            // }
            // catch (Exception ex)
            // {
            // Common.WriteToDebugWindow("App.ActivePage.Name - <none>");
            // }

            //System.Windows.Forms.MessageBox.Show(sb.ToString());
            Common.WriteToDebugWindow(sb.ToString());
        }

        // NOTE(crhodes)
        // Publish DeveloperModeEvent for things that don't have access to Common.DeveloperMode
        // e.g. Host Windows

        private static void PublishDeveloperMode(Boolean developerMode)
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.Event) startTicks = Log.EVENT("Enter", Common.LOG_CATEGORY);

            developerModeEvent.Publish(developerMode);

            if (Common.VNCLogging.Event) Log.EVENT("Enter", Common.LOG_CATEGORY, startTicks);
        }
    }
}
