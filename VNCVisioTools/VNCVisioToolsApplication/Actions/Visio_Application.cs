using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows;
using System.Windows.Controls;

using Prism.Events;

using VNC;
using VNC.Core.Events;
using VNC.Core.Presentation;
using VNC.WPF.Presentation.Views;

using VNCVisioToolsApplication.Presentation.Views;

using MSVisio = Microsoft.Office.Interop.Visio;
using VNCVisioAddIn = VNC.Visio.VSTOAddIn;

namespace VNCVisioToolsApplication.Actions
{
    public class Visio_Application
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
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

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
                "VNCVisioToolsApplication About",
                    Common.DEFAULT_WINDOW_WIDTH, Common.DEFAULT_WINDOW_HEIGHT,
                //(Int32)userControl.Width + Common.WINDOW_HOSTING_USER_CONTROL_WIDTH_PAD,
                //(Int32)userControl.Height + Common.WINDOW_HOSTING_USER_CONTROL_HEIGHT_PAD,
                ShowWindowMode.Modeless_Show,
                userControl
            );
        }

        public static void DisplayInfo()
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            MSVisio.Application app = Common.VisioApplication;

            StringBuilder sb = new StringBuilder();

            Common.WriteToDebugWindow($"App.Name - {app.Name}");

            try
            {
                Common.WriteToDebugWindow($"App.ActiveDocument.Name - {app.ActiveDocument.Name}");
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow("App.ActiveDocument.Name - <none>");
            }

            try
            {
                Common.WriteToDebugWindow($"App.ActivePage.Name - {app.ActivePage.Name}");
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow("App.ActivePage.Name - <none>");
            }

            Common.WriteToDebugWindow($"App.AddonPaths - {app.AddonPaths}");
            Common.WriteToDebugWindow($"App.CommandLine - {app.CommandLine}");
            Common.WriteToDebugWindow($"App.Documents.Count - {app.Documents.Count}");
            Common.WriteToDebugWindow($"App.DrawingPaths - {app.DrawingPaths}");
            Common.WriteToDebugWindow($"App.HelpPaths - {app.HelpPaths}");
            Common.WriteToDebugWindow($"App.IsVisio32 - {app.IsVisio32}");
            Common.WriteToDebugWindow($"App.MyShapesPath - {app.MyShapesPath}");
            Common.WriteToDebugWindow($"App.Path - {app.Path}");
            Common.WriteToDebugWindow($"App.ProcessID - {app.ProcessID}");
            Common.WriteToDebugWindow($"App.ShowChanges - {app.ShowChanges}");
            Common.WriteToDebugWindow($"App.ShowProgress - {app.ShowProgress}");
            Common.WriteToDebugWindow($"App.ShowStatusBar - {app.ShowStatusBar}");
            Common.WriteToDebugWindow($"App.ShowToolBar - {app.ShowToolbar}");
            Common.WriteToDebugWindow($"App.StartupPaths - {app.StartupPaths}");
            Common.WriteToDebugWindow($"App.StencilPaths - {app.StencilPaths}");
            Common.WriteToDebugWindow($"App.TemplatePaths - {app.TemplatePaths}");
            Common.WriteToDebugWindow($"App.TraceFlags - {app.TraceFlags}");
            Common.WriteToDebugWindow($"App.UndoEnables - {app.UndoEnabled}");
            Common.WriteToDebugWindow($"App.UserName - {app.UserName}");
            Common.WriteToDebugWindow($"App.Version - {app.Version}");

            //System.Windows.Forms.MessageBox.Show(sb.ToString());
            Common.WriteToDebugWindow(sb.ToString());
        }

        public static MSVisio.Document GetActiveDocument()
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            MSVisio.Application app = Common.VisioApplication;

            return app.ActiveDocument;
        }

        public static List<MSVisio.Document> GetOpenDocuments()
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            MSVisio.Application app = Common.VisioApplication;

            List<MSVisio.Document> openDocs = new List<MSVisio.Document>();

            foreach (MSVisio.Document doc in app.Documents)
            {
                openDocs.Add(doc);
            }

            return openDocs;
        }

        public static void LayerManager()
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}(TODO: Not Implemented)");

            // TODO(CHR): Launch WPF Layer Manager Window
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
