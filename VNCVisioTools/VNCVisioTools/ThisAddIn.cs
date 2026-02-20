using System;
using System.Reflection;

using VNC;

using VNCVisioToolsApplication.Visio;

namespace VNCVisioTools
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Int64 startTicks;
            startTicks = Common.WriteToDebugWindow("ThisAddIn_Startup()", true);
            if (Common.VNCLogging.ApplicationStart) startTicks = Log.APPLICATION_START("Enter", Common.LOG_CATEGORY);            

            GetAssemblyInfo();

            InitializeRibbonUI();

            if (Common.EnableAppEvents)
            {
                if (Common.AppEvents == null)
                {
                    Common.AppEvents = new VNCVisioToolsApplication.Events.VisioAppEvents();
                    Common.AppEvents.VisioApplication = Globals.ThisAddIn.Application;
                }
            }
            else
            {
                Common.AppEvents = null;
            }

            // NOTE(crhodes)
            // These are the events that the AddIn depends on

            Common.AddInApplicationEvents = new VNCVisioToolsApplication.Events.AddInApplicationEvents();
            Common.AddInApplicationEvents.VisioApplication = Globals.ThisAddIn.Application;

            Common.VisioApplication = Globals.ThisAddIn.Application;

            // NOTE(crhodes)
            // Initialize the AddInApplication.
            // This creates the WPF/Prism Environment in a VisioPrismAddInApplication.

            AddInApplication.InitializeApplication();

            Common.WriteToDebugWindow("ThisAddIn_Startup() Exit", startTicks, true);
            if (Common.VNCLogging.ApplicationStart) Log.APPLICATION_START("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            Common.WriteToDebugWindow("ThisAddIn_Shutdown()", true);
            if (Common.VNCLogging.ApplicationEnd) Log.APPLICATION_END("Enter/Exit()", Common.LOG_CATEGORY);            
        }

        private void InitializeRibbonUI()
        {
            Int64 startTicks = 0;
            startTicks = Common.WriteToDebugWindow("InitializeRibbonUI()", true);
            if (Common.VNCLogging.ApplicationInitializeLow) startTicks = Log.APPLICATION_INITIALIZE_LOW("Enter", Common.LOG_CATEGORY);

            Globals.Ribbons.Ribbon.rgDebug.Visible = Common.DeveloperMode = false;
            Globals.Ribbons.Ribbon.rtUILaunchApproaches.Visible = false;

            // NOTE(crhodes)
            // No need to display during normal operation.
            // More for understanding what Visio is doing during development.

            Globals.Ribbons.Ribbon.rcbEnableAppEvents.Checked = Common.EnableAppEvents = false;

            Globals.Ribbons.Ribbon.rcbDisplayEvents.Checked = Common.DisplayEvents = false;
            Globals.Ribbons.Ribbon.rcbDisplayChattyEvents.Checked = Common.DisplayChattyEvents = false;

            Common.WriteToDebugWindow("InitializeRibbonUI()", startTicks, true);
            if (Common.VNCLogging.ApplicationInitializeLow) Log.APPLICATION_INITIALIZE_LOW("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void GetAssemblyInfo()
        {
            Int64 startTicks = 0;
            startTicks = Common.WriteToDebugWindow("GetAssemblyInfo()", true);
            if (Common.VNCLogging.ApplicationInitializeLow) startTicks = Log.APPLICATION_INITIALIZE_LOW("Enter", Common.LOG_CATEGORY);

            // Get Information about ourselves

            var VNCVisioToolsAssembly = Assembly.GetExecutingAssembly();

            if (VNCVisioToolsAssembly != null)
            {
                var VNCVisioToolsAssemblyFileVersionInfo = System.Diagnostics.FileVersionInfo
                    .GetVersionInfo(VNCVisioToolsAssembly.Location);

                Common.InformationVNCVisioTools = Common.GetInformation(
                    VNCVisioToolsAssembly,
                    VNCVisioToolsAssemblyFileVersionInfo);
            }

            Common.WriteToDebugWindow("GetAssemblyInfo() Exit", startTicks, true);
            if (Common.VNCLogging.ApplicationInitializeLow) Log.APPLICATION_INITIALIZE_LOW("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
