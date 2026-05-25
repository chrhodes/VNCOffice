using System;
using System.Reflection;

using VNC;

using VNCExcelToolsApplication.Excel;

namespace VNCExcelTools
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Int64 startTicks;
            startTicks = Common.WriteToDebugWindow("Enter", true);
            if (Common.VNCLogLevel.ApplicationStart) startTicks = Log.APPLICATION_START("Enter", Common.LOG_CATEGORY);

            GetAssemblyInfo();

            InitializeRibbonUI();

            if (Common.EnableAppEvents)
            {
                if (Common.AppEvents == null)
                {
                    Common.AppEvents = new VNCExcelToolsApplication.Events.ExcelAppEvents();
                    Common.AppEvents.ExcelApplication = Globals.ThisAddIn.Application;
                }
            }
            else
            {
                Common.AppEvents = null;
            }

            // NOTE(crhodes)
            // These are the events that the AddIn depends on

            Common.AddInApplicationEvents = new VNCExcelToolsApplication.Events.AddInApplicationEvents();
            Common.AddInApplicationEvents.ExcelApplication = Globals.ThisAddIn.Application;

            Common.ExcelApplication = Globals.ThisAddIn.Application;

            // NOTE(crhodes)
            // Added this so it is available to the VNC.VSTOAddIn.Excel.  Perhaps it should be in Common.

            //VNC.VSTOAddIn.Excel.Domain.Excel.ExcelApplication = Globals.ThisAddIn.Application;

            // NOTE(crhodes)
            // Initialize the AddInApplication.
            // This creates the WPF/Prism Environment in a ExcelPrismAddInApplication.

            AddInApplication.InitializeApplication();

            Common.WriteToDebugWindow("Exit", startTicks, true);
            if (Common.VNCLogLevel.ApplicationStart) Log.APPLICATION_START("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            Common.WriteToDebugWindow("Enter/Exit", true);
            if (Common.VNCLogLevel.ApplicationEnd) Log.APPLICATION_END("Enter/Exit()", Common.LOG_CATEGORY);
        }

        private void InitializeRibbonUI()
        {
            Int64 startTicks = 0;
            startTicks = Common.WriteToDebugWindow("", true);
            if (Common.VNCLogLevel.ApplicationInitializeLow) startTicks = Log.APPLICATION_INITIALIZE_LOW("Enter", Common.LOG_CATEGORY);

            Globals.Ribbons.Ribbon.rgDebug.Visible = Common.DeveloperMode = false;
            Globals.Ribbons.Ribbon.rtUILaunchApproaches.Visible = false;

            // NOTE(crhodes)
            // No need to display during normal operation.
            // More for understanding what Excel is doing during development.

            Globals.Ribbons.Ribbon.rcbEnableAppEvents.Checked = Common.EnableAppEvents = false;

            Globals.Ribbons.Ribbon.rcbDisplayEvents.Checked = Common.DisplayEvents = false;
            Globals.Ribbons.Ribbon.rcbDisplayChattyEvents.Checked = Common.DisplayChattyEvents = false;

            Common.WriteToDebugWindow("Exit", startTicks, true);
            if (Common.VNCLogLevel.ApplicationInitializeLow) Log.APPLICATION_INITIALIZE_LOW("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void GetAssemblyInfo()
        {
            Int64 startTicks = 0;
            startTicks = Common.WriteToDebugWindow("", true);
            if (Common.VNCLogLevel.ApplicationInitializeLow) startTicks = Log.APPLICATION_INITIALIZE_LOW("Enter", Common.LOG_CATEGORY);

            // Get Information about ourselves

            var VNCExcelToolsAssembly = Assembly.GetExecutingAssembly();

            if (VNCExcelToolsAssembly != null)
            {
                var VNCExcelToolsAssemblyFileVersionInfo = System.Diagnostics.FileVersionInfo
                    .GetVersionInfo(VNCExcelToolsAssembly.Location);

                Common.InformationVNCExcelTools = Common.GetInformation(
                    VNCExcelToolsAssembly,
                    VNCExcelToolsAssemblyFileVersionInfo);
            }

            Common.WriteToDebugWindow("Exit", startTicks, true);
            if (Common.VNCLogLevel.ApplicationInitializeLow) Log.APPLICATION_INITIALIZE_LOW("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
