﻿using System;

using VNC;

using VNCVisioToolsApplication.Visio;

namespace VNCVisioTools
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Int64 startTicks;
            startTicks = Log.APPLICATION_INITIALIZE("ThisAddIn_Startup()", Common.LOG_CATEGORY);
            startTicks = Common.WriteToDebugWindow("ThisAddIn_Startup()", true);

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
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            Common.WriteToDebugWindow("ThisAddIn_Shutdown()", true);
        }

        void InitializeRibbonUI()
        {
            Globals.Ribbons.Ribbon.rgDebug.Visible = Common.DeveloperMode = false;
            Globals.Ribbons.Ribbon.rtUILaunchApproaches.Visible = false;

            // NOTE(crhodes)
            // No need to display during normal operation.
            // More for understanding what Visio is doing during development.

            Globals.Ribbons.Ribbon.rcbEnableAppEvents.Checked = Common.EnableAppEvents = false;

            Globals.Ribbons.Ribbon.rcbDisplayEvents.Checked = Common.DisplayEvents = false;
            Globals.Ribbons.Ribbon.rcbDisplayChattyEvents.Checked = Common.DisplayChattyEvents = false;
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
