using System;
using System.Threading;
using System.Windows;

using Microsoft.Office.Tools.Ribbon;
using Prism.Unity;
using SupportTools_Excel.Presentation.ViewModels;
using SupportTools_Excel.Presentation.Views;

using VNC;
using VNC.Core.Presentation;
using VNC.WPF.Presentation.Dx.Views;
using VNC.WPF.Presentation.Views;

using ExcelHlp = VNC.AddinHelper.Excel;
using VNCHlp = VNC.AddinHelper;

namespace SupportTools_Excel
{
    public partial class Ribbon
    { 
        #region Event Handlers

        private void btnExplore_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            var frm = new User_Interface.Forms.frmExploreHost();
            frm.Show();

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void btnLoadTPHost_ActiveDirectory_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            var frm = new User_Interface.Forms.frmTaskPaneHost_ActiveDirectory();
            frm.Show();

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void btnTPDevelopment_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (Common.TaskPaneDevelopment == null)
            {
                Common.TaskPaneDevelopment = VNCHlp.TaskPaneUtil.GetTaskPane(
                    () => new User_Interface.Task_Panes.TaskPane_Developer(), "Developer Utilities",
                    Globals.ThisAddIn.CustomTaskPanes, Globals.ThisAddIn.Application.Hwnd.ToString());

                // This works if the minimum size for the control has been set.
                Common.TaskPaneDevelopment.Width = Common.TaskPaneDevelopment.Control.Width;
                Common.TaskPaneDevelopment.Visible = ! Common.TaskPaneDevelopment.Visible;
            }
            else
            {
                Common.TaskPaneDevelopment.Visible = !Common.TaskPaneDevelopment.Visible;
            }

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void ddTheme_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            // TODO(crhodes):
            // This doesn't work.  Try putting it in Support Tools
            DevExpress.Xpf.Core.ThemeManager.ApplicationThemeName = DevExpress.Xpf.Core.Theme.MetropolisLightName;

            DevExpress.Xpf.Core.ThemeManager.ApplicationThemeName = ((RibbonDropDown)sender).SelectedItem.Label;

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void btnActiveDirectory_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (Common.TaskPaneActiveDirectory == null)
            {
                Common.TaskPaneActiveDirectory = VNCHlp.TaskPaneUtil.AddTaskPane(
                    new User_Interface.Task_Panes.TaskPane_ActiveDirectory(), 
                    "Active Directory Utilities", Globals.ThisAddIn.CustomTaskPanes);
                // This works if the minimum size for the control has been set.
                Common.TaskPaneActiveDirectory.Width = Common.TaskPaneActiveDirectory.Control.Width;
            }
            else
            {
                Common.TaskPaneActiveDirectory.Visible = !Common.TaskPaneActiveDirectory.Visible;
            }

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void btnAddInInfo_Click(object sender, RibbonControlEventArgs e)
        {
            DisplayAddInInfo();
        }

        private void btnDebugWindow_Click(object sender, RibbonControlEventArgs e)
        {
            DisplayDebugWindow();
        }

        private void btnDeveloperMode_Click(object sender, RibbonControlEventArgs e)
        {
            ToggleDeveloperMode();
        }

        #region TaskPane Hosts

        private void btnAppUtilities_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (Common.TaskPaneAppUtilities == null)
            {
                Common.TaskPaneAppUtilities = VNCHlp.TaskPaneUtil.GetTaskPane(
                    () => new User_Interface.Task_Panes.TaskPane_ExcelUtil(), "App Utilities",
                    Globals.ThisAddIn.CustomTaskPanes, Globals.ThisAddIn.Application.Hwnd.ToString());

                Common.TaskPaneAppUtilities.Width = Common.TaskPaneAppUtilities.Control.Width;
                Common.TaskPaneAppUtilities.Visible = !Common.TaskPaneAppUtilities.Visible;
            }
            else
            {
                Common.TaskPaneAppUtilities.Visible = !Common.TaskPaneAppUtilities.Visible;
            }

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void btnExcelUtilities_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (Common.TaskPaneUtilities == null)
            {
                Common.TaskPaneUtilities = VNCHlp.TaskPaneUtil.GetTaskPane(
                   () => new User_Interface.Task_Panes.TaskPane_Utilities(), "Excel Utilities",
                   Globals.ThisAddIn.CustomTaskPanes, Globals.ThisAddIn.Application.Hwnd.ToString());

                // This works if the minimum size for the control has been set.
                Common.TaskPaneUtilities.Width = Common.TaskPaneUtilities.Control.Width;
                Common.TaskPaneUtilities.Visible = !Common.TaskPaneUtilities.Visible;
            }
            else
            {
                Common.TaskPaneUtilities.Visible = !Common.TaskPaneUtilities.Visible;
            }

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        //private void btnSharePoint_Click(object sender, RibbonControlEventArgs e)
        //{
        //    Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

        //    Common.TaskPaneSharePoint = VNCHlp.TaskPaneUtil.GetTaskPane(
        //        () => new User_Interface.Task_Panes.TaskPane_SharePoint(), "SharePoint Utilities",
        //        Globals.ThisAddIn.CustomTaskPanes, Globals.ThisAddIn.Application.Hwnd.ToString());

        //    // This works if the minimum size for the control has been set.
        //    Common.TaskPaneSharePoint.Width = Common.TaskPaneSharePoint.Control.Width;
        //    Common.TaskPaneSharePoint.Visible = !Common.TaskPaneSharePoint.Visible;

        //    Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        //}

        private void btnTFS_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Common.TaskPaneTFS = VNCHlp.TaskPaneUtil.GetTaskPane(
                () => new User_Interface.Task_Panes.TaskPane_TFS(), "TFS Utilities",
                Globals.ThisAddIn.CustomTaskPanes, Globals.ThisAddIn.Application.Hwnd.ToString());

            // This works if the minimum size for the control has been set.
            Common.TaskPaneTFS.Width = Common.TaskPaneTFS.Control.Width;
            Common.TaskPaneTFS.Visible = !Common.TaskPaneTFS.Visible;

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region Winform Hosts

        // Load using Winform

        private void btnLoadTFSHost_Click(object sender, RibbonControlEventArgs e)
        {
            long startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            var frm = new User_Interface.Forms.frmTFSHost();
            frm.Show();

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region WPF Hosts

        public static DxThemedWindowHost ad_Host = null;

        private void btnLoadADHost_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            DxThemedWindowHost.DisplayUserControlInHost(ref ad_Host,
            "Active Directory Explorer",
            600, 900,
            //Common.DEFAULT_WINDOW_WIDTH, Common.DEFAULT_WINDOW_HEIGHT,
            ShowWindowMode.Modeless_Show,
            new User_Interface.User_Controls.wucTaskPane_ActiveDirectory());

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost ad_HostMVVM = null;

        private void btnLoadActiveDirectoryHostMVVM_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            DxThemedWindowHost.DisplayUserControlInHost(ref ad_HostMVVM,
            "Active Directory Explorer (MVVM)",
            600, 900,
            //Common.DEFAULT_WINDOW_WIDTH, Common.DEFAULT_WINDOW_HEIGHT,
            ShowWindowMode.Modeless_Show,
            new ActiveDirectoryExplorer.Presentation.Views.ActiveDirectoryExplorer());

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost azdo_Host = null;

        private void btnLoadAZDOHost_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            DxThemedWindowHost.DisplayUserControlInHost(ref azdo_Host,
            "wucTaskPane_TFS",
            600, 900,
            //Common.DEFAULT_WINDOW_WIDTH, Common.DEFAULT_WINDOW_HEIGHT,
            ShowWindowMode.Modeless_Show,
            new User_Interface.User_Controls.wucTaskPane_TFS());

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        //private void btnITRs_Click(object sender, RibbonControlEventArgs e)
        //{
        //    if(Common.TaskPaneITRs == null)
        //    {
        //        Common.TaskPaneITRs = AddinHelper.TaskPaneUtil.AddTaskPane(new User_Interface.Task_Panes.TaskPane_ITRs(), "ITRs", Globals.ThisAddIn.CustomTaskPanes);
        //                        // This works if the minimum size for the control has been set.
        //          Common.TaskPaneITRs.Width = Common.TaskPaneITRs.Control.Width;
        //    }
        //    else
        //    {
        //        Common.TaskPaneITRs.Visible = ! Common.TaskPaneITRs.Visible;
        //    }
        //}

        private void btnLogParser_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (Common.TaskPaneLogParser == null)
            {
                Common.TaskPaneLogParser = VNCHlp.TaskPaneUtil.GetTaskPane(
                      () => new User_Interface.Task_Panes.TaskPane_LogParser(), "Log Parser",
                      Globals.ThisAddIn.CustomTaskPanes, Globals.ThisAddIn.Application.Hwnd.ToString());

                // This works if the minimum size for the control has been set.
                Common.TaskPaneLogParser.Width = Common.TaskPaneLogParser.Control.Width;
                Common.TaskPaneLogParser.Visible = ! Common.TaskPaneLogParser.Visible;
            }
            else
            {
                Common.TaskPaneLogParser.Visible = ! Common.TaskPaneLogParser.Visible;
            }

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        //private void btnLTC_Click(object sender, RibbonControlEventArgs e)
        //{
        //    Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

        //    if (Common.TaskPaneLTC == null)
        //    {
        //        Common.TaskPaneLTC = VNCHlp.TaskPaneUtil.GetTaskPane(
        //            () => new User_Interface.Task_Panes.TaskPane_LTC(), "LTC Utilities",
        //            Globals.ThisAddIn.CustomTaskPanes, Globals.ThisAddIn.Application.Hwnd.ToString());

        //        // This works if the minimum size for the control has been set.
        //        Common.TaskPaneLTC.Width = Common.TaskPaneLTC.Control.Width;
        //        Common.TaskPaneLTC.Visible = ! Common.TaskPaneLTC.Visible;
        //    }
        //    else
        //    {
        //        Common.TaskPaneLTC.Visible = ! Common.TaskPaneLTC.Visible;
        //    }

        //    Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        //}

        //private void btnMTreaty_Click(object sender, RibbonControlEventArgs e)
        //{
        //    Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

        //    if (Common.TaskPaneMTreaty == null)
        //    {
        //        Common.TaskPaneMTreaty = VNCHlp.TaskPaneUtil.GetTaskPane(
        //            () => new User_Interface.Task_Panes.TaskPane_MTreaty(), "MTreaty Utilities",
        //            Globals.ThisAddIn.CustomTaskPanes, Globals.ThisAddIn.Application.Hwnd.ToString());

        //        // This works if the minimum size for the control has been set.
        //        Common.TaskPaneMTreaty.Width = Common.TaskPaneMTreaty.Control.Width;
        //        Common.TaskPaneMTreaty.Visible = !Common.TaskPaneMTreaty.Visible;
        //    }
        //    else
        //    {
        //        Common.TaskPaneMTreaty.Visible = !Common.TaskPaneMTreaty.Visible;
        //    }

        //    Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        //}

        private void btnNetworkTraces_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (Common.TaskPaneNetworkTrace == null)
            {
                Common.TaskPaneNetworkTrace = VNCHlp.TaskPaneUtil.GetTaskPane(
                    () => new User_Interface.Task_Panes.TaskPane_NetworkTrace(), "Network Traces",
                    Globals.ThisAddIn.CustomTaskPanes, Globals.ThisAddIn.Application.Hwnd.ToString());

                // This works if the minimum size for the control has been set.
                Common.TaskPaneNetworkTrace.Width = Common.TaskPaneNetworkTrace.Control.Width;
                Common.TaskPaneNetworkTrace.Visible = !Common.TaskPaneNetworkTrace.Visible;
            }
            else
            {
                Common.TaskPaneNetworkTrace.Visible = !Common.TaskPaneNetworkTrace.Visible;
            }

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        //private void btnSQLSMO_Click(object sender, RibbonControlEventArgs e)
        //{
        //    Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

        //    if (Common.TaskPaneSQLSMO == null)
        //    {
        //        Common.TaskPaneSQLSMO = VNCHlp.TaskPaneUtil.AddTaskPane(new User_Interface.Task_Panes.TaskPane_SQLSMO(), "SQL SMO", Globals.ThisAddIn.CustomTaskPanes);
        //        // This throws an exception
        //        //Globals.ThisAddIn.Application.CommandBars["SQL SMO"].Width = Common.TaskPaneSQLSMO.Width;
        //        foreach (Microsoft.Office.Core.CommandBar bar in Globals.ThisAddIn.Application.CommandBars)
        //        {
        //            string foo = bar.Name;

        //            if (foo == "SQL SMO")
        //            {
        //                // Which is curious as the bar is found!
        //                //Globals.ThisAddIn.Application.CommandBars["SQL SMO"].Width = Common.TaskPaneSQLSMO.Width;
        //            }
        //        }

        //        // This works if the minimum size for the control has been set.
        //        Common.TaskPaneSQLSMO.Width = Common.TaskPaneSQLSMO.Control.Width;
        //    }
        //    else
        //    {
        //        Common.TaskPaneSQLSMO.Visible = !Common.TaskPaneSQLSMO.Visible;
        //    }

        //    Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        //}

        private void btnWatchWindow_Click(object sender, RibbonControlEventArgs e)
        {
            DisplayWatchWindow();
        }

        private void chkDisplayEvents_Click(object sender, RibbonControlEventArgs e)
        {
            Common.DisplayEvents = chkDisplayEvents.Checked;
        }

        private void chkEnableAppEvents_Click(object sender, RibbonControlEventArgs e)
        {
            Common.HasAppEvents = chkEnableAppEvents.Checked;

            if(Common.HasAppEvents)
            {
                if(Common.AppEvents == null)
                {
                    Common.AppEvents = new Events.ExcelAppEvents();
                    Common.AppEvents.ExcelApplication = Globals.ThisAddIn.Application;
                }
            }
            else
            {
                Common.AppEvents = null;
                Common.AppEvents.ExcelApplication = null;
            }
        }

        private void chkEnableTraceLogging_Click(object sender, RibbonControlEventArgs e)
        {
            Common.EnableLogging = chkEnableTraceLogging.Checked;
        }

        private void chkDisplayXlLocationUpdates_Click(object sender, RibbonControlEventArgs e)
        {
            Common.DisplayXlLocationUpdates = chkDisplayXlLocationUpdates.Checked;
        }

        private void chkScreenUpdates_Click(object sender, RibbonControlEventArgs e)
        {
            ExcelHlp.DisplayScreenUpdates = chkScreenUpdates.Checked;
        }

        #endregion

        #region Main Function Routines

        private void DisplayAddInInfo()
        {
            VNCHlp.AddInInfo.DisplayInfo();
        }

        private void DisplayDebugWindow()
        {
            if(VNCHlp.Common.DebugWindow.Visible)
            {
                VNCHlp.Common.DebugWindow.Visible = false;
            }
            else
            {
                VNCHlp.Common.DebugWindow.Visible = true;
            }
        }

        private void DisplayWatchWindow()
        {
            VNCHlp.Common.WatchWindow.Visible = !VNCHlp.Common.WatchWindow.Visible;
        }

        private void ToggleDeveloperMode()
        {
            VNCHlp.Common.DeveloperMode = !VNCHlp.Common.DeveloperMode;
            Globals.Ribbons.Ribbon.grpDebug.Visible = VNCHlp.Common.DeveloperMode;
        }

        #endregion

    }
}
