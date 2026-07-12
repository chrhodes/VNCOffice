using System;

using Microsoft.Office.Tools.Ribbon;

using VNCExcelToolsApplication.Actions;

using VNC;
using DevExpress.XtraRichEdit.Commands;

namespace VNCExcelTools
{
    public partial class Ribbon
    {
        #region EventHandlers

        // NOTE(crhodes)
        // Wrap all calls to Excel_* in try/catch to prevent exceptions from crashing the add-in.
        // Use Common.WriteToDebugWindow(ex.Message, force:true) to handle exceptions
        // Should we also use VNC.Log??

        #region Document Action Events

        private void btnAddFooterAllPages_Click(object sender, RibbonControlEventArgs e)
        {
            if (Common.VNCLogLevel.EventHandler) Log.EVENT_HANDLER("", Common.LOG_CATEGORY);

            try
            {
                Excel_PageFormatting.AddFooterAllPages();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnAddHeaderAllPages_Click(object sender, RibbonControlEventArgs e)
        {
            if (Common.VNCLogLevel.EventHandler) Log.EVENT_HANDLER("", Common.LOG_CATEGORY);

            try
            {
                Excel_PageFormatting.AddHeaderAllPages();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnAddTableOfContents_Click(object sender, RibbonControlEventArgs e)
        {
            if (Common.VNCLogLevel.EventHandler) Log.EVENT_HANDLER("", Common.LOG_CATEGORY);

            try
            {
                Excel_Workbook.AddTableOfContents();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnLockAllWorksheets_Click(object sender, RibbonControlEventArgs e)
        {
            if (Common.VNCLogLevel.EventHandler) Log.EVENT_HANDLER("", Common.LOG_CATEGORY);

            try
            {
                Excel_Workbook.LockAllWorksheets();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnUnlockAllWorksheets_Click(object sender, RibbonControlEventArgs e)
        {
            if (Common.VNCLogLevel.EventHandler) Log.EVENT_HANDLER("", Common.LOG_CATEGORY);

            try
            {
                Excel_Workbook.UnlockAllWorksheets();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        #endregion Document Actions Events

        #region Workbook Action Events

        private void btnUnlockWorksheet_Click(object sender, RibbonControlEventArgs e)
        {
            if (Common.VNCLogLevel.EventHandler) Log.EVENT_HANDLER("", Common.LOG_CATEGORY);

            try
            {
                Excel_Worksheet.UnlockWorksheet();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnLockWorksheet_Click(object sender, RibbonControlEventArgs e)
        {
            if (Common.VNCLogLevel.EventHandler) Log.EVENT_HANDLER("", Common.LOG_CATEGORY);

            try
            {
                Excel_Worksheet.LockWorksheet();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        #endregion

        #region Page Events

        private void btnAllLandscape_Click(object sender, RibbonControlEventArgs e)
        {
            if (Common.VNCLogLevel.EventHandler) Log.EVENT_HANDLER("", Common.LOG_CATEGORY);

            try
            {
                Excel_PageFormatting.AllLandscape();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnLandscape_Click(object sender, RibbonControlEventArgs e)
        {
            if (Common.VNCLogLevel.EventHandler) Log.EVENT_HANDLER("", Common.LOG_CATEGORY);

            try
            {
                Excel_PageFormatting.Landscape();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnAllPortrait_Click(object sender, RibbonControlEventArgs e)
        {
            if (Common.VNCLogLevel.EventHandler) Log.EVENT_HANDLER("", Common.LOG_CATEGORY);

            try
            {
                Excel_PageFormatting.AllPortrait();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnPortrait_Click(object sender, RibbonControlEventArgs e)
        {
            if (Common.VNCLogLevel.EventHandler) Log.EVENT_HANDLER("", Common.LOG_CATEGORY);

            try
            {
                Excel_PageFormatting.Portrait();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        #endregion

        #region Utilities Events

        private void btnFolderMap_Click(object sender, RibbonControlEventArgs e)
        {
            if (Common.VNCLogLevel.EventHandler) Log.EVENT_HANDLER("", Common.LOG_CATEGORY);

            try
            {
                Excel_CustomUI.FolderMap();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        #endregion

        #region Help Events

        private void btnDisplayAddInInfo_Click(object sender, RibbonControlEventArgs e)
        {
            if (Common.VNCLogLevel.EventHandler) Log.EVENT_HANDLER("", Common.LOG_CATEGORY);
            // NOTE(crhodes)
            // This is for the old approach
            //VNC.VSTOAddIn.AddInInfo.DisplayInfo();

            // NOTE(crhodes)
            // This is for the new approach

            try
            {
                Excel_Application.DisplayAddInInfo();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnToggleDeveloperMode_Click(object sender, RibbonControlEventArgs e)
        {
            if (Common.VNCLogLevel.EventHandler) Log.EVENT_HANDLER("", Common.LOG_CATEGORY);
            VNC.VSTOAddIn.Common.DeveloperMode = !VNC.VSTOAddIn.Common.DeveloperMode;
            Globals.Ribbons.Ribbon.rgDebug.Visible = VNC.VSTOAddIn.Common.DeveloperMode;
        }

        #endregion

        #region Debug Events

        private void btnDebugWindow_Click(object sender, RibbonControlEventArgs e)
        {
            if (Common.VNCLogLevel.EventHandler) Log.EVENT_HANDLER("", Common.LOG_CATEGORY);
            VNC.VSTOAddIn.Common.DebugWindow.Visible = !VNC.VSTOAddIn.Common.DebugWindow.Visible;
        }

        private void btnWatchWindow_Click(object sender, RibbonControlEventArgs e)
        {
            if (Common.VNCLogLevel.EventHandler) Log.EVENT_HANDLER("", Common.LOG_CATEGORY);
            VNC.VSTOAddIn.Common.WatchWindow.Visible = !VNC.VSTOAddIn.Common.WatchWindow.Visible;
        }

        private void rcbDisplayChattyEvents_Click(object sender, RibbonControlEventArgs e)
        {
            if (Common.VNCLogLevel.EventHandler) Log.EVENT_HANDLER("", Common.LOG_CATEGORY);
            Common.DisplayChattyEvents = rcbDisplayChattyEvents.Checked;
        }

        private void rcbDisplayEvents_Click(object sender, RibbonControlEventArgs e)
        {
            if (Common.VNCLogLevel.EventHandler) Log.EVENT_HANDLER("", Common.LOG_CATEGORY);
            VNCExcelToolsApplication.Common.DisplayEvents = rcbDisplayEvents.Checked;
        }

        private void rcbEnableAppEvents_Click(object sender, RibbonControlEventArgs e)
        {
            if (Common.VNCLogLevel.EventHandler) Log.EVENT_HANDLER("", Common.LOG_CATEGORY);
            Common.EnableAppEvents = rcbEnableAppEvents.Checked;

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
                Common.AppEvents.ExcelApplication = null;
                Common.AppEvents = null;
            }
        }

        //private void rcbLogToDebugWindow_Click(object sender, RibbonControlEventArgs e)
        //{
        //    MessageBox.Show(System.Reflection.MethodInfo.GetCurrentMethod().Name);
        //}

        private void rcbToggleDeveloperUIMode_Click(object sender, RibbonControlEventArgs e)
        {
            if (Common.VNCLogLevel.EventHandler) Log.EVENT_HANDLER("", Common.LOG_CATEGORY);
            // TODO(crhodes)
            // This is for changing the visibility of MVVM stuff.

            try
            {
                Excel_Application.DeveloperModeUI(rcbToggleDeveloperUIMode.Checked);
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnTestExcelLogging_Click(object sender, RibbonControlEventArgs e)
        {
            if (Common.VNCLogLevel.EventHandler) Log.EVENT_HANDLER("", Common.LOG_CATEGORY);

            try
            {
                Excel_CustomUI.TestExcelLogging();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnLoggingConfiguration_Click(object sender, RibbonControlEventArgs e)
        {
            if (Common.VNCLogLevel.EventHandler) Log.EVENT_HANDLER("", Common.LOG_CATEGORY);

            try
            {
                Excel_CustomUI.LoggingConfiguration();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void rcbUILaunchApproaches_Click(object sender, RibbonControlEventArgs e)
        {
            if (Common.VNCLogLevel.EventHandler) Log.EVENT_HANDLER("", Common.LOG_CATEGORY);

            Globals.Ribbons.Ribbon.rtUILaunchApproaches.Visible = Globals.Ribbons.Ribbon.rcbUILaunchApproaches.Checked;
        }

        #endregion

        #endregion
    }
}
