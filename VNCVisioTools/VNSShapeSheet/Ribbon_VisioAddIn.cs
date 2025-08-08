using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;

using Microsoft.Office.Tools.Ribbon;

using VNCShapeSheet;

using VNCShapeSheetApplication;

using Visio = Microsoft.Office.Interop.Visio;

namespace VNCShapeSheet
{
    public partial class Ribbon
    {
        // TODO(crhodes)
        // This Region can be removed.  It is left here to show functionality from
        // original lessons

        #region EventHandlers

        #region Help Events

        private void btnDisplayAddInInfo_Click(object sender, RibbonControlEventArgs e)
        {
            VNC.VSTOAddIn.AddInInfo.DisplayInfo();
        }

        private void btnToggleDeveloperMode_Click(object sender, RibbonControlEventArgs e)
        {
            VNC.VSTOAddIn.Common.DeveloperMode = !VNC.VSTOAddIn.Common.DeveloperMode;
            Globals.Ribbons.Ribbon.rgDebug.Visible = VNC.VSTOAddIn.Common.DeveloperMode;
        }

        #endregion

        #region Debug Events

        private void btnToggleDeveloperUIMode_Click(object sender, RibbonControlEventArgs e)
        {
            // TODO(crhodes)
            // This is for changing the visibility of MVVM stuff. 
        }

        private void btnDebugWindow_Click(object sender, RibbonControlEventArgs e)
        {
            VNC.VSTOAddIn.Common.DebugWindow.Visible = !VNC.VSTOAddIn.Common.DebugWindow.Visible;
        }

        private void btnWatchWindow_Click(object sender, RibbonControlEventArgs e)
        {
            VNC.VSTOAddIn.Common.WatchWindow.Visible = !VNC.VSTOAddIn.Common.WatchWindow.Visible;
        }

        private void rcbLogToDebugWindow_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show(System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        private void rcbEnableAppEvents_Click(object sender, RibbonControlEventArgs e)
        {
            Common.EnableAppEvents = rcbEnableAppEvents.Checked;

            if (Common.EnableAppEvents)
            {
                if (Common.AppEvents == null)
                {
                    Common.AppEvents = new VNCShapeSheetApplication.Events.VisioAppEvents();
                    Common.AppEvents.VisioApplication = Globals.ThisAddIn.Application;
                }
            }
            else
            {
                Common.AppEvents.VisioApplication = null;
                Common.AppEvents = null;
            }
        }

        private void rcbDisplayEvents_Click(object sender, RibbonControlEventArgs e)
        {
            VNCShapeSheetApplication.Common.DisplayEvents = rcbDisplayEvents.Checked;
        }

        private void rcbDisplayChattyEvents_Click(object sender, RibbonControlEventArgs e)
        {
            Common.DisplayChattyEvents = rcbDisplayChattyEvents.Checked;
        }

        #endregion

        #endregion

    }
}
