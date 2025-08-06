using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;

using Microsoft.Office.Tools.Ribbon;

using VNCVisioTools;

using VNCVisioToolsApplication.Actions;

using Visio = Microsoft.Office.Interop.Visio;

namespace VNCVisioTools
{
    public partial class Ribbon
    {
        #region EventHandlers

        #region Document Actions Events

        //public static DxThemedWindowHost duplicatePage_Host = null;

        private void btnAddDefaultLayers_Click(object sender, RibbonControlEventArgs e)
        {
            //Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Visio_Document.AddDefaultLayers();

            //Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }
        private void btnAddFooter_Click(object sender, RibbonControlEventArgs e)
        {
            //Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Visio_Document.AddFooter();

            //Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void btnAddHeader_Click(object sender, RibbonControlEventArgs e)
        {
            //Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Visio_Document.AddHeader();

            //Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void btnAddNavigationLinks_Click(object sender, RibbonControlEventArgs e)
        {
            //Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Visio_Document.AddNavigationLinks();

            //Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void btnAddTableOfContents_Click(object sender, RibbonControlEventArgs e)
        {
            //Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Visio_Document.CreateTableOfContents();

            //Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void btnAllPageOff_Click(object sender, RibbonControlEventArgs e)
        {
            //Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            string layerName = cmbLayers.Text;

            if (layerName.Length > 0)
            {
                Visio_Document.DisplayLayer(layerName, false);
            }

            //Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void btnAllPageOn_Click(object sender, RibbonControlEventArgs e)
        {
            //Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            string layerName = cmbLayers.Text;

            if (layerName.Length > 0)
            {
                Visio_Document.DisplayLayer(layerName, true);
            }

            //Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void btnAutoSizePagesOff_Click(object sender, RibbonControlEventArgs e)
        {
            //Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Visio_Document.AutoSizePagesOff();

            //Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void btnAutoSizePagesOn_Click(object sender, RibbonControlEventArgs e)
        {
            //Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Visio_Document.AutoSizePagesOn();

            //Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void btnDeletePages_Click(object sender, RibbonControlEventArgs e)
        {
            //Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Visio_Document.DeletePages();

            //Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void btnDisplayPageNames_Click(object sender, RibbonControlEventArgs e)
        {
            //Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Visio_Document.DisplayPageNames();

            //Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void btnGetApplicationInfo_Click(object sender, RibbonControlEventArgs e)
        {
            //Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Visio_Application.DisplayInfo();

            //Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void btnGetDocumentInfo_Click(object sender, RibbonControlEventArgs e)
        {
            //Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);
            try
            {
                Visio_Document.DisplayInfo();
            }
            catch (Exception ex)
            {
                var foo = ex.Message;
            }
            

            //Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void btnGetStencilInfo_Click(object sender, RibbonControlEventArgs e)
        {
            //Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Visio_Stencil.DisplayInfo();

            //Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void btnLayerManager_Click(object sender, RibbonControlEventArgs e)
        {
            // TODO(crhodes)
            // Do we have a new LayerManager?
        }

        private void btnLoadLayers_Click(object sender, RibbonControlEventArgs e)
        {
            //Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Visio_Application.LayerManager();

            //Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void btnPrintPages_Click(object sender, RibbonControlEventArgs e)
        {
            //Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Visio_Document.PrintPages();

            //Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void btnRemoveLayers_Click(object sender, RibbonControlEventArgs e)
        {
            //Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Visio_Document.RemoveLayers();

            //Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void btnSavePages_Click(object sender, RibbonControlEventArgs e)
        {
            //Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Visio_Document.SavePages();

            //Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void btnSortAllPages_Click(object sender, RibbonControlEventArgs e)
        {
            //Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Visio_Document.SortAllPages();

            //Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void btnSyncPageNames_Click(object sender, RibbonControlEventArgs e)
        {
            //Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Visio_Document.SyncPageNames();

            //Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void btnUpdatePageNameShapes_Click(object sender, RibbonControlEventArgs e)
        {
            //Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Visio_Document.UpdatePageNameShapes();

            //Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion Document Actions Events

        #region Page Actions Events

        private void btnAddNavLinks_Click(object sender, RibbonControlEventArgs e)
        {
            //Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Visio_Page.AddNavigationLinks(Globals.ThisAddIn.Application.ActivePage);

            //Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void btnAutoSizePageOff_Click(object sender, RibbonControlEventArgs e)
        {
            //Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Visio_Page.AutoSizePageOff();

            //Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void btnAutoSizePageOn_Click(object sender, RibbonControlEventArgs e)
        {
            //Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Visio_Page.AutoSizePageOn();

            //Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void btnGetPageInfo_Click(object sender, RibbonControlEventArgs e)
        {
            //Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Visio_Page.DisplayInfo(Globals.ThisAddIn.Application.ActivePage);

            //Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void btnPrintPage_Click(object sender, RibbonControlEventArgs e)
        {
            //Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Visio_Page.PrintPage();

            //Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void btnSavePage_Click(object sender, RibbonControlEventArgs e)
        {
            //Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Visio_Page.SavePage(Globals.ThisAddIn.Application.ActivePage);

            //Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void btnSyncPageNamesPage_Click(object sender, RibbonControlEventArgs e)
        {
            //Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Visio_Page.SyncPageNames();

            //Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void btnUpdatePageNameShapesPage_Click(object sender, RibbonControlEventArgs e)
        {
            //Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Visio_Page.UpdatePageNameShapes(Globals.ThisAddIn.Application.ActivePage);

            //Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion Page Actions Events

        #region Layer Actions Events

        private void btnAddDefaultLayers_Page_Click(object sender, RibbonControlEventArgs e)
        {
            //Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Visio_Page.AddDefaultLayers();

            //Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void btnLockBackground_Click(object sender, RibbonControlEventArgs e)
        {
            //Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Visio_Page.LockLayer("Background");

            //Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void btnPageOff_Click(object sender, RibbonControlEventArgs e)
        {
            //Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            string layerName = cmbLayers.Text;

            if (layerName.Length > 0)
            {
                Visio_Page.DisplayLayer(Globals.ThisAddIn.Application.ActivePage, layerName, false);
            }

            //Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void btnPageOn_Click(object sender, RibbonControlEventArgs e)
        {
            //Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            string layerName = cmbLayers.Text;

            if (layerName.Length > 0)
            {
                Visio_Page.DisplayLayer(Globals.ThisAddIn.Application.ActivePage, layerName, true);
            }

            //Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void btnRemoveLayers_Page_Click(object sender, RibbonControlEventArgs e)
        {
            //Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Visio_Page.RemoveLayers();

            //Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void btnUnlockBackground_Click(object sender, RibbonControlEventArgs e)
        {
            //Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Visio_Page.UnlockLayer("Background");

            //Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion Layer Actions Events

        #region Visio_Shape Events

        private void btn0PtMargins_Click(object sender, RibbonControlEventArgs e)
        {
            //Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Visio_Shape.SetMargins("0 pt");

            //Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void btn1PtMargins_Click(object sender, RibbonControlEventArgs e)
        {
            //Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Visio_Shape.SetMargins("1 pt");

            //Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void btn2PtMargins_Click(object sender, RibbonControlEventArgs e)
        {
            //Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Visio_Shape.SetMargins("2 pt");

            //Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void btnAddColorSupport_Click(object sender, RibbonControlEventArgs e)
        {
            //Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Visio_Shape.AddColorSupportToSelection();

            //Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void btnAddHyperLink_Click(object sender, RibbonControlEventArgs e)
        {
            //Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Visio_Shape.AddHyperlinkToPage_FromShapeText();

            //Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void btnAddIDAndTextSupport_Click(object sender, RibbonControlEventArgs e)
        {
            //Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Visio_Shape.Add_IDandTextSupport_ToSelection();

            //Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void btnAddIDSupport_Click(object sender, RibbonControlEventArgs e)
        {
            //Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Visio_Shape.Add_IDSupport_ToSelection();

            //Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void btnAddIsPageName_Click(object sender, RibbonControlEventArgs e)
        {
            //Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Visio_Shape.Add_User_IsPageName();

            //Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void btnAddTextControl_Click(object sender, RibbonControlEventArgs e)
        {
            //Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Visio_Shape.Add_TextControl_ToSelection();

            //Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void btnGetShapeInfo_Click(object sender, RibbonControlEventArgs e)
        {
            //Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Visio_Shape.GatherInfo();
        }

        private void btnMakeLinkableMaster_Click(object sender, RibbonControlEventArgs e)
        {
            //Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Visio_Shape.MakeLinkableMaster();

            //Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void btnMoveToBackgroundLayer_Click(object sender, RibbonControlEventArgs e)
        {

            Visio_Shape.MoveToBackgroundLayer();

            //Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion Visio_Shape Events

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

        private void btnDebugWindow_Click(object sender, RibbonControlEventArgs e)
        {
            VNC.VSTOAddIn.Common.DebugWindow.Visible = !VNC.VSTOAddIn.Common.DebugWindow.Visible;
        }

        private void btnWatchWindow_Click(object sender, RibbonControlEventArgs e)
        {
            VNC.VSTOAddIn.Common.WatchWindow.Visible = !VNC.VSTOAddIn.Common.WatchWindow.Visible;
        }

        private void rcbDisplayChattyEvents_Click(object sender, RibbonControlEventArgs e)
        {
            Common.DisplayChattyEvents = rcbDisplayChattyEvents.Checked;
        }

        private void rcbDisplayEvents_Click(object sender, RibbonControlEventArgs e)
        {
            VNCVisioToolsApplication.Common.DisplayEvents = rcbDisplayEvents.Checked;
        }

        private void rcbEnableAppEvents_Click(object sender, RibbonControlEventArgs e)
        {
            Common.EnableAppEvents = rcbEnableAppEvents.Checked;

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
                Common.AppEvents.VisioApplication = null;
                Common.AppEvents = null;
            }
        }

        //private void rcbLogToDebugWindow_Click(object sender, RibbonControlEventArgs e)
        //{
        //    MessageBox.Show(System.Reflection.MethodInfo.GetCurrentMethod().Name);
        //}

        private void rcbToggleDeveloperUIMode_Click(object sender, RibbonControlEventArgs e)
        {
            
            // TODO(crhodes)
            // This is for changing the visibility of MVVM stuff. 
        }

        #endregion

        #endregion

    }
}
