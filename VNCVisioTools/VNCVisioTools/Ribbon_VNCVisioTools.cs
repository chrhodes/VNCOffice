using System;

using Microsoft.Office.Tools.Ribbon;

using VNCVisioToolsApplication.Actions;

namespace VNCVisioTools
{
    public partial class Ribbon
    {
        #region EventHandlers

        // TODO(crhodes)
        // Should all the calls be wrapped in a try/catch like btnGetDocumentInfo_Click?

        //wrap all calls to Visio_* in try/catch to prevent exceptions from crashing the add-in.
        //Use Common.WriteToDebugWindow(ex.Message, force:true) to handle excep        #region Document Actions Events

        #region Document Action Events

        private void btnAddDefaultLayers_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Visio_Document.AddDefaultLayers();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }
        private void btnAddFooter_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Visio_Document.AddFooter();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnAddHeader_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Visio_Document.AddHeader();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnAddNavigationLinks_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Visio_Document.AddNavigationLinks();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnAddTableOfContents_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Visio_Document.CreateTableOfContents();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnAllPageOff_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                string layerName = cmbLayers.Text;

                if (layerName.Length > 0)
                {
                    Visio_Document.DisplayLayer(layerName, false);
                }
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnAllPageOn_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                string layerName = cmbLayers.Text;

                if (layerName.Length > 0)
                {
                    Visio_Document.DisplayLayer(layerName, true);
                }
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnAutoSizePagesOff_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Visio_Document.AutoSizePagesOff();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnAutoSizePagesOn_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Visio_Document.AutoSizePagesOn();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnDeletePages_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Visio_Document.DeletePages();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnDisplayPageNames_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Visio_Document.DisplayPageNames();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnGetApplicationInfo_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Visio_Application.DisplayInfo();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnGetDocumentInfo_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Visio_Document.DisplayInfo();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnGetStencilInfo_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Visio_Stencil.DisplayInfo();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnLayerManager_Click(object sender, RibbonControlEventArgs e)
        {
            // TODO(crhodes)
            // Do we have a new LayerManager?
        }

        private void btnLoadLayers_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Visio_Application.LayerManager();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnPrintPages_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Visio_Document.PrintPages();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnRemoveLayers_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Visio_Document.RemoveLayers();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnSavePages_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Visio_Document.SavePages();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnSortAllPages_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Visio_Document.SortAllPages();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnSyncPageNames_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Visio_Document.SyncPageNames();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnUpdatePageNameShapes_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Visio_Document.UpdatePageNameShapes();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        #endregion Document Actions Events

        #region Page Actions Events

        private void btnAddNavLinks_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Visio_Page.AddNavigationLinks(Globals.ThisAddIn.Application.ActivePage);
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnAutoSizePageOff_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Visio_Page.AutoSizePageOff();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnAutoSizePageOn_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Visio_Page.AutoSizePageOn();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnGetPageInfo_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Visio_Page.DisplayInfo(Globals.ThisAddIn.Application.ActivePage);
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnPrintPage_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Visio_Page.PrintPage();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnSavePage_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Visio_Page.SavePage(Globals.ThisAddIn.Application.ActivePage);
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnSyncPageNamesPage_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Visio_Page.SyncPageNames();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnUpdatePageNameShapesPage_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Visio_Page.UpdatePageNameShapes(Globals.ThisAddIn.Application.ActivePage);
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        #endregion Page Actions Events

        #region Layer Actions Events

        private void btnAddDefaultLayers_Page_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Visio_Page.AddDefaultLayers();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnLockBackground_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Visio_Page.LockLayer("Background");
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnPageOff_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                string layerName = cmbLayers.Text;

                if (layerName.Length > 0)
                {
                    Visio_Page.DisplayLayer(Globals.ThisAddIn.Application.ActivePage, layerName, false);
                }
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnPageOn_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                string layerName = cmbLayers.Text;

                if (layerName.Length > 0)
                {
                    Visio_Page.DisplayLayer(Globals.ThisAddIn.Application.ActivePage, layerName, true);
                }
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnRemoveLayers_Page_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Visio_Page.RemoveLayers();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnUnlockBackground_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Visio_Page.UnlockLayer("Background");
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        #endregion Layer Actions Events

        #region Visio_Shape Events

        private void btn0PtMargins_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Visio_Shape.SetMargins("0 pt");
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btn1PtMargins_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Visio_Shape.SetMargins("1 pt");
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btn2PtMargins_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Visio_Shape.SetMargins("2 pt");
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnAddColorSupport_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Visio_Shape.AddColorSupportToSelection();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnAddHyperLink_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Visio_Shape.AddHyperlinkToPage_FromShapeText();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnAddIDAndTextSupport_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Visio_Shape.Add_IDandTextSupport_ToSelection();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnAddIDSupport_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Visio_Shape.Add_IDSupport_ToSelection();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnAddIsPageName_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Visio_Shape.Add_User_IsPageName();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnAddTextControl_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Visio_Shape.Add_TextControl_ToSelection();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnGetShapeInfo_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Visio_Shape.GatherInfo();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnMakeLinkableMaster_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Visio_Shape.MakeLinkableMaster();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnMoveToBackgroundLayer_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Visio_Shape.MoveToBackgroundLayer();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        #endregion Visio_Shape Events

        #region CustomUI Events

        private void btnCommandCockpit_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Visio_CustomUI.CommandCockpit();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnLinq2Excel_Click(object sender, RibbonControlEventArgs e)
        {
            //Visio_CustomUI.Linq2Excel();
        }

        private void btnDuplicatePage_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Visio_CustomUI.DuplicatePage();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnRenamePages_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Visio_CustomUI.RenamePages();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnMovePages_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Visio_CustomUI.MovePages();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnEditControlRows_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Visio_CustomUI.EditControlRows();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnEditParagraph_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Visio_CustomUI.EditParagraph();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnEditControlPoints_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Visio_CustomUI.EditControlPoints();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnEditText_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Visio_CustomUI.EditText();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnCustomUI_Car_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Visio_CustomUI.CustomUI_Car();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        #endregion

        #region Document Base Pages Events

        private void btnAddArchitectureBasePages_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Visio_Document.AddArchitectureBasePages();
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
