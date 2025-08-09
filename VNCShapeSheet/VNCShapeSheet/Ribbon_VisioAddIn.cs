using System;
using System.Windows.Forms;

using Microsoft.Office.Tools.Ribbon;
using VNCShapeSheetApplication.Actions;
namespace VNCShapeSheet
{
    public partial class Ribbon
    {
        #region EventHandlers

        #region Document (Object) Events

        private void btnDocumentProperties_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShapeSheetUI.DocumentProperties();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        #endregion

        #region Document (Rows) Events

        private void btnDocumentHyperlinks_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShapeSheetUI.DocumentHyperlinks();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnDocumentScratch_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShapeSheetUI.DocumentScratch();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnDocumentShapeData_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShapeSheetUI.DocumentShapeData();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnDocumentUserDefinedCells_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShapeSheetUI.DocumentUserDefinedCells();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        #endregion

        #region Page (Object) Events

        private void btnPageLayout_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShapeSheetUI.PageLayout();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnPageProperties_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShapeSheetUI.PageProperties();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnPrintProperties_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShapeSheetUI.PrintProperties();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnRulerAndGrid_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShapeSheetUI.RulerAndGrid();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnPageThemeProperties_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShapeSheetUI.PageThemeProperties();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        #endregion

        #region Page (Rows) Events

        private void btnLayers_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShapeSheetUI.Layers();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnPageActions_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShapeSheetUI.ActionsPage();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnPageActionTags_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShapeSheetUI.ActionTagsPage();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnPageHyperlinks_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShapeSheetUI.PageHyperlinks();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnPageScratch_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShapeSheetUI.PageScratch();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnPageShapeData_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShapeSheetUI.PageShapeData();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnPageUserDefinedCells_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShapeSheetUI.PageUserDefinedCells();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        #endregion

        #region Shape (Object) Events

        private void btn1DEndpoints_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShapeSheetUI.OneDEndpoints();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btn3DRotationProperties_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShapeSheetUI.ThreeDRotationProperties();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnAdditionalEffectProperties_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShapeSheetUI.AdditionalEffectProperties();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnAlignment_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShapeSheetUI.Alignment();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnBevelProperties_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShapeSheetUI.BevelProperties();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnChangeShapeBehavior_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShapeSheetUI.ChangeShapeBehavior();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnEvents_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShapeSheetUI.Events();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnFillFormat_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShapeSheetUI.FillFormat();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnGlueInfo_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShapeSheetUI.GlueInfo();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnGradientProperties_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShapeSheetUI.GradientProperties();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnGroupProperties_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShapeSheetUI.GroupProperties();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnImageProperties_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShapeSheetUI.ImageProperties();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnLayerMembership_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShapeSheetUI.LayerMembership();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnLineFormat_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShapeSheetUI.LineFormat();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnMiscelleaneous_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShapeSheetUI.Miscelleaneous();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnProtection_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShapeSheetUI.Protection();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnQuickStyle_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShapeSheetUI.QuickStyle();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnShapeLayout_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShapeSheetUI.ShapeLayout();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnShapeTransform_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShapeSheetUI.ShapeTransform();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnTextBlockFormat_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShapeSheetUI.TextBlockFormat();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnTextTransform_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShapeSheetUI.TextTransform();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnThemeProperties_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShapeSheetUI.ShapeThemeProperties();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        #endregion

        #region Shape (Rows) Events

        private void btnActions_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShapeSheetUI.ActionsShape();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnActionTags_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShapeSheetUI.ActionTagsShape();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnCharacter_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShapeSheetUI.Character();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnConnectionPoints_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShapeSheetUI.ConnectionPoints();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnControls_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShapeSheetUI.Controls();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnGeometry_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShapeSheetUI.Geometry();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnGradientStops_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShapeSheetUI.GradientStops();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnParagraph_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShapeSheetUI.Paragraph();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnShapeHyperlinks_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShapeSheetUI.ShapeHyperlinks();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnShapeScratch_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShapeSheetUI.ShapeScratch();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnShapeShapeData_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShapeSheetUI.ShapeShapeData();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnTabs_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShapeSheetUI.Tabs();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnShapeUserDefinedCells_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShapeSheetUI.ShapeUserDefinedCells();
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
