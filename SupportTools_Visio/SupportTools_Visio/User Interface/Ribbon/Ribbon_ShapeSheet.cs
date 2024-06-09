using System;

using Microsoft.Office.Tools.Ribbon;

using SupportTools_Visio.Domain;

using VNC;
using VNC.Core.Presentation;
using VNC.WPF.Presentation.Dx.Views;

using VNCVisioAddIn = VNC.Visio.VSTOAddIn;

namespace SupportTools_Visio
{
    public partial class Ribbon
    {
        #region Event Handlers

        #region ShapeSheet UI Events Document Related

        public static DxThemedWindowHost _documentPropertiesHost = null;

        private void btnDocumentProperties_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            //DxThemedWindowHost.DisplayUserControlInHost(ref _documentPropertiesHost,
            //    "Document Properties",
            //    600, 450,
            //    //Common.DEFAULT_WINDOW_WIDTH, Common.DEFAULT_WINDOW_HEIGHT,
            //    DxThemedWindowHost.ShowWindowMode.Modeless,
            //    new Presentation.Views.DocumentShapeSheetSection(
            //        new Presentation.ViewModels.DocumentPropertiesViewModel(),
            //        new Presentation.Views.DocumentProperties()));
            DxThemedWindowHost.DisplayUserControlInHost(ref _pageActionsHost,
                "Document Properties",
                600, 600,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetSection(
                    new Presentation.ViewModels.ObjectViewModel<VNCVisioAddIn.Domain.DocumentPropertiesRow, Presentation.ModelWrappers.DocumentPropertiesWrapper>(
                        "Update Properties",
                        VNC.Visio.VSTOAddIn.Domain.DocumentPropertiesRow.GetRow,
                        ShapeType.Document),
                    new Presentation.Views.DocumentProperties()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region ShapeSheet UI Events Page Related

        public static DxThemedWindowHost _pagePageLayoutHost = null;

        private void btnPageLayout_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            DxThemedWindowHost.DisplayUserControlInHost(ref _pagePagePropertiesHost,
                "Page PageLayout",
                600, 600,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetSection(
                    new Presentation.ViewModels.ObjectViewModel<VNCVisioAddIn.Domain.PageLayoutRow, Presentation.ModelWrappers.PageLayoutWrapper>(
                        "Update Page Properties",
                        VNC.Visio.VSTOAddIn.Domain.PageLayoutRow.GetRow,
                        ShapeType.Page),
                    new Presentation.Views.PageLayout()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost _pagePagePropertiesHost = null;

        private void btnPageProperties_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            //DxThemedWindowHost.DisplayUserControlInHost(ref _pagePagePropertiesHost,
            //    "Page Properties",
            //    600, 450,
            //    //Common.DEFAULT_WINDOW_WIDTH, Common.DEFAULT_WINDOW_HEIGHT,
            //    DxThemedWindowHost.ShowWindowMode.Modeless,
            //    new Presentation.Views.PageShapeSheetSection(
            //        new Presentation.ViewModels.PagePropertiesViewModel(),
            //        new Presentation.Views.PageProperties()));
            DxThemedWindowHost.DisplayUserControlInHost(ref _pagePagePropertiesHost,
                "Page Properties",
                600, 600,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetSection(
                    new Presentation.ViewModels.ObjectViewModel<VNCVisioAddIn.Domain.PagePropertiesRow, Presentation.ModelWrappers.PagePropertiesWrapper>(
                        "Update Page Properties",
                        VNC.Visio.VSTOAddIn.Domain.PagePropertiesRow.Get_PageProperties,
                        ShapeType.Page),
                    new Presentation.Views.PageProperties()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost _pagePrintPropertiesHost = null;

        private void btnPrintProperties_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            DxThemedWindowHost.DisplayUserControlInHost(ref _pagePrintPropertiesHost,
                "Page PrintProperties",
                600, 600,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetSection(
                    new Presentation.ViewModels.ObjectViewModel<VNCVisioAddIn.Domain.PrintPropertiesRow, Presentation.ModelWrappers.PrintPropertiesWrapper>(
                        "Update PrintProperties",
                        VNC.Visio.VSTOAddIn.Domain.PrintPropertiesRow.Get_PrintProperties,
                        ShapeType.Page),
                    new Presentation.Views.PrintProperties()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost _pageRulerAndGridsHost = null;

        private void btnRulerAndGrid_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            DxThemedWindowHost.DisplayUserControlInHost(ref _pageRulerAndGridsHost,
                "Page Ruler & Grid",
                600, 600,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetSection(
                    new Presentation.ViewModels.ObjectViewModel<VNCVisioAddIn.Domain.RulerAndGridRow, Presentation.ModelWrappers.RulerAndGridWrapper>(
                        "Update Ruler & Grid",
                        VNC.Visio.VSTOAddIn.Domain.RulerAndGridRow.Get_RulerAndGrid,
                        ShapeType.Page),
                    new Presentation.Views.RulerAndGrid()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost _pageThemePropertiesHost = null;

        private void btnPageThemeProperties_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            DxThemedWindowHost.DisplayUserControlInHost(ref _pageThemePropertiesHost,
                "Page ThemeProperties",
                600, 800,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetSection(
                    new Presentation.ViewModels.ObjectViewModel<VNCVisioAddIn.Domain.ThemePropertiesRow, Presentation.ModelWrappers.ThemePropertiesWrapper>(
                        "Update ThemeProperties",
                        VNC.Visio.VSTOAddIn.Domain.ThemePropertiesRow.Get_ThemeProperties,
                        ShapeType.Page),
                    new Presentation.Views.ThemeProperties()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region ShapeSheet UI Events Shape Related

        public static DxThemedWindowHost _shapeOneDEndpointsHost = null;

        private void btn1DEndpoints_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            DxThemedWindowHost.DisplayUserControlInHost(ref _shapeOneDEndpointsHost,
                "Shape 1-D Endpoints",
                600, 600,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetSection(
                    new Presentation.ViewModels.ObjectViewModel<VNCVisioAddIn.Domain.OneDEndPointsRow, Presentation.ModelWrappers.OneDEndPointsWrapper>(
                        "Update 1-D Endpoints",
                        VNC.Visio.VSTOAddIn.Domain.OneDEndPointsRow.Get_OneDEndPoints,
                        ShapeType.Shape),
                    new Presentation.Views.OneDEndPoints()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost _shapeThreeDRotationPropertiesHost = null;

        private void btn3DRotationProperties_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            DxThemedWindowHost.DisplayUserControlInHost(ref _shapeThreeDRotationPropertiesHost,
                "3D Rotation Properties",
                600, 450,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetSection(
                    new Presentation.ViewModels.ObjectViewModel<VNCVisioAddIn.Domain.ThreeDRotationPropertiesRow, Presentation.ModelWrappers.ThreeDRotationPropertiesWrapper>(
                        "Update 3-D Rotation Properties",
                        VNC.Visio.VSTOAddIn.Domain.ThreeDRotationPropertiesRow.GetRow,
                        ShapeType.Shape),
                    new Presentation.Views.OneDEndPoints()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost _shapeAdditionalEffectPropertiesHost = null;

        private void btnAdditionalEffectProperties_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            DxThemedWindowHost.DisplayUserControlInHost(ref _shapeAdditionalEffectPropertiesHost,
                "Additional Effect Properties",
                600, 450,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetSection(
                    new Presentation.ViewModels.ObjectViewModel<VNCVisioAddIn.Domain.AdditionalEffectPropertiesRow, Presentation.ModelWrappers.AdditionalEffectPropertiesWrapper>(
                        "Update Additional Effect Properties",
                        VNC.Visio.VSTOAddIn.Domain.AdditionalEffectPropertiesRow.Get_AdditionalEffectProperties,
                        ShapeType.Shape),
                    new Presentation.Views.AdditionalEffectProperties()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost _shapeBevelPropertiesHost = null;

        private void btnBevelProperties_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            DxThemedWindowHost.DisplayUserControlInHost(ref _shapeBevelPropertiesHost,
            "Bevel Properties",
            600, 450,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetSection(
                    new Presentation.ViewModels.ObjectViewModel<VNCVisioAddIn.Domain.BevelPropertiesRow, Presentation.ModelWrappers.BevelPropertiesWrapper>(
                        "Update Bevel Properties",
                        VNC.Visio.VSTOAddIn.Domain.BevelPropertiesRow.Get_BevelProperties,
                        ShapeType.Shape),
                    new Presentation.Views.BevelProperties()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost _shapeChangeShapeBehaviorHost = null;

        private void btnChangeShapeBehavior_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            DxThemedWindowHost.DisplayUserControlInHost(ref _shapeChangeShapeBehaviorHost,
            "Change Shape Behavior",
            600, 450,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetSection(
                    new Presentation.ViewModels.ObjectViewModel<VNCVisioAddIn.Domain.ChangeShapeBehaviorRow, Presentation.ModelWrappers.ChangeShapeBehaviorWrapper>(
                        "Update ChangeShapeBehavior",
                        VNC.Visio.VSTOAddIn.Domain.ChangeShapeBehaviorRow.Get_ChangeShapeBehavior,
                        ShapeType.Shape),
                    new Presentation.Views.ChangeShapeBehavior()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost _shapeEventsHost = null;

        private void btnEvents_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            DxThemedWindowHost.DisplayUserControlInHost(ref _shapeEventsHost,
            "Events",
            600, 450,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetSection(
                    new Presentation.ViewModels.ObjectViewModel<VNCVisioAddIn.Domain.EventsRow, Presentation.ModelWrappers.EventsWrapper>(
                        "Update Events",
                        VNC.Visio.VSTOAddIn.Domain.EventsRow.Get_Events,
                        ShapeType.Shape),
                    new Presentation.Views.Events()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost _shapeFillFormatHost = null;

        private void btnFillFormat_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            DxThemedWindowHost.DisplayUserControlInHost(ref _shapeFillFormatHost,
            "Fill Format",
            600, 450,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetSection(
                    new Presentation.ViewModels.ObjectViewModel<VNCVisioAddIn.Domain.FillFormatRow, Presentation.ModelWrappers.FillFormatWrapper>(
                        "Update Fill Format",
                        VNC.Visio.VSTOAddIn.Domain.FillFormatRow.GetRow,
                        ShapeType.Shape),
                    new Presentation.Views.FillFormat()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost _shapeGlueInfoHost = null;

        private void btnGlueInfo_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            DxThemedWindowHost.DisplayUserControlInHost(ref _shapeGlueInfoHost,
            "Glue Info",
            600, 450,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetSection(
                    new Presentation.ViewModels.ObjectViewModel<VNCVisioAddIn.Domain.GlueInfoRow, Presentation.ModelWrappers.GlueInfoWrapper>(
                        "Update Glue Info",
                        VNC.Visio.VSTOAddIn.Domain.GlueInfoRow.Get_GlueInfo,
                        ShapeType.Shape),
                    new Presentation.Views.GlueInfo()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }
        public static DxThemedWindowHost _shapeGradientPropertiesHost = null;

        private void btnGradientProperties_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            DxThemedWindowHost.DisplayUserControlInHost(ref _shapeGradientPropertiesHost,
            "Gradient Properties",
            600, 450,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetSection(
                    new Presentation.ViewModels.ObjectViewModel<VNCVisioAddIn.Domain.GradientPropertiesRow, Presentation.ModelWrappers.GradientPropertiesWrapper>(
                        "Update Gradient Properties",
                        VNC.Visio.VSTOAddIn.Domain.GradientPropertiesRow.Get_GradientProperties,
                        ShapeType.Shape),
                    new Presentation.Views.GradientProperties()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost _shapeShapeLayoutHost = null;

        private void btnShapeLayout_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            DxThemedWindowHost.DisplayUserControlInHost(ref _shapeShapeLayoutHost,
            "Shape Layout",
            600, 450,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetSection(
                    new Presentation.ViewModels.ObjectViewModel<VNCVisioAddIn.Domain.ShapeLayoutRow, Presentation.ModelWrappers.ShapeLayoutWrapper>(
                        "Update Shape Layout",
                        VNC.Visio.VSTOAddIn.Domain.ShapeLayoutRow.GetRow,
                        ShapeType.Shape),
                    new Presentation.Views.ShapeLayout()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost _shapeGroupPropertiesHost = null;

        private void btnGroupProperties_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            DxThemedWindowHost.DisplayUserControlInHost(ref _shapeGroupPropertiesHost,
            "Group Properties",
            600, 450,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetSection(
                    new Presentation.ViewModels.ObjectViewModel<VNCVisioAddIn.Domain.GroupPropertiesRow, Presentation.ModelWrappers.GroupPropertiesWrapper>(
                        "Update Group Properties",
                        VNC.Visio.VSTOAddIn.Domain.GroupPropertiesRow.Get_GroupProperties,
                        ShapeType.Shape),
                    new Presentation.Views.GroupProperties()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost _shapeImagePropertiesHost = null;

        private void btnImageProperties_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            DxThemedWindowHost.DisplayUserControlInHost(ref _shapeImagePropertiesHost,
                "Image Properties",
                600, 450,
                ShowWindowMode.Modeless_Show,
                    new Presentation.Views.ShapeSheetSection(
                        new Presentation.ViewModels.ObjectViewModel<VNCVisioAddIn.Domain.ImagePropertiesRow, Presentation.ModelWrappers.ImagePropertiesWrapper>(
                            "Update Group Properties",
                            VNC.Visio.VSTOAddIn.Domain.ImagePropertiesRow.Get_ImageProperties,
                            ShapeType.Shape),
                        new Presentation.Views.ImageProperties()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost ssLayerMembership_ShapeSheetSectionHost = null;

        private void btnLayerMembership_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            DxThemedWindowHost.DisplayUserControlInHost(ref ssLayerMembership_ShapeSheetSectionHost,
            "Layer Membership",
            600, 450,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetSection(
                    new Presentation.ViewModels.ObjectViewModel<VNCVisioAddIn.Domain.LayerMembershipRow, Presentation.ModelWrappers.LayerMembershipWrapper>(
                        "Update Layer Membership",
                        VNC.Visio.VSTOAddIn.Domain.LayerMembershipRow.GetRow,
                        ShapeType.Shape),
                    new Presentation.Views.LayerMembership()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost _shapeLineFormatHost = null;

        private void btnLineFormat_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            DxThemedWindowHost.DisplayUserControlInHost(ref _shapeLineFormatHost,
                "Line Format",
                600, 450,
                ShowWindowMode.Modeless_Show,
                    new Presentation.Views.ShapeSheetSection(
                        new Presentation.ViewModels.ObjectViewModel<VNCVisioAddIn.Domain.LineFormatRow, Presentation.ModelWrappers.LineFormatWrapper>(
                            "Update LineFormat",
                            VNC.Visio.VSTOAddIn.Domain.LineFormatRow.Get_LineFormat,
                            ShapeType.Shape),
                        new Presentation.Views.LineFormat()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost _shapeMiscellaneousHost = null;

        private void btnMiscelleaneous_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            DxThemedWindowHost.DisplayUserControlInHost(ref _shapeMiscellaneousHost,
                "Miscellaneous",
                600, 450,
                ShowWindowMode.Modeless_Show,
                    new Presentation.Views.ShapeSheetSection(
                        new Presentation.ViewModels.ObjectViewModel<VNCVisioAddIn.Domain.MiscellaneousRow, Presentation.ModelWrappers.MiscellaneousWrapper>(
                            "Update Miscellaneous",
                            VNC.Visio.VSTOAddIn.Domain.MiscellaneousRow.GetRow,
                            ShapeType.Shape),
                        new Presentation.Views.Miscellaneous()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost _shapeProtectionHost = null;

        private void btnProtection_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            DxThemedWindowHost.DisplayUserControlInHost(ref _shapeProtectionHost,
                "Protection",
                600, 450,
                ShowWindowMode.Modeless_Show,
                    new Presentation.Views.ShapeSheetSection(
                        new Presentation.ViewModels.ObjectViewModel<VNCVisioAddIn.Domain.ProtectionRow, Presentation.ModelWrappers.ProtectionWrapper>(
                            "Update Protection",
                            VNC.Visio.VSTOAddIn.Domain.ProtectionRow.Get_Protection,
                            ShapeType.Shape),
                        new Presentation.Views.Protection()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost _shapeQuickStyleHost = null;

        private void btnQuickStyle_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            DxThemedWindowHost.DisplayUserControlInHost(ref _shapeQuickStyleHost,
                "Quick Style",
                600, 450,
                ShowWindowMode.Modeless_Show,
                    new Presentation.Views.ShapeSheetSection(
                        new Presentation.ViewModels.ObjectViewModel<VNCVisioAddIn.Domain.QuickStyleRow, Presentation.ModelWrappers.QuickStyleWrapper>(
                            "Update Layer Membership",
                            VNC.Visio.VSTOAddIn.Domain.QuickStyleRow.GetRow,
                            ShapeType.Shape),
                        new Presentation.Views.QuickStyle()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost _shapeShapeTransformHost = null;

        private void btnShapeTransform_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            DxThemedWindowHost.DisplayUserControlInHost(ref _shapeShapeTransformHost,
                "Shape Transform",
                800,600,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetSection(
                    new Presentation.ViewModels.ObjectViewModel<VNCVisioAddIn.Domain.ShapeTransformRow, Presentation.ModelWrappers.ShapeTransformWrapper>(
                        "Update Shape Transform",
                        VNC.Visio.VSTOAddIn.Domain.ShapeTransformRow.Get_ShapeTransform,
                        ShapeType.Shape),
                    new Presentation.Views.ShapeTransform()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost _shapeTextBlockFormatHost = null;

        private void btnTextBlockFormat_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            DxThemedWindowHost.DisplayUserControlInHost(ref _shapeTextBlockFormatHost,
            "Shape TextBlock Format",
            600, 450,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetSection(
                    new Presentation.ViewModels.ObjectViewModel<VNCVisioAddIn.Domain.TextBlockFormatRow, Presentation.ModelWrappers.TextBlockFormatWrapper>(
                        "Update Text Block Format",
                        VNC.Visio.VSTOAddIn.Domain.TextBlockFormatRow.GetRow,
                        ShapeType.Shape),
                    new Presentation.Views.TextBlockFormat()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost _shapeTextTransformHost = null;

        private void btnTextTransform_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            DxThemedWindowHost.DisplayUserControlInHost(ref _shapeTextTransformHost,
            "Shape Text Transform",
            600, 450,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetSection(
                    new Presentation.ViewModels.ObjectViewModel<VNCVisioAddIn.Domain.TextTransformRow, Presentation.ModelWrappers.TextTransformWrapper>(
                        "Update 3-D Rotation Properties",
                        VNC.Visio.VSTOAddIn.Domain.TextTransformRow.Get_TextTransform,
                        ShapeType.Shape),
                    new Presentation.Views.TextTransform()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }
        public static DxThemedWindowHost _shapeThemePropertiesHost = null;

        private void btnShapeThemeProperties_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            DxThemedWindowHost.DisplayUserControlInHost(ref _shapeThemePropertiesHost,
                "Shape ThemeProperties",
                600, 800,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetSection(
                    new Presentation.ViewModels.ObjectViewModel<VNCVisioAddIn.Domain.ThemePropertiesRow, Presentation.ModelWrappers.ThemePropertiesWrapper>(
                        "Update ThemeProperties",
                        VNC.Visio.VSTOAddIn.Domain.ThemePropertiesRow.Get_ThemeProperties,
                        ShapeType.Shape),
                    new Presentation.Views.ThemeProperties()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region UI Events Shape - Row Based

        #region Actions

        public static DxThemedWindowHost _pageActionsHost = null;

        private void btnActionsPage_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            DxThemedWindowHost.DisplayUserControlInHost(ref _pageActionsHost,
                "Actions (Page)",
                600, 800,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetSection(
                    new Presentation.ViewModels.RowsViewModel<VNCVisioAddIn.Domain.ActionRow, Presentation.ModelWrappers.ActionRowWrapper>(
                        "Update Actions",
                        VNC.Visio.VSTOAddIn.Domain.ActionRow.GetRows,
                        ShapeType.Page),
                    new Presentation.Views.Actions()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost _shapeActionsHost = null;

        private void btnActionsShape_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            DxThemedWindowHost.DisplayUserControlInHost(ref _shapeActionsHost,
                "Actions (Shape)",
                600, 800,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetSection(
                    new Presentation.ViewModels.RowsViewModel<VNCVisioAddIn.Domain.ActionRow, Presentation.ModelWrappers.ActionRowWrapper>(
                        "Update Actions",
                        VNC.Visio.VSTOAddIn.Domain.ActionRow.GetRows,
                        ShapeType.Shape),
                    new Presentation.Views.Actions()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region ActionTags

        public static DxThemedWindowHost _pageActionTagsHost = null;

        private void btnActionTagsPage_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            DxThemedWindowHost.DisplayUserControlInHost(ref _pageActionTagsHost,
                "ActionsTags (Page)",
                600, 750,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetSection(
                    new Presentation.ViewModels.RowsViewModel<VNCVisioAddIn.Domain.ActionTagRow, Presentation.ModelWrappers.ActionTagRowWrapper>(
                        "Hello Natalie", 
                        VNC.Visio.VSTOAddIn.Domain.ActionTagRow.GetRows,
                        ShapeType.Page),
                    new Presentation.Views.ActionTags()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost _shapeActionTagsHost = null;

        private void btnActionTagsShape_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            DxThemedWindowHost.DisplayUserControlInHost(ref _shapeActionTagsHost,
                "ActionsTags (Shape)",
                600, 750,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetSection(
                    new Presentation.ViewModels.RowsViewModel<VNCVisioAddIn.Domain.ActionTagRow, Presentation.ModelWrappers.ActionTagRowWrapper>(
                        "Hello Natalie", 
                        VNC.Visio.VSTOAddIn.Domain.ActionTagRow.GetRows,
                        ShapeType.Shape),
                    new Presentation.Views.ActionTags()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        private void btnCharacter_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);


            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost ssConnectionPoints_ShapeSheetSectionHost = null;

        private void btnConnectionPoints_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            DxThemedWindowHost.DisplayUserControlInHost(ref ssConnectionPoints_ShapeSheetSectionHost,
            "Connection Points",
            600, 450,
            //Common.DEFAULT_WINDOW_WIDTH, Common.DEFAULT_WINDOW_HEIGHT,
                ShowWindowMode.Modeless_Show,
            new Presentation.Views.ShapeSheetSection(
                new Presentation.ViewModels.ConnectionPointsViewModel(),
                new Presentation.Views.ConnectionPoints()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost ssControls_ShapeSheetSectionHost = null;

        private void btnControls_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            DxThemedWindowHost.DisplayUserControlInHost(ref ssControls_ShapeSheetSectionHost,
            "Controls",
            600, 450,
            //Common.DEFAULT_WINDOW_WIDTH, Common.DEFAULT_WINDOW_HEIGHT,
                ShowWindowMode.Modeless_Show,
            new Presentation.Views.ShapeSheetSection(
                new Presentation.ViewModels.ControlsViewModel(),
                new Presentation.Views.Controls()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void btnGeometry_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);


            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost ssGradientStops_ShapeSheetSectionHost = null;

        private void btnGradientStops_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);


            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #region Hyperlinks

        public static DxThemedWindowHost _documentHyperLinksHost = null;

        private void btnDocumentHyperlinks_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            DxThemedWindowHost.DisplayUserControlInHost(ref _documentHyperLinksHost,
                "Hyperlinks (Document)",
                800, 700,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetSection(
                    new Presentation.ViewModels.RowsViewModel<VNCVisioAddIn.Domain.HyperlinkRow, Presentation.ModelWrappers.HyperlinkRowWrapper>(
                        "Update Hyperlinks",
                        VNC.Visio.VSTOAddIn.Domain.HyperlinkRow.GetRows,
                        ShapeType.Document),
                    new Presentation.Views.Hyperlinks()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost _pageHyperLinksHost = null;

        private void btnPageHyperlinks_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            DxThemedWindowHost.DisplayUserControlInHost(ref _pageHyperLinksHost,
                "Hyperlinks (Page)",
                800, 700,
                ShowWindowMode.Modeless_Show,
                    new Presentation.Views.ShapeSheetSection(
                    new Presentation.ViewModels.RowsViewModel<VNCVisioAddIn.Domain.HyperlinkRow, Presentation.ModelWrappers.HyperlinkRowWrapper>(
                        "Update Hyperlinks",
                        VNC.Visio.VSTOAddIn.Domain.HyperlinkRow.GetRows,
                        ShapeType.Page),
                    new Presentation.Views.Hyperlinks()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost _shapeHyperlinksHost = null;

        private void btnShapeHyperlinks_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            DxThemedWindowHost.DisplayUserControlInHost(ref _shapeHyperlinksHost,
                "Hyperlinks (Shape)",
                800, 700,
                ShowWindowMode.Modeless_Show,
                    new Presentation.Views.ShapeSheetSection(
                    new Presentation.ViewModels.RowsViewModel<VNCVisioAddIn.Domain.HyperlinkRow, Presentation.ModelWrappers.HyperlinkRowWrapper>(
                        "Update Hyperlinks",
                        VNC.Visio.VSTOAddIn.Domain.HyperlinkRow.GetRows,
                        ShapeType.Shape),
                    new Presentation.Views.Hyperlinks()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region Layers

        public static DxThemedWindowHost _pageLayersHost = null;

        private void btnLayers_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            DxThemedWindowHost.DisplayUserControlInHost(ref _pageLayersHost,
                "Layers (Page)",
                800, 700,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetSection(
                    new Presentation.ViewModels.RowsViewModel<VNCVisioAddIn.Domain.LayerRow, Presentation.ModelWrappers.LayerRowWrapper>(
                        "Update Layers",
                        VNC.Visio.VSTOAddIn.Domain.LayerRow.GetRows,
                        ShapeType.Page),
                    new Presentation.Views.Layers()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        private void btnParagraph_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);


        }

        private void btnTabs_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);


            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #region Scratch

        public static DxThemedWindowHost _documentScratchHost = null;

        private void btnDocumentScratch_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            DxThemedWindowHost.DisplayUserControlInHost(ref _documentScratchHost,
                "Scratch (Document)",
                800, 700,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetSection(
                    new Presentation.ViewModels.RowsViewModel<VNCVisioAddIn.Domain.ScratchRow, Presentation.ModelWrappers.ScratchRowWrapper>(
                        "Update Scratch",
                        VNC.Visio.VSTOAddIn.Domain.ScratchRow.GetRows,
                        ShapeType.Document),
                    new Presentation.Views.Scratch()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost _pageScratchHost = null;

        private void btnPageScratch_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            DxThemedWindowHost.DisplayUserControlInHost(ref _pageScratchHost,
                "Scratch (Page)",
                800, 700,
                ShowWindowMode.Modeless_Show,
                    new Presentation.Views.ShapeSheetSection(
                    new Presentation.ViewModels.RowsViewModel<VNCVisioAddIn.Domain.ScratchRow, Presentation.ModelWrappers.ScratchRowWrapper>(
                        "Update Scratch",
                        VNC.Visio.VSTOAddIn.Domain.ScratchRow.GetRows,
                        ShapeType.Page),
                    new Presentation.Views.Scratch()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost _shapeScratchHost = null;

        private void btnShapeScratch_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            DxThemedWindowHost.DisplayUserControlInHost(ref _shapeScratchHost,
                "Scratch (Shape)",
                800, 700,
                ShowWindowMode.Modeless_Show,
                    new Presentation.Views.ShapeSheetSection(
                    new Presentation.ViewModels.RowsViewModel<VNCVisioAddIn.Domain.ScratchRow, Presentation.ModelWrappers.ScratchRowWrapper>(
                        "Update Scratch",
                        VNC.Visio.VSTOAddIn.Domain.ScratchRow.GetRows,
                        ShapeType.Shape),
                    new Presentation.Views.Scratch()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region ShapeData

        public static DxThemedWindowHost _documentShapeDataHost = null;

        private void btnDocumentShapeData_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            DxThemedWindowHost.DisplayUserControlInHost(ref _documentShapeDataHost,
                "Shape Data (Document)",
                800, 700,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetSection(
                    new Presentation.ViewModels.RowsViewModel<VNCVisioAddIn.Domain.ShapeDataRow, Presentation.ModelWrappers.ShapeDataRowWrapper>(
                        "Update ShapeData",
                        VNC.Visio.VSTOAddIn.Domain.ShapeDataRow.GetRows,
                        ShapeType.Document),
                    new Presentation.Views.ShapeData()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost _pageShapeDataHost = null;

        private void btnPageShapeData_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            DxThemedWindowHost.DisplayUserControlInHost(ref _pageShapeDataHost,
                "Shape Data (Page)",
                800, 700,
                ShowWindowMode.Modeless_Show,
                    new Presentation.Views.ShapeSheetSection(
                    new Presentation.ViewModels.RowsViewModel<VNCVisioAddIn.Domain.ShapeDataRow, Presentation.ModelWrappers.ShapeDataRowWrapper>(
                        "Update ShapeData",
                        VNC.Visio.VSTOAddIn.Domain.ShapeDataRow.GetRows,
                        ShapeType.Page),
                    new Presentation.Views.ShapeData()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost _shapeShapeDataHost = null;

        private void btnShapeShapeData_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            DxThemedWindowHost.DisplayUserControlInHost(ref _shapeShapeDataHost,
                "Scratch (Shape)",
                800, 700,
                ShowWindowMode.Modeless_Show,
                    new Presentation.Views.ShapeSheetSection(
                    new Presentation.ViewModels.RowsViewModel<VNCVisioAddIn.Domain.ShapeDataRow, Presentation.ModelWrappers.ShapeDataRowWrapper>(
                        "Update ShapeData",
                        VNC.Visio.VSTOAddIn.Domain.ShapeDataRow.GetRows,
                        ShapeType.Shape),
                    new Presentation.Views.ShapeData()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region UserDefinedCells

        public static DxThemedWindowHost _documentUserDefineCellsHost = null;

        private void btnDocumentUserDefinedCells_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            DxThemedWindowHost.DisplayUserControlInHost(ref _documentUserDefineCellsHost,
                "User-Defined Cells (Document)",
                800, 700,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetSection(
                    new Presentation.ViewModels.RowsViewModel<VNCVisioAddIn.Domain.UserDefinedCellRow, Presentation.ModelWrappers.UserDefinedCellRowWrapper>(
                        "Update ShapeData",
                        VNC.Visio.VSTOAddIn.Domain.UserDefinedCellRow.Get_UserDefinedCellsRows,
                        ShapeType.Document),
                    new Presentation.Views.UserDefinedCells()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost _pageUserDefineCellsHost = null;

        private void btnPageUserDefinedCells_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            DxThemedWindowHost.DisplayUserControlInHost(ref _pageUserDefineCellsHost,
                "User-Defined Cells (Page)",
                800, 700,
                ShowWindowMode.Modeless_Show,
                    new Presentation.Views.ShapeSheetSection(
                    new Presentation.ViewModels.RowsViewModel<VNCVisioAddIn.Domain.UserDefinedCellRow, Presentation.ModelWrappers.UserDefinedCellRowWrapper>(
                        "Update ShapeData",
                        VNC.Visio.VSTOAddIn.Domain.UserDefinedCellRow.Get_UserDefinedCellsRows,
                        ShapeType.Page),
                    new Presentation.Views.UserDefinedCells()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost _shapeUserDefineCellsHost = null;

        private void btnShapeUserDefinedCells_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            DxThemedWindowHost.DisplayUserControlInHost(ref _shapeUserDefineCellsHost,
                "User-Defined Cells (Shape)",
                800, 700,
                ShowWindowMode.Modeless_Show,
                    new Presentation.Views.ShapeSheetSection(
                    new Presentation.ViewModels.RowsViewModel<VNCVisioAddIn.Domain.UserDefinedCellRow, Presentation.ModelWrappers.UserDefinedCellRowWrapper>(
                        "Update ShapeData",
                        VNC.Visio.VSTOAddIn.Domain.UserDefinedCellRow.Get_UserDefinedCellsRows,
                        ShapeType.Shape),
                    new Presentation.Views.UserDefinedCells()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #endregion

        #endregion Event Handlers
    }
}