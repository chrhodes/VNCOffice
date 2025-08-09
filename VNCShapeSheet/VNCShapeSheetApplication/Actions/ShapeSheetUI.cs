using System;
using System.Reflection;
using System.Windows.Forms;

using VNC;
using VNC.Core.Presentation;
using VNC.WPF.Presentation.Dx.Views;

using VNCVisioAddIn = VNC.Visio.VSTOAddIn;

namespace VNCShapeSheetApplication.Actions
{
    public class ShapeSheetUI
    {
        #region Event Handlers

        #region UI Events - Object based - Uses ObjectViewModel

        #region ShapeSheet UI Events Document Related

        public static DxThemedWindowHost _documentPropertiesHost = null;

        public static void DocumentProperties()
        {
            Int64 startTicks = Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            if (_documentPropertiesHost is null) _documentPropertiesHost = new DxThemedWindowHost();

            _documentPropertiesHost.DisplayUserControlInHost(
                "DocumentProperties",
                650, 550,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetObjectSection(
                    new Presentation.ViewModels.ObjectViewModel<VNCVisioAddIn.Domain.DocumentPropertiesRow, 
                        VNCVisioAddIn.Presentation.ModelWrappers.DocumentPropertiesWrapper>(
                        "Update DocumentProperties",
                        VNCVisioAddIn.Domain.DocumentPropertiesRow.GetRow,
                        VNCVisioAddIn.Domain.DocumentPropertiesRow.SetRow,
                        VNCVisioAddIn.Domain.ShapeType.Document),
                    new Presentation.Views.DocumentProperties()));

            //Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}() Exit", startTicks);
        }

        #endregion

        #region Style Related

        // TODO(crhodes)
        // 

        #endregion

        #region ShapeSheet UI Events Page Related

        public static DxThemedWindowHost _pagePageLayoutHost = null;

        public static void PageLayout()
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (_pagePageLayoutHost is null) _pagePageLayoutHost = new DxThemedWindowHost();

            _pagePageLayoutHost.DisplayUserControlInHost(
                "PageLayout",
                550, 700,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetObjectSection(
                    new Presentation.ViewModels.ObjectViewModel<VNCVisioAddIn.Domain.PageLayoutRow,
                        VNCVisioAddIn.Presentation.ModelWrappers.PageLayoutWrapper>(
                        "Update PageLayout",
                        VNCVisioAddIn.Domain.PageLayoutRow.GetRow,
                        VNCVisioAddIn.Domain.PageLayoutRow.SetRow,
                        VNCVisioAddIn.Domain.ShapeType.Page),
                    new Presentation.Views.PageLayout()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost _pagePagePropertiesHost = null;

        public static void PageProperties()
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (_pagePagePropertiesHost is null) _pagePagePropertiesHost = new DxThemedWindowHost();

            _pagePagePropertiesHost.DisplayUserControlInHost(
                "PageProperties",
                600, 575,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetObjectSection(
                    new Presentation.ViewModels.ObjectViewModel<VNCVisioAddIn.Domain.PagePropertiesRow, 
                        VNCVisioAddIn.Presentation.ModelWrappers.PagePropertiesWrapper>(
                        "Update PageProperties",
                        VNCVisioAddIn.Domain.PagePropertiesRow.GetRow,
                        VNCVisioAddIn.Domain.PagePropertiesRow.SetRow,
                        VNCVisioAddIn.Domain.ShapeType.Page),
                    new Presentation.Views.PageProperties()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost _pagePrintPropertiesHost = null;

        public static void PrintProperties()
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (_pagePrintPropertiesHost is null) _pagePrintPropertiesHost = new DxThemedWindowHost();

            _pagePrintPropertiesHost.DisplayUserControlInHost(
                "PrintProperties",
                600, 600,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetObjectSection(
                    new Presentation.ViewModels.ObjectViewModel<VNCVisioAddIn.Domain.PrintPropertiesRow, 
                        VNCVisioAddIn.Presentation.ModelWrappers.PrintPropertiesWrapper>(
                        "Update PrintProperties",
                        VNCVisioAddIn.Domain.PrintPropertiesRow.GetRow,
                        VNCVisioAddIn.Domain.PrintPropertiesRow.SetRow,
                        VNCVisioAddIn.Domain.ShapeType.Page),
                    new Presentation.Views.PrintProperties()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost _pageRulerAndGridsHost = null;

        public static void RulerAndGrid()
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (_pageRulerAndGridsHost is null) _pageRulerAndGridsHost = new DxThemedWindowHost();

            _pageRulerAndGridsHost.DisplayUserControlInHost(
                "Ruler & Grid",
                600, 300,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetObjectSection(
                    new Presentation.ViewModels.ObjectViewModel<VNCVisioAddIn.Domain.RulerAndGridRow, 
                        VNCVisioAddIn.Presentation.ModelWrappers.RulerAndGridWrapper>(
                        "Update Ruler & Grid",
                        VNCVisioAddIn.Domain.RulerAndGridRow.GetRow,
                        VNCVisioAddIn.Domain.RulerAndGridRow.SetRow,
                        VNCVisioAddIn.Domain.ShapeType.Page),
                    new Presentation.Views.RulerAndGrid()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost _pageThemePropertiesHost = null;

        public static void PageThemeProperties()
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (_pageThemePropertiesHost is null) _pageThemePropertiesHost = new DxThemedWindowHost();

            _pageThemePropertiesHost.DisplayUserControlInHost(
                "ThemeProperties",
                600, 400,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetObjectSection(
                    new Presentation.ViewModels.ObjectViewModel<VNCVisioAddIn.Domain.ThemePropertiesRow, 
                        VNCVisioAddIn.Presentation.ModelWrappers.ThemePropertiesWrapper>(
                        "Update ThemeProperties",
                        VNCVisioAddIn.Domain.ThemePropertiesRow.GetRow,
                        VNCVisioAddIn.Domain.ThemePropertiesRow.SetRow,
                        VNCVisioAddIn.Domain.ShapeType.Page),
                    new Presentation.Views.ThemeProperties()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region ShapeSheet UI Events Shape Related

        public static DxThemedWindowHost _shapeOneDEndpointsHost = null;

        public static void OneDEndpoints()
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (_shapeOneDEndpointsHost is null) _shapeOneDEndpointsHost = new DxThemedWindowHost();

            _shapeOneDEndpointsHost.DisplayUserControlInHost(
                "1-D Endpoints",
                600, 300,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetObjectSection(
                    new Presentation.ViewModels.ObjectViewModel<VNCVisioAddIn.Domain.OneDEndPointsRow, 
                        VNCVisioAddIn.Presentation.ModelWrappers.OneDEndPointsWrapper>(
                        "Update 1-D Endpoints",
                        VNCVisioAddIn.Domain.OneDEndPointsRow.GetRow,
                        VNCVisioAddIn.Domain.OneDEndPointsRow.SetRow,
                        VNCVisioAddIn.Domain.ShapeType.Shape),
                    new Presentation.Views.OneDEndPoints()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost _shapeThreeDRotationPropertiesHost = null;

        public static void ThreeDRotationProperties()
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (_shapeThreeDRotationPropertiesHost is null) _shapeThreeDRotationPropertiesHost = new DxThemedWindowHost();

            _shapeThreeDRotationPropertiesHost.DisplayUserControlInHost(
                "3-D RotationProperties",
                600, 450,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetObjectSection(
                    new Presentation.ViewModels.ObjectViewModel<VNCVisioAddIn.Domain.ThreeDRotationPropertiesRow, 
                        VNCVisioAddIn.Presentation.ModelWrappers.ThreeDRotationPropertiesWrapper>(
                        "Update 3-D RotationProperties",
                        VNCVisioAddIn.Domain.ThreeDRotationPropertiesRow.GetRow,
                        VNCVisioAddIn.Domain.ThreeDRotationPropertiesRow.SetRow,
                        VNCVisioAddIn.Domain.ShapeType.Shape),
                    new Presentation.Views.OneDEndPoints()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost _shapeAdditionalEffectPropertiesHost = null;

        public static void AdditionalEffectProperties()
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (_shapeAdditionalEffectPropertiesHost is null) _shapeAdditionalEffectPropertiesHost = new DxThemedWindowHost();

            _shapeAdditionalEffectPropertiesHost.DisplayUserControlInHost(
                "AdditionalEffectProperties",
                600, 450,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetObjectSection(
                    new Presentation.ViewModels.ObjectViewModel<VNCVisioAddIn.Domain.AdditionalEffectPropertiesRow, 
                        VNCVisioAddIn.Presentation.ModelWrappers.AdditionalEffectPropertiesWrapper>(
                        "Update AdditionalEffectProperties",
                        VNCVisioAddIn.Domain.AdditionalEffectPropertiesRow.GetRow,
                        VNCVisioAddIn.Domain.AdditionalEffectPropertiesRow.SetRow,
                        VNCVisioAddIn.Domain.ShapeType.Shape),
                    new Presentation.Views.AdditionalEffectProperties()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost _shapeAlignmentHost = null;

        public static void Alignment()
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (_shapeAlignmentHost is null) _shapeAlignmentHost = new DxThemedWindowHost();

            _shapeAlignmentHost.DisplayUserControlInHost(
                "Alignment",
                600, 450,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetObjectSection(
                    new Presentation.ViewModels.ObjectViewModel<VNCVisioAddIn.Domain.AlignmentRow, 
                        VNCVisioAddIn.Presentation.ModelWrappers.AlignmentRowWrapper>(
                        "Update Alignment",
                        VNCVisioAddIn.Domain.AlignmentRow.GetRow,
                        VNCVisioAddIn.Domain.AlignmentRow.SetRow,
                        VNCVisioAddIn.Domain.ShapeType.Shape),
                    new Presentation.Views.Alignment()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost _shapeBevelPropertiesHost = null;

        public static void BevelProperties()
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (_shapeBevelPropertiesHost is null) _shapeBevelPropertiesHost = new DxThemedWindowHost();

            _shapeBevelPropertiesHost.DisplayUserControlInHost(
            "BevelProperties",
            600, 450,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetObjectSection(
                    new Presentation.ViewModels.ObjectViewModel<VNCVisioAddIn.Domain.BevelPropertiesRow, 
                        VNCVisioAddIn.Presentation.ModelWrappers.BevelPropertiesRowWrapper>(
                        "Update BevelProperties",
                        VNCVisioAddIn.Domain.BevelPropertiesRow.GetRow,
                        VNCVisioAddIn.Domain.BevelPropertiesRow.SetRow,
                        VNCVisioAddIn.Domain.ShapeType.Shape),
                    new Presentation.Views.BevelProperties()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost _shapeChangeShapeBehaviorHost = null;

        public static void ChangeShapeBehavior()
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (_shapeChangeShapeBehaviorHost is null) _shapeChangeShapeBehaviorHost = new DxThemedWindowHost();

            _shapeChangeShapeBehaviorHost.DisplayUserControlInHost(
            "ChangeShapeBehavior",
            600, 450,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetObjectSection(
                    new Presentation.ViewModels.ObjectViewModel<VNCVisioAddIn.Domain.ChangeShapeBehaviorRow, 
                        VNCVisioAddIn.Presentation.ModelWrappers.ChangeShapeBehaviorWrapper>(
                        "Update ChangeShapeBehavior",
                        VNCVisioAddIn.Domain.ChangeShapeBehaviorRow.GetRow,
                        VNCVisioAddIn.Domain.ChangeShapeBehaviorRow.SetRow,
                        VNCVisioAddIn.Domain.ShapeType.Shape),
                    new Presentation.Views.ChangeShapeBehavior()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost _shapeEventsHost = null;

        public static void Events()
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (_shapeEventsHost is null) _shapeEventsHost = new DxThemedWindowHost();

            _shapeEventsHost.DisplayUserControlInHost(
            "Events",
            600, 450,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetObjectSection(
                    new Presentation.ViewModels.ObjectViewModel<VNCVisioAddIn.Domain.EventsRow, 
                        VNCVisioAddIn.Presentation.ModelWrappers.EventsWrapper>(
                        "Update Events",
                        VNCVisioAddIn.Domain.EventsRow.GetRow,
                        VNCVisioAddIn.Domain.EventsRow.SetRow,
                        VNCVisioAddIn.Domain.ShapeType.Shape),
                    new Presentation.Views.Events()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost _shapeFillFormatHost = null;

        public static void FillFormat()
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (_shapeFillFormatHost is null) _shapeFillFormatHost = new DxThemedWindowHost();

            _shapeFillFormatHost.DisplayUserControlInHost(
            "FillFormat",
            600, 450,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetObjectSection(
                    new Presentation.ViewModels.ObjectViewModel<VNCVisioAddIn.Domain.FillFormatRow, 
                        VNCVisioAddIn.Presentation.ModelWrappers.FillFormatWrapper>(
                        "Update FillFormat",
                        VNCVisioAddIn.Domain.FillFormatRow.GetRow,
                        VNCVisioAddIn.Domain.FillFormatRow.SetRow,
                        VNCVisioAddIn.Domain.ShapeType.Shape),
                    new Presentation.Views.FillFormat()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost _shapeGlueInfoHost = null;

        public static void GlueInfo()
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (_shapeGlueInfoHost is null) _shapeGlueInfoHost = new DxThemedWindowHost();

            _shapeGlueInfoHost.DisplayUserControlInHost(
            "GlueInfo",
            600, 450,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetObjectSection(
                    new Presentation.ViewModels.ObjectViewModel<VNCVisioAddIn.Domain.GlueInfoRow, 
                        VNCVisioAddIn.Presentation.ModelWrappers.GlueInfoWrapper>(
                        "Update GlueInfo",
                        VNCVisioAddIn.Domain.GlueInfoRow.GetRow,
                        VNCVisioAddIn.Domain.GlueInfoRow.SetRow,
                        VNCVisioAddIn.Domain.ShapeType.Shape),
                    new Presentation.Views.GlueInfo()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost _shapeGradientPropertiesHost = null;

        public static void GradientProperties()
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (_shapeGradientPropertiesHost is null) _shapeGradientPropertiesHost = new DxThemedWindowHost();

            _shapeGradientPropertiesHost.DisplayUserControlInHost(
            "GradientProperties",
            600, 450,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetObjectSection(
                    new Presentation.ViewModels.ObjectViewModel<VNCVisioAddIn.Domain.GradientPropertiesRow, 
                        VNCVisioAddIn.Presentation.ModelWrappers.GradientPropertiesWrapper>(
                        "Update GradientProperties",
                        VNCVisioAddIn.Domain.GradientPropertiesRow.GetRow,
                        VNCVisioAddIn.Domain.GradientPropertiesRow.SetRow,
                        VNCVisioAddIn.Domain.ShapeType.Shape),
                    new Presentation.Views.GradientProperties()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }


        public static DxThemedWindowHost _shapeGroupPropertiesHost = null;

        public static void GroupProperties()
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (_shapeGroupPropertiesHost is null) _shapeGroupPropertiesHost = new DxThemedWindowHost();

            _shapeGroupPropertiesHost.DisplayUserControlInHost(
            "GroupProperties",
            600, 450,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetObjectSection(
                    new Presentation.ViewModels.ObjectViewModel<VNCVisioAddIn.Domain.GroupPropertiesRow, 
                        VNCVisioAddIn.Presentation.ModelWrappers.GroupPropertiesWrapper>(
                        "Update GroupProperties",
                        VNCVisioAddIn.Domain.GroupPropertiesRow.GetRow,
                        VNCVisioAddIn.Domain.GroupPropertiesRow.SetRow,
                        VNCVisioAddIn.Domain.ShapeType.Shape),
                    new Presentation.Views.GroupProperties()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost _shapeImagePropertiesHost = null;

        public static void ImageProperties()
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (_shapeImagePropertiesHost is null) _shapeImagePropertiesHost = new DxThemedWindowHost();

            _shapeImagePropertiesHost.DisplayUserControlInHost(
                "ImageProperties",
                600, 450,
                ShowWindowMode.Modeless_Show,
                    new Presentation.Views.ShapeSheetObjectSection(
                        new Presentation.ViewModels.ObjectViewModel<VNCVisioAddIn.Domain.ImagePropertiesRow, 
                            VNCVisioAddIn.Presentation.ModelWrappers.ImagePropertiesWrapper>(
                            "Update ImageProperties",
                            VNCVisioAddIn.Domain.ImagePropertiesRow.GetRow,
                            VNCVisioAddIn.Domain.ImagePropertiesRow.SetRow,
                            VNCVisioAddIn.Domain.ShapeType.Shape),
                        new Presentation.Views.ImageProperties()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost ssLayerMembership_ShapeSheetSectionHost = null;

        public static void LayerMembership()
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (ssLayerMembership_ShapeSheetSectionHost is null) ssLayerMembership_ShapeSheetSectionHost = new DxThemedWindowHost();

            ssLayerMembership_ShapeSheetSectionHost.DisplayUserControlInHost(
            "LayerMembership",
            600, 450,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetObjectSection(
                    new Presentation.ViewModels.ObjectViewModel<VNCVisioAddIn.Domain.LayerMembershipRow, 
                        VNCVisioAddIn.Presentation.ModelWrappers.LayerMembershipWrapper>(
                        "Update LayerMembership",
                        VNCVisioAddIn.Domain.LayerMembershipRow.GetRow,
                        VNCVisioAddIn.Domain.LayerMembershipRow.SetRow,
                        VNCVisioAddIn.Domain.ShapeType.Shape),
                    new Presentation.Views.LayerMembership()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost _shapeLineFormatHost = null;

        public static void LineFormat()
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (_shapeLineFormatHost is null) _shapeLineFormatHost = new DxThemedWindowHost();

            _shapeLineFormatHost.DisplayUserControlInHost(
                "LineFormat",
                600, 450,
                ShowWindowMode.Modeless_Show,
                    new Presentation.Views.ShapeSheetObjectSection(
                        new Presentation.ViewModels.ObjectViewModel<VNCVisioAddIn.Domain.LineFormatRow, 
                            VNCVisioAddIn.Presentation.ModelWrappers.LineFormatWrapper>(
                            "Update LineFormat",
                            VNCVisioAddIn.Domain.LineFormatRow.GetRow,
                            VNCVisioAddIn.Domain.LineFormatRow.SetRow,
                            VNCVisioAddIn.Domain.ShapeType.Shape),
                        new Presentation.Views.LineFormat()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost _shapeMiscellaneousHost = null;

        public static void Miscelleaneous()
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (_shapeMiscellaneousHost is null) _shapeMiscellaneousHost = new DxThemedWindowHost();

            _shapeMiscellaneousHost.DisplayUserControlInHost(
                "Miscellaneous",
                600, 450,
                ShowWindowMode.Modeless_Show,
                    new Presentation.Views.ShapeSheetObjectSection(
                        new Presentation.ViewModels.ObjectViewModel<VNCVisioAddIn.Domain.MiscellaneousRow, 
                            VNCVisioAddIn.Presentation.ModelWrappers.MiscellaneousWrapper>(
                            "Update Miscellaneous",
                            VNCVisioAddIn.Domain.MiscellaneousRow.GetRow,
                            VNCVisioAddIn.Domain.MiscellaneousRow.SetRow,
                            VNCVisioAddIn.Domain.ShapeType.Shape),
                        new Presentation.Views.Miscellaneous()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost _shapeProtectionHost = null;

        public static void Protection()
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (_shapeProtectionHost is null) _shapeProtectionHost = new DxThemedWindowHost();

            _shapeProtectionHost.DisplayUserControlInHost(
                "Protection",
                600, 450,
                ShowWindowMode.Modeless_Show,
                    new Presentation.Views.ShapeSheetObjectSection(
                        new Presentation.ViewModels.ObjectViewModel<VNCVisioAddIn.Domain.ProtectionRow,
                            VNCVisioAddIn.Presentation.ModelWrappers.ProtectionWrapper>(
                            "Update Protection",
                            VNCVisioAddIn.Domain.ProtectionRow.GetRow,
                            VNCVisioAddIn.Domain.ProtectionRow.SetRow,
                            VNCVisioAddIn.Domain.ShapeType.Shape),
                        new Presentation.Views.Protection()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost _shapeQuickStyleHost = null;

        public static void QuickStyle()
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (_shapeQuickStyleHost is null) _shapeQuickStyleHost = new DxThemedWindowHost();

            _shapeQuickStyleHost.DisplayUserControlInHost(
                "QuickStyle",
                600, 450,
                ShowWindowMode.Modeless_Show,
                    new Presentation.Views.ShapeSheetObjectSection(
                        new Presentation.ViewModels.ObjectViewModel<VNCVisioAddIn.Domain.QuickStyleRow, 
                            VNCVisioAddIn.Presentation.ModelWrappers.QuickStyleWrapper>(
                            "Update QuickStyle",
                            VNCVisioAddIn.Domain.QuickStyleRow.GetRow,
                            VNCVisioAddIn.Domain.QuickStyleRow.SetRow,
                            VNCVisioAddIn.Domain.ShapeType.Shape),
                        new Presentation.Views.QuickStyle()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost _shapeLayoutHost = null;

        public static void ShapeLayout()
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (_shapeLayoutHost is null) _shapeLayoutHost = new DxThemedWindowHost();

            _shapeLayoutHost.DisplayUserControlInHost(
            "ShapeLayout",
            600, 450,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetObjectSection(
                    new Presentation.ViewModels.ObjectViewModel<VNCVisioAddIn.Domain.ShapeLayoutRow, 
                        VNCVisioAddIn.Presentation.ModelWrappers.ShapeLayoutWrapper>(
                        "Update ShapeLayout",
                        VNCVisioAddIn.Domain.ShapeLayoutRow.GetRow,
                        VNCVisioAddIn.Domain.ShapeLayoutRow.SetRow,
                        VNCVisioAddIn.Domain.ShapeType.Shape),
                    new Presentation.Views.ShapeLayout()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost _shapeTransformHost = null;

        public static void ShapeTransform()
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (_shapeTransformHost is null) _shapeTransformHost = new DxThemedWindowHost();

            _shapeTransformHost.DisplayUserControlInHost(
                "ShapeTransform",
                800, 600,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetObjectSection(
                    new Presentation.ViewModels.ObjectViewModel<VNCVisioAddIn.Domain.ShapeTransformRow, 
                        VNCVisioAddIn.Presentation.ModelWrappers.ShapeTransformWrapper>(
                        "Update ShapeTransform",
                        VNCVisioAddIn.Domain.ShapeTransformRow.GetRow,
                        VNCVisioAddIn.Domain.ShapeTransformRow.SetRow,
                        VNCVisioAddIn.Domain.ShapeType.Shape),
                    new Presentation.Views.ShapeTransform()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost _shapeTextBlockFormatHost = null;

        public static void TextBlockFormat()
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (_shapeTextBlockFormatHost is null) _shapeTextBlockFormatHost = new DxThemedWindowHost();

            _shapeTextBlockFormatHost.DisplayUserControlInHost(
            "TextBlockFormat",
            600, 450,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetObjectSection(
                    new Presentation.ViewModels.ObjectViewModel<VNCVisioAddIn.Domain.TextBlockFormatRow, 
                        VNCVisioAddIn.Presentation.ModelWrappers.TextBlockFormatWrapper>(
                        "Update TextBlockFormat",
                        VNCVisioAddIn.Domain.TextBlockFormatRow.GetRow,
                        VNCVisioAddIn.Domain.TextBlockFormatRow.SetRow,
                        VNCVisioAddIn.Domain.ShapeType.Shape),
                    new Presentation.Views.TextBlockFormat()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost _shapeTextTransformHost = null;

        public static void TextTransform()
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (_shapeTextTransformHost is null) _shapeTextTransformHost = new DxThemedWindowHost();

            _shapeTextTransformHost.DisplayUserControlInHost(
            "TextTransform",
            600, 450,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetObjectSection(
                    new Presentation.ViewModels.ObjectViewModel<VNCVisioAddIn.Domain.TextTransformRow, 
                        VNCVisioAddIn.Presentation.ModelWrappers.TextTransformWrapper>(
                        "Update TextTransform",
                        VNCVisioAddIn.Domain.TextTransformRow.GetRow,
                        VNCVisioAddIn.Domain.TextTransformRow.SetRow,
                        VNCVisioAddIn.Domain.ShapeType.Shape),
                    new Presentation.Views.TextTransform()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }
        public static DxThemedWindowHost _shapeThemePropertiesHost = null;

        public static void ShapeThemeProperties()
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (_shapeThemePropertiesHost is null) _shapeThemePropertiesHost = new DxThemedWindowHost();

            _shapeThemePropertiesHost.DisplayUserControlInHost(
                "Shape ThemeProperties",
                600, 800,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetObjectSection(
                    new Presentation.ViewModels.ObjectViewModel<VNCVisioAddIn.Domain.ThemePropertiesRow, 
                        VNCVisioAddIn.Presentation.ModelWrappers.ThemePropertiesWrapper>(
                        "Update ThemeProperties",
                        VNCVisioAddIn.Domain.ThemePropertiesRow.GetRow,
                        VNCVisioAddIn.Domain.ThemePropertiesRow.SetRow,
                        VNCVisioAddIn.Domain.ShapeType.Shape),
                    new Presentation.Views.ThemeProperties()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #endregion

        #region UI Events - Row Based - Uses RowViewModel

        #region Actions

        public static DxThemedWindowHost _pageActionsHost = null;

        public static void ActionsPage()
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (_pageActionsHost is null) _pageActionsHost = new DxThemedWindowHost();

            _pageActionsHost.DisplayUserControlInHost(
                "Actions (Page)",
                600, 800,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetRowsSection(
                    new Presentation.ViewModels.RowsViewModel<VNCVisioAddIn.Domain.ActionRow, 
                        VNCVisioAddIn.Presentation.ModelWrappers.ActionRowWrapper>(
                        "Update Actions",
                        VNCVisioAddIn.Domain.ActionRow.GetRows,
                        VNCVisioAddIn.Domain.ShapeType.Page),
                    new Presentation.Views.Actions()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost _shapeActionsHost = null;

        public static void ActionsShape()
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (_shapeActionsHost is null) _shapeActionsHost = new DxThemedWindowHost();

            _shapeActionsHost.DisplayUserControlInHost(
                "Actions (Shape)",
                600, 800,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetRowsSection(
                    new Presentation.ViewModels.RowsViewModel<VNCVisioAddIn.Domain.ActionRow, 
                        VNCVisioAddIn.Presentation.ModelWrappers.ActionRowWrapper>(
                        "Update Actions",
                        VNCVisioAddIn.Domain.ActionRow.GetRows,
                        VNCVisioAddIn.Domain.ShapeType.Shape),
                    new Presentation.Views.Actions()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region ActionTags

        public static DxThemedWindowHost _pageActionTagsHost = null;

        public static void ActionTagsPage()
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (_pageActionTagsHost is null) _pageActionTagsHost = new DxThemedWindowHost();

            _pageActionTagsHost.DisplayUserControlInHost(
                "ActionsTags (Page)",
                600, 750,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetRowsSection(
                    new Presentation.ViewModels.RowsViewModel<VNCVisioAddIn.Domain.ActionTagRow, 
                        VNCVisioAddIn.Presentation.ModelWrappers.ActionTagRowWrapper>(
                        "Update ActionTags (Page)",
                        VNCVisioAddIn.Domain.ActionTagRow.GetRows,
                        VNCVisioAddIn.Domain.ShapeType.Page),
                    new Presentation.Views.ActionTags()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost _shapeActionTagsHost = null;

        public static void ActionTagsShape()
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (_shapeActionTagsHost is null) _shapeActionTagsHost = new DxThemedWindowHost();

            _shapeActionTagsHost.DisplayUserControlInHost(
                "ActionsTags (Shape)",
                600, 750,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetRowsSection(
                    new Presentation.ViewModels.RowsViewModel<VNCVisioAddIn.Domain.ActionTagRow, 
                        VNCVisioAddIn.Presentation.ModelWrappers.ActionTagRowWrapper>(
                        "Update ActionTags (Shape)",
                        VNCVisioAddIn.Domain.ActionTagRow.GetRows,
                        VNCVisioAddIn.Domain.ShapeType.Shape),
                    new Presentation.Views.ActionTags()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region Character

        public static DxThemedWindowHost _characterHost = null;

        public static void Character()
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (_characterHost is null) _characterHost = new DxThemedWindowHost();

            _characterHost.DisplayUserControlInHost(
                "Character",
                600, 750,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetRowsSection(
                    new Presentation.ViewModels.RowsViewModel<VNCVisioAddIn.Domain.CharacterRow, 
                        VNCVisioAddIn.Presentation.ModelWrappers.CharacterRowWrapper>(
                        "Update Character",
                        VNCVisioAddIn.Domain.CharacterRow.GetRows,
                        VNCVisioAddIn.Domain.ShapeType.Shape),
                    new Presentation.Views.Character()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region ConnectionPoints

        public static DxThemedWindowHost _connectionPointsHost = null;

        public static void ConnectionPoints()
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (_connectionPointsHost is null) _connectionPointsHost = new DxThemedWindowHost();

            _connectionPointsHost.DisplayUserControlInHost(
                "ConnectionPoints",
                600, 750,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetRowsSection(
                    new Presentation.ViewModels.RowsViewModel<VNCVisioAddIn.Domain.ConnectionPointRow, 
                        VNCVisioAddIn.Presentation.ModelWrappers.ConnectionPointRowWrapper>(
                        "Update ConnectionPoints",
                        VNCVisioAddIn.Domain.ConnectionPointRow.GetRows,
                        VNCVisioAddIn.Domain.ShapeType.Shape),
                    new Presentation.Views.ConnectionPoints()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region Controls

        public static DxThemedWindowHost _controlsHost = null;

        public static void Controls()
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (_controlsHost is null) _controlsHost = new DxThemedWindowHost();

            _controlsHost.DisplayUserControlInHost(
                 "Controls",
                 600, 750,
                 ShowWindowMode.Modeless_Show,
                 new Presentation.Views.ShapeSheetRowsSection(
                     new Presentation.ViewModels.RowsViewModel<VNCVisioAddIn.Domain.ControlsRow, 
                        VNCVisioAddIn.Presentation.ModelWrappers.ControlsRowWrapper>(
                         "Update Controls",
                         VNCVisioAddIn.Domain.ControlsRow.GetRows,
                         VNCVisioAddIn.Domain.ShapeType.Shape),
                     new Presentation.Views.Controls()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region Geometry

        public static DxThemedWindowHost _geometryHost = null;

        public static void Geometry()
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            MessageBox.Show("// TODO(crhodes) - Not Implemented Yet");

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region FillGradientStops

        public static DxThemedWindowHost _fillGradientStopsHost = null;

        public static void GradientStops()
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (_fillGradientStopsHost is null) _fillGradientStopsHost = new DxThemedWindowHost();

            _fillGradientStopsHost.DisplayUserControlInHost(
                 "FillGradientStops",
                 600, 750,
                 ShowWindowMode.Modeless_Show,
                 new Presentation.Views.ShapeSheetRowsSection(
                     new Presentation.ViewModels.RowsViewModel<VNCVisioAddIn.Domain.FillGradientStopRow,
                        VNCVisioAddIn.Presentation.ModelWrappers.FillGradientStopRowWrapper>(
                         "Update FillGradientStops",
                         VNCVisioAddIn.Domain.FillGradientStopRow.GetRows,
                         VNCVisioAddIn.Domain.ShapeType.Shape),
                     new Presentation.Views.FillGradientStops()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region Hyperlinks

        public static DxThemedWindowHost _documentHyperLinksHost = null;

        public static void DocumentHyperlinks()
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (_documentHyperLinksHost is null) _documentHyperLinksHost = new DxThemedWindowHost();

            _documentHyperLinksHost.DisplayUserControlInHost(
                "Hyperlinks (Document)",
                800, 700,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetRowsSection(
                    new Presentation.ViewModels.RowsViewModel<VNCVisioAddIn.Domain.HyperlinkRow,
                        VNCVisioAddIn.Presentation.ModelWrappers.HyperlinkRowWrapper>(
                        "Update Hyperlinks (Document)",
                        VNCVisioAddIn.Domain.HyperlinkRow.GetRows,
                        VNCVisioAddIn.Domain.ShapeType.Document),
                    new Presentation.Views.Hyperlinks()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost _pageHyperLinksHost = null;

        public static void PageHyperlinks()
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (_pageHyperLinksHost is null) _pageHyperLinksHost = new DxThemedWindowHost();

            _pageHyperLinksHost.DisplayUserControlInHost(
                "Hyperlinks (Page)",
                800, 700,
                ShowWindowMode.Modeless_Show,
                    new Presentation.Views.ShapeSheetRowsSection(
                    new Presentation.ViewModels.RowsViewModel<VNCVisioAddIn.Domain.HyperlinkRow,
                        VNCVisioAddIn.Presentation.ModelWrappers.HyperlinkRowWrapper>(
                        "Update Hyperlinks (Page)",
                        VNCVisioAddIn.Domain.HyperlinkRow.GetRows,
                        VNCVisioAddIn.Domain.ShapeType.Page),
                    new Presentation.Views.Hyperlinks()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost _shapeHyperlinksHost = null;

        public static void ShapeHyperlinks()
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (_shapeHyperlinksHost is null) _shapeHyperlinksHost = new DxThemedWindowHost();

            _shapeHyperlinksHost.DisplayUserControlInHost(
                "Hyperlinks (Shape)",
                800, 700,
                ShowWindowMode.Modeless_Show,
                    new Presentation.Views.ShapeSheetRowsSection(
                    new Presentation.ViewModels.RowsViewModel<VNCVisioAddIn.Domain.HyperlinkRow, 
                        VNCVisioAddIn.Presentation.ModelWrappers.HyperlinkRowWrapper>(
                        "Update Hyperlinks (Shape)",
                        VNCVisioAddIn.Domain.HyperlinkRow.GetRows,
                        VNCVisioAddIn.Domain.ShapeType.Shape),
                    new Presentation.Views.Hyperlinks()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region Layers

        public static DxThemedWindowHost _pageLayersHost = null;

        public static void Layers()
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (_pageLayersHost is null) _pageLayersHost = new DxThemedWindowHost();

            _pageLayersHost.DisplayUserControlInHost(
                "Layers (Page)",
                800, 700,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetRowsSection(
                    new Presentation.ViewModels.RowsViewModel<VNCVisioAddIn.Domain.LayerRow, 
                        VNCVisioAddIn.Presentation.ModelWrappers.LayerRowWrapper>(
                        "Update Layers (Page)",
                        VNCVisioAddIn.Domain.LayerRow.GetRows,
                        VNCVisioAddIn.Domain.ShapeType.Page),
                    new Presentation.Views.Layers()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region LineGradientStops

        public static DxThemedWindowHost _lineGradientStopsHost = null;

        public static void LineStops()
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (_lineGradientStopsHost is null) _lineGradientStopsHost = new DxThemedWindowHost();

            _lineGradientStopsHost.DisplayUserControlInHost(
                 "LineGradientStops",
                 600, 750,
                 ShowWindowMode.Modeless_Show,
                 new Presentation.Views.ShapeSheetRowsSection(
                     new Presentation.ViewModels.RowsViewModel<VNCVisioAddIn.Domain.LineGradientStopRow, 
                        VNCVisioAddIn.Presentation.ModelWrappers.LineGradientStopRowWrapper>(
                         "Update LineGradientStops",
                         VNCVisioAddIn.Domain.LineGradientStopRow.GetRows,
                         VNCVisioAddIn.Domain.ShapeType.Shape),
                     new Presentation.Views.LineGradientStops()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region Paragraph

        public static DxThemedWindowHost _paragraphHost = null;

        public static void Paragraph()
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (_paragraphHost is null) _paragraphHost = new DxThemedWindowHost();

            _paragraphHost.DisplayUserControlInHost(
                "Paragraph",
                800, 700,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetRowsSection(
                    new Presentation.ViewModels.RowsViewModel<VNCVisioAddIn.Domain.LayerRow, 
                        VNCVisioAddIn.Presentation.ModelWrappers.LayerRowWrapper>(
                        "Update Paragraph",
                        VNCVisioAddIn.Domain.LayerRow.GetRows,
                        VNCVisioAddIn.Domain.ShapeType.Page),
                    new Presentation.Views.Layers()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region Scratch

        public static DxThemedWindowHost _documentScratchHost = null;

        public static void DocumentScratch()
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (_documentScratchHost is null) _documentScratchHost = new DxThemedWindowHost();

            _documentScratchHost.DisplayUserControlInHost(
                "Scratch (Document)",
                800, 700,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetRowsSection(
                    new Presentation.ViewModels.RowsViewModel<VNCVisioAddIn.Domain.ScratchRow, 
                        VNCVisioAddIn.Presentation.ModelWrappers.ScratchRowWrapper>(
                        "Update Scratch (Document)",
                        VNCVisioAddIn.Domain.ScratchRow.GetRows,
                        VNCVisioAddIn.Domain.ShapeType.Document),
                    new Presentation.Views.Scratch()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost _pageScratchHost = null;

        public static void PageScratch()
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (_pageScratchHost is null) _pageScratchHost = new DxThemedWindowHost();

            _pageScratchHost.DisplayUserControlInHost(
                "Scratch (Page)",
                800, 700,
                ShowWindowMode.Modeless_Show,
                    new Presentation.Views.ShapeSheetRowsSection(
                    new Presentation.ViewModels.RowsViewModel<VNCVisioAddIn.Domain.ScratchRow, 
                        VNCVisioAddIn.Presentation.ModelWrappers.ScratchRowWrapper>(
                        "Update Scratch (Page)",
                        VNCVisioAddIn.Domain.ScratchRow.GetRows,
                        VNCVisioAddIn.Domain.ShapeType.Page),
                    new Presentation.Views.Scratch()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost _shapeScratchHost = null;

        public static void ShapeScratch()
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (_shapeScratchHost is null) _shapeScratchHost = new DxThemedWindowHost();

            _shapeScratchHost.DisplayUserControlInHost(
                "Scratch (Shape)",
                800, 700,
                ShowWindowMode.Modeless_Show,
                    new Presentation.Views.ShapeSheetRowsSection(
                    new Presentation.ViewModels.RowsViewModel<VNCVisioAddIn.Domain.ScratchRow, 
                        VNCVisioAddIn.Presentation.ModelWrappers.ScratchRowWrapper>(
                        "Update Scratch (Shape)",
                        VNCVisioAddIn.Domain.ScratchRow.GetRows,
                        VNCVisioAddIn.Domain.ShapeType.Shape),
                    new Presentation.Views.Scratch()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region ShapeData

        public static DxThemedWindowHost _documentShapeDataHost = null;

        public static void DocumentShapeData()
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (_documentShapeDataHost is null) _documentShapeDataHost = new DxThemedWindowHost();

            _documentShapeDataHost.DisplayUserControlInHost(
                "Shape Data (Document)",
                800, 700,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetRowsSection(
                    new Presentation.ViewModels.RowsViewModel<VNCVisioAddIn.Domain.ShapeDataRow, 
                        VNCVisioAddIn.Presentation.ModelWrappers.ShapeDataRowWrapper>(
                        "Update ShapeData (Document)",
                        VNCVisioAddIn.Domain.ShapeDataRow.GetRows,
                        VNCVisioAddIn.Domain.ShapeType.Document),
                    new Presentation.Views.ShapeData()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost _pageShapeDataHost = null;

        public static void PageShapeData()
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (_pageShapeDataHost is null) _pageShapeDataHost = new DxThemedWindowHost();

            _pageShapeDataHost.DisplayUserControlInHost(
                "Shape Data (Page)",
                800, 700,
                ShowWindowMode.Modeless_Show,
                    new Presentation.Views.ShapeSheetRowsSection(
                    new Presentation.ViewModels.RowsViewModel<VNCVisioAddIn.Domain.ShapeDataRow, 
                        VNCVisioAddIn.Presentation.ModelWrappers.ShapeDataRowWrapper>(
                        "Update ShapeData (Page)",
                        VNCVisioAddIn.Domain.ShapeDataRow.GetRows,
                        VNCVisioAddIn.Domain.ShapeType.Page),
                    new Presentation.Views.ShapeData()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost _shapeShapeDataHost = null;

        public static void ShapeShapeData()
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (_shapeShapeDataHost is null) _shapeShapeDataHost = new DxThemedWindowHost();

            _shapeShapeDataHost.DisplayUserControlInHost(
                "Shape Data (Shape)",
                800, 700,
                ShowWindowMode.Modeless_Show,
                    new Presentation.Views.ShapeSheetRowsSection(
                    new Presentation.ViewModels.RowsViewModel<VNCVisioAddIn.Domain.ShapeDataRow, 
                        VNCVisioAddIn.Presentation.ModelWrappers.ShapeDataRowWrapper>(
                        "Update ShapeData (Shape)",
                        VNCVisioAddIn.Domain.ShapeDataRow.GetRows,
                        VNCVisioAddIn.Domain.ShapeType.Shape),
                    new Presentation.Views.ShapeData()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region Tabs

        public static DxThemedWindowHost _tabsHost = null;

        public static void Tabs()
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (_tabsHost is null) _tabsHost = new DxThemedWindowHost();

            _tabsHost.DisplayUserControlInHost(
                "Tabs",
                800, 700,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetRowsSection(
                    new Presentation.ViewModels.RowsViewModel<VNCVisioAddIn.Domain.TabsRow, 
                        VNCVisioAddIn.Presentation.ModelWrappers.TabRowWrapper>(
                        "Update Tabs",
                        VNCVisioAddIn.Domain.TabsRow.GetRows,
                        VNCVisioAddIn.Domain.ShapeType.Page),
                    new Presentation.Views.Tabs()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region TextFields

        public static DxThemedWindowHost _textFieldsHost = null;

        public static void TextFields()
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (_textFieldsHost is null) _textFieldsHost = new DxThemedWindowHost();

            _textFieldsHost.DisplayUserControlInHost(
                "Tabs",
                800, 700,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetRowsSection(
                    new Presentation.ViewModels.RowsViewModel<VNCVisioAddIn.Domain.TextFieldRow, 
                        VNCVisioAddIn.Presentation.ModelWrappers.TextFieldRowWrapper>(
                        "Update Tabs",
                        VNCVisioAddIn.Domain.TextFieldRow.GetRows,
                        VNCVisioAddIn.Domain.ShapeType.Page),
                    new Presentation.Views.TextFields()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region UserDefinedCells

        public static DxThemedWindowHost _documentUserDefinedCellsHost = null;

        public static void DocumentUserDefinedCells()
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (_documentUserDefinedCellsHost is null) _documentUserDefinedCellsHost = new DxThemedWindowHost();

            _documentUserDefinedCellsHost.DisplayUserControlInHost(
                "User-Defined Cells (Document)",
                800, 700,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.ShapeSheetRowsSection(
                    new Presentation.ViewModels.RowsViewModel<VNCVisioAddIn.Domain.UserDefinedCellRow, 
                        VNCVisioAddIn.Presentation.ModelWrappers.UserDefinedCellRowWrapper>(
                        "Update User-Defined Cells  (Document)",
                        VNCVisioAddIn.Domain.UserDefinedCellRow.GetRows,
                        VNCVisioAddIn.Domain.ShapeType.Document),
                    new Presentation.Views.UserDefinedCells()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost _pageUserDefinedCellsHost = null;

        public static void PageUserDefinedCells()
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (_pageUserDefinedCellsHost is null) _pageUserDefinedCellsHost = new DxThemedWindowHost();

            _pageUserDefinedCellsHost.DisplayUserControlInHost(
                "User-Defined Cells (Page)",
                800, 700,
                ShowWindowMode.Modeless_Show,
                    new Presentation.Views.ShapeSheetRowsSection(
                    new Presentation.ViewModels.RowsViewModel<VNCVisioAddIn.Domain.UserDefinedCellRow, 
                        VNCVisioAddIn.Presentation.ModelWrappers.UserDefinedCellRowWrapper>(
                        "Update User-Defined Cells (Page)",
                        VNCVisioAddIn.Domain.UserDefinedCellRow.GetRows,
                        VNCVisioAddIn.Domain.ShapeType.Page),
                    new Presentation.Views.UserDefinedCells()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost _shapeUserDefinedCellsHost = null;

        public static void ShapeUserDefinedCells()
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (_shapeUserDefinedCellsHost is null) _shapeUserDefinedCellsHost = new DxThemedWindowHost();

            _shapeUserDefinedCellsHost.DisplayUserControlInHost(
                "User-Defined Cells (Shape)",
                800, 700,
                ShowWindowMode.Modeless_Show,
                    new Presentation.Views.ShapeSheetRowsSection(
                    new Presentation.ViewModels.RowsViewModel<VNCVisioAddIn.Domain.UserDefinedCellRow, 
                        VNCVisioAddIn.Presentation.ModelWrappers.UserDefinedCellRowWrapper>(
                        "Update User-Defined Cells (Shape)",
                        VNCVisioAddIn.Domain.UserDefinedCellRow.GetRows,
                        VNCVisioAddIn.Domain.ShapeType.Shape),
                    new Presentation.Views.UserDefinedCells()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #endregion

        #endregion Event Handlers

        protected void UpdateShapeSheetWindowHosts()
        {
            if (_characterHost != null) _characterHost.DeveloperUIMode = Common.DeveloperUIMode;
            if (_connectionPointsHost != null) _connectionPointsHost.DeveloperUIMode = Common.DeveloperUIMode;
            if (_controlsHost != null) _controlsHost.DeveloperUIMode = Common.DeveloperUIMode;
            if (_documentHyperLinksHost != null) _documentHyperLinksHost.DeveloperUIMode = Common.DeveloperUIMode;
            if (_documentPropertiesHost != null) _documentPropertiesHost.DeveloperUIMode = Common.DeveloperUIMode;
            if (_documentScratchHost != null) _documentScratchHost.DeveloperUIMode = Common.DeveloperUIMode;
            if (_documentShapeDataHost != null) _documentShapeDataHost.DeveloperUIMode = Common.DeveloperUIMode;
            if (_documentUserDefinedCellsHost != null) _documentUserDefinedCellsHost.DeveloperUIMode = Common.DeveloperUIMode;
            if (_fillGradientStopsHost != null) _fillGradientStopsHost.DeveloperUIMode = Common.DeveloperUIMode;
            if (_geometryHost != null) _geometryHost.DeveloperUIMode = Common.DeveloperUIMode;
            if (_lineGradientStopsHost != null) _lineGradientStopsHost.DeveloperUIMode = Common.DeveloperUIMode;
            if (_pageActionsHost != null) _pageActionsHost.DeveloperUIMode = Common.DeveloperUIMode;
            if (_pageActionTagsHost != null) _pageActionTagsHost.DeveloperUIMode = Common.DeveloperUIMode;
            if (_pageHyperLinksHost != null) _pageHyperLinksHost.DeveloperUIMode = Common.DeveloperUIMode;
            if (_pageLayersHost != null) _pageLayersHost.DeveloperUIMode = Common.DeveloperUIMode;
            if (_pagePageLayoutHost != null) _pagePageLayoutHost.DeveloperUIMode = Common.DeveloperUIMode;
            if (_pagePagePropertiesHost != null) _pagePagePropertiesHost.DeveloperUIMode = Common.DeveloperUIMode;
            if (_pagePrintPropertiesHost != null) _pagePrintPropertiesHost.DeveloperUIMode = Common.DeveloperUIMode;
            if (_pageRulerAndGridsHost != null) _pageRulerAndGridsHost.DeveloperUIMode = Common.DeveloperUIMode;
            if (_pageScratchHost != null) _pageScratchHost.DeveloperUIMode = Common.DeveloperUIMode;
            if (_pageShapeDataHost != null) _pageShapeDataHost.DeveloperUIMode = Common.DeveloperUIMode;
            if (_pageThemePropertiesHost != null) _pageThemePropertiesHost.DeveloperUIMode = Common.DeveloperUIMode;
            if (_pageUserDefinedCellsHost != null) _pageUserDefinedCellsHost.DeveloperUIMode = Common.DeveloperUIMode;
            if (_paragraphHost != null) _paragraphHost.DeveloperUIMode = Common.DeveloperUIMode;
            if (_shapeActionsHost != null) _shapeActionsHost.DeveloperUIMode = Common.DeveloperUIMode;
            if (_shapeActionTagsHost != null) _shapeActionTagsHost.DeveloperUIMode = Common.DeveloperUIMode;
            if (_shapeAdditionalEffectPropertiesHost != null) _shapeAdditionalEffectPropertiesHost.DeveloperUIMode = Common.DeveloperUIMode;
            if (_shapeAlignmentHost != null) _shapeAlignmentHost.DeveloperUIMode = Common.DeveloperUIMode;
            if (_shapeBevelPropertiesHost != null) _shapeBevelPropertiesHost.DeveloperUIMode = Common.DeveloperUIMode;
            if (_shapeChangeShapeBehaviorHost != null) _shapeChangeShapeBehaviorHost.DeveloperUIMode = Common.DeveloperUIMode;
            if (_shapeEventsHost != null) _shapeEventsHost.DeveloperUIMode = Common.DeveloperUIMode;
            if (_shapeFillFormatHost != null) _shapeFillFormatHost.DeveloperUIMode = Common.DeveloperUIMode;
            if (_shapeGlueInfoHost != null) _shapeGlueInfoHost.DeveloperUIMode = Common.DeveloperUIMode;
            if (_shapeGradientPropertiesHost != null) _shapeGradientPropertiesHost.DeveloperUIMode = Common.DeveloperUIMode;
            if (_shapeGroupPropertiesHost != null) _shapeGroupPropertiesHost.DeveloperUIMode = Common.DeveloperUIMode;
            if (_shapeHyperlinksHost != null) _shapeHyperlinksHost.DeveloperUIMode = Common.DeveloperUIMode;
            if (_shapeImagePropertiesHost != null) _shapeImagePropertiesHost.DeveloperUIMode = Common.DeveloperUIMode;
            if (_shapeLayoutHost != null) _shapeLayoutHost.DeveloperUIMode = Common.DeveloperUIMode;
            if (_shapeLineFormatHost != null) _shapeLineFormatHost.DeveloperUIMode = Common.DeveloperUIMode;
            if (_shapeMiscellaneousHost != null) _shapeMiscellaneousHost.DeveloperUIMode = Common.DeveloperUIMode;
            if (_shapeOneDEndpointsHost != null) _shapeOneDEndpointsHost.DeveloperUIMode = Common.DeveloperUIMode;
            if (_shapeProtectionHost != null) _shapeProtectionHost.DeveloperUIMode = Common.DeveloperUIMode;
            if (_shapeQuickStyleHost != null) _shapeQuickStyleHost.DeveloperUIMode = Common.DeveloperUIMode;
            if (_shapeScratchHost != null) _shapeScratchHost.DeveloperUIMode = Common.DeveloperUIMode;
            if (_shapeShapeDataHost != null) _shapeShapeDataHost.DeveloperUIMode = Common.DeveloperUIMode;
            if (_shapeTextBlockFormatHost != null) _shapeTextBlockFormatHost.DeveloperUIMode = Common.DeveloperUIMode;
            if (_shapeTextTransformHost != null) _shapeTextTransformHost.DeveloperUIMode = Common.DeveloperUIMode;
            if (_shapeThemePropertiesHost != null) _shapeThemePropertiesHost.DeveloperUIMode = Common.DeveloperUIMode;
            if (_shapeThreeDRotationPropertiesHost != null) _shapeThreeDRotationPropertiesHost.DeveloperUIMode = Common.DeveloperUIMode;
            if (_shapeTransformHost != null) _shapeTransformHost.DeveloperUIMode = Common.DeveloperUIMode;
            if (_shapeUserDefinedCellsHost != null) _shapeUserDefinedCellsHost.DeveloperUIMode = Common.DeveloperUIMode;
            if (_tabsHost != null) _tabsHost.DeveloperUIMode = Common.DeveloperUIMode;
            if (_textFieldsHost != null) _textFieldsHost.DeveloperUIMode = Common.DeveloperUIMode;
        }
    }
}
