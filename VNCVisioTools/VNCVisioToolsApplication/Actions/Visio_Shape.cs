using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Reflection;
using System.Text;

using EnvDTE;

using MSVisio = Microsoft.Office.Interop.Visio;
using VNCVisioAddIn = VNC.Visio.VSTOAddIn;

using VNC.Core;

using VNC;
using System.Windows;

namespace VNCVisioToolsApplication.Actions
{
    public class Visio_Shape
    {
        #region Enums, Fields, Properties, Structures

        #endregion

        #region Events

        #endregion

        #region Action Handlers

        #endregion

        #region Helpers

        #endregion

        #region Main Function Routines

        public static void Add_IDandTextSupport_ToSelection()
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            MSVisio.Application app = Common.VisioApplication;

            MSVisio.Selection selection = app.ActiveWindow.Selection;

            foreach (MSVisio.Shape shape in selection)
            {
                Add_IDandTextSupport(shape);
            }
        }

        public static void Add_IDSupport_ToSelection()
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            MSVisio.Application app = Common.VisioApplication;

            MSVisio.Selection selection = app.ActiveWindow.Selection;

            foreach (MSVisio.Shape shape in selection)
            {
                Add_IDSupport(shape);
            }
        }

        public static void Add_TextControl_ToSelection()
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            MSVisio.Application app = Common.VisioApplication;

            MSVisio.Selection selection = app.ActiveWindow.Selection;

            foreach (MSVisio.Shape shape in selection)
            {
                Add_TextTransformControl(shape);
            }
        }

        public static void Add_User_IsPageName()
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            MSVisio.Application app = Common.VisioApplication;

            StringBuilder sb = new StringBuilder();

            MSVisio.Page page = app.ActivePage;

            MSVisio.Selection selection = app.ActiveWindow.Selection;

            Common.WriteToDebugWindow($" Page({page.NameU}) selection.Count: {selection.Count}");

            foreach (MSVisio.Shape shape in selection)
            {
                Common.WriteToDebugWindow($"  Shape({shape.Name})");

                try
                {
                    var isPageName = shape.CellExists["User.IsPageName", 1];
                    var isPageName0 = shape.CellExists["User.IsPageName", 0];

                    if (isPageName != 0)
                    {
                        MSVisio.Cell cell = shape.Cells["User.IsPageName"];

                        cell.ResultIU = 1.0;
                        Common.WriteToDebugWindow($"  Shape({shape.Name}).Cell(Section:{cell.Section} RowName:{cell.RowName} Name:{cell.Name})");
                    }
                    else
                    {
                        shape.AddNamedRow((short)MSVisio.VisSectionIndices.visSectionUser, "IsPageName", 0);
                        shape.Cells["User.isPageName"].ResultIU = 1.0;
                    }

                    UpdatePageNameShape(shape, page.NameU);
                }
                catch (Exception ex)
                {
                    Common.WriteToDebugWindow(ex.ToString(), force: true);
                }
            }

            Common.WriteToDebugWindow($"Visio_Shape.Add_User_IsPageName() End");
        }

        // All string.Format replaced with string interpolation
        public static void AddColorSupportToSelection()
        {
            MSVisio.Application app = Common.VisioApplication;

            MSVisio.Selection selection = app.ActiveWindow.Selection;

            foreach (MSVisio.Shape shape in selection)
            {
                Add_ColorSupport(shape);
            }
        }

        public static void AddHyperLink(MSVisio.Shape shape, string pageName)
        {
            try
            {
                MSVisio.Hyperlink hlink = shape.Hyperlinks.Add();
                hlink.SubAddress = pageName;
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.ToString(), force: true);
            }
        }

        public static void AddHyperlinkToPage_FromShapeText()
        {
            MSVisio.Application app = Common.VisioApplication;

            MSVisio.Selection selection = app.ActiveWindow.Selection;

            foreach (MSVisio.Shape shape in selection)
            {
                string pageName = shape.Text;
                AddHyperLink(shape, pageName);
            }
        }

        public static void ClearConnectionPoints(string tag)
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            MSVisio.Application app = Common.VisioApplication;

            MSVisio.Page page = app.ActivePage;

            MSVisio.Selection selection = app.ActiveWindow.Selection;

            foreach (MSVisio.Shape shape in selection)
            {
                Clear_ConnectionPoints(shape, tag);
            }
        }

        public static void DisplayInfo(MSVisio.Shape shape)
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            var isPageName = shape.CellExists["User.IsPageName", 0];

            if (isPageName != 0)
            {
                MSVisio.Cell cell = shape.Cells["User.IsPageName"];
            }

            Common.WriteToDebugWindow($"   Shape(ID:{shape.ID}  Name:{shape.Name}  Text:>{shape.Text}<)");
        }

        public static void GatherInfo()
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            MSVisio.Application app = Common.VisioApplication;

            StringBuilder sb = new StringBuilder();

            MSVisio.Page page = app.ActivePage;

            MSVisio.Selection selection = app.ActiveWindow.Selection;

            foreach (MSVisio.Shape shape in selection)
            {
                DisplayInfo(shape);
            }
        }

        public static void HandleShapeAdded(MSVisio.Shape shape)
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");
            var isPageName = shape.CellExists["User.IsPageName", 0];

            if (isPageName != 0)
            {
                MSVisio.Cell cell = shape.Cells["User.IsPageName"];

                if (cell.ResultIU > 0)
                {
                    MSVisio.Application app = Common.VisioApplication;
                    MSVisio.Page page = app.ActivePage;
                    shape.Text = page.NameU;
                }
            }
        }

        public static void LinkShapeToPage(MSVisio.Application app, string doc, string page, string shape, string shapeu, String[] args)
        {
            string pageLevel = args[0];
            string separator = "";

            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}() PageLevel:{pageLevel}");

            MSVisio.Shape activeShape = app.ActivePage.Shapes[shape];
            Common.WriteToDebugWindow($"  Shape(Name:{activeShape.Name}  Text:{activeShape.Text}");

            if (pageLevel.Length > 0)
            {
                separator = "-";
            }

            string pageName = $"{pageLevel}{separator}{activeShape.Characters.TextAsString.Replace("\n", " ")}";

            MSVisio.Hyperlink newHyperLink = activeShape.AddHyperlink();
            newHyperLink.SubAddress = pageName;
        }

        public static void ListInvocationsInMethod(MSVisio.Application app, string doc, string page, string shape, string shapeu, string[] array)
        {


        }

        public static void ListMethodsInClass(MSVisio.Application app, string doc, string page, string shape, string shapeu, string[] array)
        {


        }
        
        public static void MakeLinkableMaster()
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            MSVisio.Application app = Common.VisioApplication;

            MSVisio.Selection selection = app.ActiveWindow.Selection;

            foreach (MSVisio.Shape shape in selection)
            {
                MakeLinkableMaster(shape);
            }
        }

        public static void MoveToBackgroundLayer()
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            MSVisio.Application app = Common.VisioApplication;
            MSVisio.Page currentPage = app.ActivePage;
            MSVisio.Layer backgroundLayer = null;
            string layerLock = "1";

            try
            {
                backgroundLayer = currentPage.Layers["Background"];
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.ToString(), force: true);
            }

            if (backgroundLayer == null)
            {
                backgroundLayer = Visio_Page.AddLayer(currentPage, "Background");
            }
            else
            {
                layerLock = backgroundLayer.CellsC[(short)MSVisio.VisCellIndices.visLayerLock].FormulaU;
            }

            MSVisio.Selection selection = app.ActiveWindow.Selection;

            foreach (MSVisio.Shape shape in selection)
            {
                backgroundLayer.Add(shape, 1);
            }

            backgroundLayer.CellsC[(short)MSVisio.VisCellIndices.visLayerLock].FormulaU = layerLock;
        }

        public static void SetMargins(string points)
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            MSVisio.Application app = Common.VisioApplication;

            MSVisio.Selection selection = app.ActiveWindow.Selection;

            foreach (MSVisio.Shape shape in selection)
            {
                SetAllMargins(shape, points);
            }
        }

        public static void UpdatePageNameShape(MSVisio.Shape shape, string pageName)
        {
            var isPageName = shape.CellExistsU["User.IsPageName", 0];

            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}({shape.Name}  isPageName:{isPageName})");

            if (isPageName != 0)
            {
                MSVisio.Cell cell = shape.CellsU["User.IsPageName"];

                Common.WriteToDebugWindow($"    Shape({shape.Name}).Cell(Section:{cell.Section} RowName:{cell.RowName} Name:{cell.Name} Value:{cell.ResultIU})");

                if (cell.ResultIU > 0)
                {
                    shape.Text = pageName;
                }
            }
        }

        internal static void Add_ConnectionPoints(List<VNCVisioAddIn.Domain.ConnectionPointRow> connectionPoints)
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            MSVisio.Application app = Common.VisioApplication;

            MSVisio.Page page = app.ActivePage;

            MSVisio.Selection selection = app.ActiveWindow.Selection;

            foreach (MSVisio.Shape shape in selection)
            {
                Add_ConnectionPoints(shape, connectionPoints);
            }
        }

        static void Add_Connection_Row(MSVisio.Shape shape,
            VNCVisioAddIn.Domain.ConnectionPointRow connectionPoint)
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            short sectionConnectionPts = (short)MSVisio.VisSectionIndices.visSectionConnectionPts;
            short tagConnectionPts = (short)MSVisio.VisRowTags.visTagCnnctPt;
            short tagConnectionNamed = (short)MSVisio.VisRowTags.visTagCnnctNamedABCD;

            //short rowID = shape.AddRow(sectionConnectionPts, (short)VisRowIndices.visRowLast, tagConnectionPts);
            short rowID = shape.AddNamedRow(sectionConnectionPts, connectionPoint.Name, tagConnectionNamed);


            //shape.CellsSRC[
            //    (short)MSVisio.VisSectionIndices.visSectionConnectionPts,
            //    rowID,
            //    (short)MSVisio.VisCellIndices.visCnnctD].FormulaU = "Top";

            //shape.CellsSRC[
            //    (short)MSVisio.VisSectionIndices.visSectionConnectionPts,
            //    rowID,
            //    (short)MSVisio.VisCellIndices.visCnnctX].RowNameU = connectionPoint.Name;

            shape.CellsSRC[
                sectionConnectionPts,
                rowID,
                (short)MSVisio.VisCellIndices.visCnnctX].FormulaU = connectionPoint.X;

            shape.CellsSRC[
                sectionConnectionPts,
                 rowID,
                (short)MSVisio.VisCellIndices.visCnnctY].FormulaU = connectionPoint.Y;

            shape.CellsSRC[
                sectionConnectionPts,
                rowID,
                (short)MSVisio.VisCellIndices.visCnnctDirX].FormulaU = connectionPoint.DirX;

            shape.CellsSRC[
                sectionConnectionPts,
                 rowID,
                (short)MSVisio.VisCellIndices.visCnnctDirY].FormulaU = connectionPoint.DirY;

            shape.CellsSRC[
                sectionConnectionPts,
                 rowID,
                (short)MSVisio.VisCellIndices.visCnnctType].FormulaU = connectionPoint.Type;

            shape.CellsSRC[
                sectionConnectionPts,
                 rowID,
                (short)MSVisio.VisCellIndices.visCnnctD].FormulaU = connectionPoint.D.WrapInDblQuotes();
        }

        static void Add_ConnectionPoints(MSVisio.Shape shape, List<VNCVisioAddIn.Domain.ConnectionPointRow> connectionPoints)
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");
            // TODO(crhodes)
            // Add a remove Connection Points method to clear things out.

            try
            {
                foreach (var connectionPoint in connectionPoints)
                {
                    Add_Connection_Row(shape, connectionPoint);
                }
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.ToString(), force: true);
            }
        }

        static void Clear_ConnectionPoints(MSVisio.Shape shape, string tag)
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            short sectionConnection = (short)MSVisio.VisSectionIndices.visSectionConnectionPts;

            try
            {
                short exists = shape.SectionExists[sectionConnection, 0];

                if (exists != 0)
                {
                    int rows = shape.RowCount[sectionConnection];

                    MSVisio.Section connectionPoints = shape.Section[sectionConnection];

                    for (short row = 0; row < rows; row++)
                    {
                        MSVisio.Cell cnnctX = shape.CellsSRC[sectionConnection, row, (short)MSVisio.VisCellIndices.visCnnctX];
                        MSVisio.Cell cnnctY = shape.CellsSRC[sectionConnection, row, (short)MSVisio.VisCellIndices.visCnnctY];
                        MSVisio.Cell cnnctDirX = shape.CellsSRC[sectionConnection, row, (short)MSVisio.VisCellIndices.visCnnctDirX];
                        MSVisio.Cell cnnctA = shape.CellsSRC[sectionConnection, row, (short)MSVisio.VisCellIndices.visCnnctA];
                        MSVisio.Cell cnnctDirY = shape.CellsSRC[sectionConnection, row, (short)MSVisio.VisCellIndices.visCnnctDirY];
                        MSVisio.Cell cnnctB = shape.CellsSRC[sectionConnection, row, (short)MSVisio.VisCellIndices.visCnnctB];
                        MSVisio.Cell cnnctType = shape.CellsSRC[sectionConnection, row, (short)MSVisio.VisCellIndices.visCnnctType];
                        MSVisio.Cell cnnctC = shape.CellsSRC[sectionConnection, row, (short)MSVisio.VisCellIndices.visCnnctC];
                        MSVisio.Cell cnnctD = shape.CellsSRC[sectionConnection, row, (short)MSVisio.VisCellIndices.visCnnctD];

                        var cpX = cnnctX.FormulaU;
                        var cpY = cnnctY.FormulaU;
                        var cpDirX = cnnctDirX.FormulaU;
                        var cpA = cnnctA.FormulaU;
                        var cpDirY = cnnctDirY.FormulaU;
                        var cpB = cnnctB.FormulaU;
                        var cpType = cnnctType.FormulaU;
                        var cpC = cnnctC.FormulaU;
                        var cpD = cnnctD.FormulaU;

                        string value = cnnctD.FormulaU;

                        if (value.Contains(tag))
                        {
                            shape.DeleteRow(sectionConnection, row);
                        }
                    }

                    switch (tag)
                    {
                        case "Lefts":
                            break;
                        case "Tops":
                            break;
                        case "Rights":
                            break;
                        case "Bottoms":
                            break;
                        case "All":
                            shape.DeleteSection((short)MSVisio.VisSectionIndices.visSectionConnectionPts);
                            break;
                        default:
                            MessageBox.Show($"Unknown tag: {tag}");
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.ToString(), force: true);
            }
        }
        private static void MakeLinkableMaster(MSVisio.Shape shape)
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            try
            {
                VNCVisioAddIn.Helpers.Validate_Action_SectionExists(shape);

                VNCVisioAddIn.Helpers.Delete_Section_Row(shape, MSVisio.VisSectionIndices.visSectionAction, "Actions", "CreateActivityPage");
                VNCVisioAddIn.Helpers.Delete_Section_Row(shape, MSVisio.VisSectionIndices.visSectionAction, "Actions", "LinkShapeToPage");

                VNCVisioAddIn.Helpers.Add_ActionSection_Row(shape,
                    "CreateActivityPage",
                    "=RUNADDONWARGS(\"QueueMarkerEvent\",\"CreatePageForShape,,, Page % 20Base\")",
                    "Create Page for Shape");
                VNCVisioAddIn.Helpers.Add_ActionSection_Row(shape,
                    "LinkShapeToPage",
                    "=RUNADDONWARGS(\"QueueMarkerEvent\",\"LinkShapeToPage, \")",
                    "Link Shape to Page");

                VNCVisioAddIn.Helpers.Validate_Prop_SectionExists(shape);

                VNCVisioAddIn.Helpers.Delete_Section_Row(shape, MSVisio.VisSectionIndices.visSectionProp, "Prop", "HyperLink");
                VNCVisioAddIn.Helpers.Delete_Section_Row(shape, MSVisio.VisSectionIndices.visSectionProp, "Prop", "ReturnLink");
                VNCVisioAddIn.Helpers.Delete_Section_Row(shape, MSVisio.VisSectionIndices.visSectionProp, "Prop", "PageName");
                VNCVisioAddIn.Helpers.Delete_Section_Row(shape, MSVisio.VisSectionIndices.visSectionProp, "Prop", "ExternalLink");
                VNCVisioAddIn.Helpers.Delete_Section_Row(shape, MSVisio.VisSectionIndices.visSectionProp, "Prop", "HyperLinkPrefix");

                VNCVisioAddIn.Helpers.Add_Prop_Row(shape, "PageName", "PageName", (short)MSVisio.VisCellVals.visPropTypeString, null, "<page name>".WrapInDblQuotes(), prompt: "Enter Page Name", ask: "TRUE");
                VNCVisioAddIn.Helpers.Add_Prop_Row(shape, "HyperLink", "HyperLink", (short)MSVisio.VisCellVals.visPropTypeString, null, "");
                VNCVisioAddIn.Helpers.Add_Prop_Row(shape, "ReturnLink", "ReturnLink", (short)MSVisio.VisCellVals.visPropTypeString, null, "Page Shapes.vssx,PageLink Arrow Left".WrapInDblQuotes());
                VNCVisioAddIn.Helpers.Add_Prop_Row(shape, "ExternalLink", "ExternalLink", (short)MSVisio.VisCellVals.visPropTypeString, null, "");
                VNCVisioAddIn.Helpers.Add_Prop_Row(shape, "HyperLinkPrefix", "HyperLinkPrefix", (short)MSVisio.VisCellVals.visPropTypeString, null, "");

                // For now assume the shape does not have any hyperlinks we care about.

                shape.DeleteSection((short)MSVisio.VisSectionIndices.visSectionHyperlink);

                VNCVisioAddIn.Helpers.Validate_HyperLink_SectionExists(shape);

                MSVisio.Hyperlink newHyperLink = shape.AddHyperlink();

                // This doesn't work as the value is treated as a string.
                //newHyperLink.SubAddress = "Prop.HyperLink";

                // Need to go at it as a CellSRC

                shape.CellsSRC[
                    (short)MSVisio.VisSectionIndices.visSectionHyperlink,
                    (short)MSVisio.VisRowIndices.visRow1stHyperlink,
                    (short)MSVisio.VisCellIndices.visHLinkSubAddress].FormulaU = "GUARD(Prop.HyperLink)";

                // This creates a section that we could update, but the shape.Characters also adds a TextField row
                // we don't need two.
                //Validate_TextField_SectionExists(shape);

                // Not sure how to go about this.  Macro recorder shows this

                MSVisio.Characters characters = shape.Characters;
                characters.AddCustomFieldU("GUARD(Prop.PageName)", (short)MSVisio.VisFieldFormats.visFmtNumGenNoUnits);

                // TODO(crhodes)
                // Need to protect the Text so not accidentally overridden.
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.ToString(), force:true);
            }
        }
        
        #endregion

        #region Private Methods

        private static void Add_ColorSupport(MSVisio.Shape shape)
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");
            // Have to add these in the right order as there are some dependencies

            string value = string.Empty;

            value = "AliceBlue; AntiqueWhite; Aqua; Aquamarine; Azure; Beige; Bisque; Black; BlanchedAlmond; Blue; BlueViolet; Brown; BurlyWood; CadetBlue; Chartreuse; Chocolate; Coral; CornflowerBlue; Cornsilk; Crimson; Cyan; DarkBlue; DarkCyan; DarkGoldenrod; DarkGray; DarkGreen; DarkKhaki; DarkMagenta; DarkOliveGreen; DarkOrange; DarkOrchid; DarkRed; DarkSalmon; DarkSeaGreen; DarkSlateBlue; DarkSlateGray; DarkTurquoise; DarkViolet; DeepPink; DeepSkyBlue; DimGray; DodgerBlue; Firebrick; FloralWhite; ForestGreen; Fuchsia; Gainsboro; GhostWhite; Gold; Goldenrod; Gray; Green; GreenYellow; Honeydew; HotPink; IndianRed; Indigo; Ivory; Khaki; Lavender; LavenderBlush; LawnGreen; LemonChiffon; LightBlue; LightCoral; LightCyan; LightGoldenrodYellow; LightGreen; LightGray; LightPink; LightSalmon; LightSeaGreen; LightSkyBlue; LightSlateGray; LightSteelBlue; LightYellow; Lime; LimeGreen; Linen; Magenta; Maroon; MediumAquamarine; MediumBlue; MediumOrchid; MediumPurple; MediumSeaGreen; MediumSlateBlue; MediumSpringGreen; MediumTurquoise; MediumVioletRed; MidnightBlue; MintCream; MistyRose; Moccasin; NavajoWhite; Navy; OldLace; Olive; OliveDrab; Orange; OrangeRed; Orchid; PaleGoldenrod; PaleGreen; PaleTurquoise; PaleVioletRed; PapayaWhip; PeachPuff; Peru; Pink; Plum; PowderBlue; Purple; Red; RosyBrown; RoyalBlue; SaddleBrown; Salmon; SandyBrown; SeaGreen; SeaShell; Sienna; Silver; SkyBlue; SlateBlue; SlateGray; Snow; SpringGreen; SteelBlue; Tan; Teal; Thistle; Tomato; Turquoise; Violet; Wheat; White; WhiteSmoke; Yellow; YellowGreen";
            value = string.Format("\"{0}\"", value);
            VNCVisioAddIn.Helpers.Add_User_Row(shape, "colorNames", value);

            value = "RGB(240, 248, 255); RGB(250, 235, 215); RGB(0, 255, 255); RGB(127, 255, 212); RGB(240, 255, 255); RGB(245, 245, 220); RGB(255, 228, 196); RGB(0, 0, 0); RGB(255, 235, 205); RGB(0, 0, 255); RGB(138, 43, 226); RGB(165, 42, 42); RGB(222, 184, 135); RGB(95, 158, 160); RGB(127, 255, 0); RGB(210, 105, 30); RGB(255, 127, 80); RGB(100, 149, 237); RGB(255, 248, 220); RGB(220, 20, 60); RGB(0, 255, 255); RGB(0, 0, 139); RGB(0, 139, 139); RGB(184, 134, 11); RGB(169, 169, 169); RGB(0, 100, 0); RGB(189, 183, 107); RGB(139, 0, 139); RGB(85, 107, 47); RGB(255, 140, 0); RGB(153, 50, 204); RGB(139, 0, 0); RGB(233, 150, 122); RGB(143, 188, 139); RGB(72, 61, 139); RGB(47, 79, 79); RGB(0, 206, 209); RGB(148, 0, 211); RGB(255, 20, 147); RGB(0, 191, 255); RGB(105, 105, 105); RGB(30, 144, 255); RGB(178, 34, 34); RGB(255, 250, 240); RGB(34, 139, 34); RGB(255, 0, 255); RGB(220, 220, 220); RGB(248, 248, 255); RGB(255, 215, 0); RGB(218, 165, 32); RGB(128, 128, 128); RGB(0, 128, 0); RGB(173, 255, 47); RGB(240, 255, 240); RGB(255, 105, 180); RGB(205, 92, 92); RGB(75, 0, 130); RGB(255, 255, 240); RGB(240, 230, 140); RGB(230, 230, 250); RGB(255, 240, 245); RGB(124, 252, 0); RGB(255, 250, 205); RGB(173, 216, 230); RGB(240, 128, 128); RGB(224, 255, 255); RGB(250, 250, 210); RGB(144, 238, 144); RGB(211, 211, 211); RGB(255, 182, 193); RGB(255, 160, 122); RGB(32, 178, 170); RGB(135, 206, 250); RGB(119, 136, 153); RGB(176, 196, 222); RGB(255, 255, 224); RGB(0, 255, 0); RGB(50, 205, 50); RGB(250, 240, 230); RGB(255, 0, 255); RGB(128, 0, 0); RGB(102, 205, 170); RGB(0, 0, 205); RGB(186, 85, 211); RGB(147, 112, 219); RGB(60, 179, 113); RGB(123, 104, 238); RGB(0, 250, 154); RGB(72, 209, 204); RGB(199, 21, 133); RGB(25, 25, 112); RGB(245, 255, 250); RGB(255, 228, 225); RGB(255, 228, 181); RGB(255, 222, 173); RGB(0, 0, 128); RGB(253, 245, 230); RGB(128, 128, 0); RGB(107, 142, 35); RGB(255, 165, 0); RGB(255, 69, 0); RGB(218, 112, 214); RGB(238, 232, 170); RGB(152, 251, 152); RGB(175, 238, 238); RGB(219, 112, 147); RGB(255, 239, 213); RGB(255, 218, 185); RGB(205, 133, 63); RGB(255, 192, 203); RGB(221, 160, 221); RGB(176, 224, 230); RGB(128, 0, 128); RGB(255, 0, 0); RGB(188, 143, 143); RGB(65, 105, 225); RGB(139, 69, 19); RGB(250, 128, 114); RGB(244, 164, 96); RGB(46, 139, 87); RGB(255, 245, 238); RGB(160, 82, 45); RGB(192, 192, 192); RGB(135, 206, 235); RGB(106, 90, 205); RGB(112, 128, 144); RGB(255, 250, 250); RGB(0, 255, 127); RGB(70, 130, 180); RGB(210, 180, 140); RGB(0, 128, 128); RGB(216, 191, 216); RGB(255, 99, 71); RGB(64, 224, 208); RGB(238, 130, 238); RGB(245, 222, 179); RGB(255, 255, 255); RGB(245, 245, 245); RGB(255, 255, 0); RGB(154, 205, 50)";
            value = string.Format("\"{0}\"", value);
            VNCVisioAddIn.Helpers.Add_User_Row(shape, "colorValues", value);

            VNCVisioAddIn.Helpers.Add_Prop_Row(shape, "Color", "Color", (short)MSVisio.VisCellVals.visPropTypeListFix, "User.colorNames", "", "", "");

            value = "INDEX(LOOKUP(Prop.Color,User.colorNames),User.colorValues)";
            VNCVisioAddIn.Helpers.Add_User_Row(shape, "Color", value);
        }

        private static void Add_IDandTextSupport(MSVisio.Shape shape)
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            VNCVisioAddIn.Helpers.Validate_Prop_SectionExists(shape);

            VNCVisioAddIn.Helpers.Add_Prop_Row(shape, rowName: "ID", label: "ID", type: (short)MSVisio.VisCellVals.visPropTypeString, format: null, value: "Xnnn".WrapInDblQuotes());
            VNCVisioAddIn.Helpers.Add_Prop_Row(shape, "ShowID", "Show ID", (short)MSVisio.VisCellVals.visPropTypeBool, null, "TRUE".WrapInDblQuotes());

            VNCVisioAddIn.Helpers.Add_Prop_Row(shape, "Text", "Text", (short)MSVisio.VisCellVals.visPropTypeString, null, "<text>".WrapInDblQuotes());
            VNCVisioAddIn.Helpers.Add_Prop_Row(shape, "TextLeft", "Text Left", (short)MSVisio.VisCellVals.visPropTypeString, null, "<left text>".WrapInDblQuotes());
            VNCVisioAddIn.Helpers.Add_Prop_Row(shape, "TextTop", "Text Top", (short)MSVisio.VisCellVals.visPropTypeString, null, "<top text>".WrapInDblQuotes());
            VNCVisioAddIn.Helpers.Add_Prop_Row(shape, "TextRight", "Text Right", (short)MSVisio.VisCellVals.visPropTypeString, null, "<right text>".WrapInDblQuotes());
            VNCVisioAddIn.Helpers.Add_Prop_Row(shape, "TextBottom", "Text Bottom", (short)MSVisio.VisCellVals.visPropTypeString, null, "<bottom text>".WrapInDblQuotes());

            VNCVisioAddIn.Helpers.Add_Prop_Row(shape, "ShowLeftText", "Show Left Text", (short)MSVisio.VisCellVals.visPropTypeBool, null, "FALSE".WrapInDblQuotes());
            VNCVisioAddIn.Helpers.Add_Prop_Row(shape, "ShowTopText", "Show Top Text", (short)MSVisio.VisCellVals.visPropTypeBool, null, "FALSE".WrapInDblQuotes());
            VNCVisioAddIn.Helpers.Add_Prop_Row(shape, "ShowRightText", "Show Right Text", (short)MSVisio.VisCellVals.visPropTypeBool, null, "FALSE".WrapInDblQuotes());
            VNCVisioAddIn.Helpers.Add_Prop_Row(shape, "ShowBottomText", "Show Bottom Text", (short)MSVisio.VisCellVals.visPropTypeBool, null, "FALSE".WrapInDblQuotes());

            VNCVisioAddIn.Helpers.Add_Prop_Row(shape, "SizeText", "Size Text", (short)MSVisio.VisCellVals.visPropTypeNumber, "0.0".WrapInDblQuotes(), "1.0");
            VNCVisioAddIn.Helpers.Add_Prop_Row(shape, "SizeLeftText", "Size Left Text", (short)MSVisio.VisCellVals.visPropTypeNumber, "0.0".WrapInDblQuotes(), "1.0");
            VNCVisioAddIn.Helpers.Add_Prop_Row(shape, "SizeTopText", "Size Top Text", (short)MSVisio.VisCellVals.visPropTypeNumber, "0.0".WrapInDblQuotes(), "1.0");
            VNCVisioAddIn.Helpers.Add_Prop_Row(shape, "SizeRightText", "Size Right Text", (short)MSVisio.VisCellVals.visPropTypeNumber, "0.0".WrapInDblQuotes(), "1.0");
            VNCVisioAddIn.Helpers.Add_Prop_Row(shape, "SizeBottomText", "Size Bottom Text", (short)MSVisio.VisCellVals.visPropTypeNumber, value: "1.0", format: "0.0".WrapInDblQuotes());
            VNCVisioAddIn.Helpers.Add_Prop_Row(shape, "SizeIDText", "Size ID Text", (short)MSVisio.VisCellVals.visPropTypeNumber, value: "1.0", format: "0.0".WrapInDblQuotes());

            VNCVisioAddIn.Helpers.Add_Prop_Row(shape, "GroupName", "Group Name", (short)MSVisio.VisCellVals.visPropTypeString, null, "<group name>".WrapInDblQuotes());
            VNCVisioAddIn.Helpers.Add_Prop_Row(shape, "TabColor", "Tab Color (RGB)", (short)MSVisio.VisCellVals.visPropTypeString, null, "RGB(128,128,128)");
        }
        private static void Add_IDSupport(MSVisio.Shape shape)
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            VNCVisioAddIn.Helpers.Validate_Prop_SectionExists(shape);

            VNCVisioAddIn.Helpers.Add_Prop_Row(shape, "ID", "ID", (short)MSVisio.VisCellVals.visPropTypeString, null, "000".WrapInDblQuotes());
            VNCVisioAddIn.Helpers.Add_Prop_Row(shape, "ShowID", "Show ID", (short)MSVisio.VisCellVals.visPropTypeBool, null, "TRUE".WrapIn("\""));
            VNCVisioAddIn.Helpers.Add_Prop_Row(shape, "Text", "Text", (short)MSVisio.VisCellVals.visPropTypeString, null, "<text>".WrapInDblQuotes());
            VNCVisioAddIn.Helpers.Add_Prop_Row(shape, "SizeText", "Size Text", (short)MSVisio.VisCellVals.visPropTypeNumber, format: "0.0".WrapInDblQuotes(), value: "1.0");
        }

        // TODO(crhodes):
        // This section should be reviewed and if appropriate lifted out into the ShapeEditor 
        // and associated configuration file

        private static void Add_TextTransformControl(MSVisio.Shape shape)
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            VNCVisioAddIn.Helpers.Populate_Controls_Section(shape,
                "Width*0.5",
                "Height*0.5",
                "Controls.Row_1",
                "Controls.Row_1.Y",
                "0",
                "0",
                "TRUE",
                "Reposition Text");

            VNCVisioAddIn.Helpers.Set_TextXForm_Section(shape,
                "GUARD(Width*2)",
                "GUARD(Height*2)",
                "GUARD(Controls.Row_1)",
                "GUARD(Controls.Row_1.Y)",
                "TxtWidth*0.5",
                "TxtHeight*0.5",
                "0 deg"
                );

            //Set_Paragraph_Section()
            // TODO(crhodes)
            // DO we need to call the first method
            //VNCVisioAddIn.Domain.TextBlockFormat textBlock = new VNCVisioAddIn.Domain.TextBlockFormat();
            //VNCVisioAddIn.Domain.TextBlockFormat.Set_TextBlockFormat_Section(shape, textBlock);

            VNCVisioAddIn.Domain.TextBlockFormatRow.SetRow(shape);
        }

        private static void SetAllMargins(MSVisio.Shape shape, string points)
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            //VNCVisioAddIn.Domain.TextBlockFormat.Validate_TextBlockFormat_SectionExists(shape);

            VNCVisioAddIn.Domain.TextBlockFormatRow.SetMargins(shape, points, points, points, points);
        }

        private static void SetMargins(MSVisio.Shape shape, string leftPoints, string topPoints, string rightPoints, string bottomPoints)
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            //VNCVisioAddIn.Domain.TextBlockFormat.Validate_TextBlockFormat_SectionExists(shape);

            VNCVisioAddIn.Domain.TextBlockFormatRow.SetMargins(shape, leftPoints, topPoints, rightPoints, bottomPoints);
        }

        #endregion

        #region Utility Routines

        public static void MoveToBackgroundLayer(MSVisio.Application app, string doc, string page, string shape, string shapeu)
        {


        }

        public static void UpdateTextSections(VNCVisioAddIn.Domain.TextBlockFormatRow textBlockFormat)
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            MSVisio.Application app = Common.VisioApplication;

            MSVisio.Page page = app.ActivePage;

            MSVisio.Selection selection = app.ActiveWindow.Selection;

            foreach (MSVisio.Shape shape in selection)
            {
                VNCVisioAddIn.Domain.TextBlockFormatRow.SetRow(shape, textBlockFormat);
            }
        }

  

        #endregion

        public static bool HasTextTransformSection(MSVisio.Shape shape)
        {
            bool result = false;

            MSVisio.Section section = shape.Section[(short)MSVisio.VisSectionIndices.visSectionObject];
            MSVisio.Row sectionRow = section[(short)MSVisio.VisRowIndices.visRowTextXForm];

            return result;
        }

        internal static ObservableCollection<VNCVisioAddIn.Domain.ConnectionPointRow> Get_ConnectionPointRows(MSVisio.Shape shape)
        {
            var rows = new ObservableCollection<VNCVisioAddIn.Domain.ConnectionPointRow>();

            MSVisio.Section section = shape.Section[(short)MSVisio.VisSectionIndices.visSectionConnectionPts];

            var rowCount = section.Count;

            for (short i = 0; i < rowCount; i++)
            {
                VNCVisioAddIn.Domain.ConnectionPointRow connectionPointRow = new VNCVisioAddIn.Domain.ConnectionPointRow();

                var row = section[i];

                connectionPointRow.Name = row.NameU;

                connectionPointRow.X = row[(short)MSVisio.VisCellIndices.visCnnctX].FormulaU;
                connectionPointRow.Y = row[(short)MSVisio.VisCellIndices.visCnnctY].FormulaU;
                connectionPointRow.DirX = row[(short)MSVisio.VisCellIndices.visCnnctDirX].FormulaU;
                connectionPointRow.A = row[(short)MSVisio.VisCellIndices.visCnnctA].FormulaU;
                connectionPointRow.DirY = row[(short)MSVisio.VisCellIndices.visCnnctDirY].FormulaU;
                connectionPointRow.B = row[(short)MSVisio.VisCellIndices.visCnnctB].FormulaU;
                connectionPointRow.Type = row[(short)MSVisio.VisCellIndices.visCnnctType].FormulaU;
                connectionPointRow.C = row[(short)MSVisio.VisCellIndices.visCnnctC].FormulaU;
                connectionPointRow.D = row[(short)MSVisio.VisCellIndices.visCnnctD].FormulaU;

                rows.Add(connectionPointRow);
            }

            return rows;
        }
    }
}
