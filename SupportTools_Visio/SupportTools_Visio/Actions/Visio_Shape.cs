using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Reflection;
using System.Text;
using System.Windows;

using Microsoft.Office.Interop.Visio;

using SupportTools_Visio.Domain;

//using static VNC.Helper;
using VNC;
using VNC.Core;

using Visio = Microsoft.Office.Interop.Visio;
using VNCVisioAddIn = VNC.Visio.VSTOAddIn;

namespace SupportTools_Visio.Actions
{
    class Visio_Shape
    {

        #region Main Function Routines

        public static void ListInvocationsInMethod(Visio.Application app, string doc, string page, string shape, string shapeu, string[] array)
        {
            

        }

        public static void ListMethodsInClass(Visio.Application app, string doc, string page, string shape, string shapeu, string[] array)
        {


        }

        public static void Add_User_IsPageName()
        {
            VNCVisioAddIn.Common.DisplayInDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            Visio.Application app = Globals.ThisAddIn.Application;

            StringBuilder sb = new StringBuilder();

            Visio.Page page = app.ActivePage;

            Visio.Selection selection = app.ActiveWindow.Selection;

            VNCVisioAddIn.Common.DisplayInDebugWindow(string.Format(" Page({0}) selection.Count: {1}", page.NameU, selection.Count));

            //for (int i = 0; i < selection.Count; i++)
            //{
            //    var item = selection[i];

            //    VNCVisioAddIn.Common.DisplayInDebugWindow(string.Format("  Shape({0})", item.Name));
            //}

            foreach (Visio.Shape shape in selection)
            {
                VNCVisioAddIn.Common.DisplayInDebugWindow(string.Format("  Shape({0})", shape.Name));

                try
                {
                    var isPageName = shape.CellExists["User.IsPageName", 1];
                    var isPageName0 = shape.CellExists["User.IsPageName", 0];

                    if (isPageName != 0)
                    {
                        Visio.Cell cell = shape.Cells["User.IsPageName"];

                        cell.ResultIU = 1.0;
                        VNCVisioAddIn.Common.DisplayInDebugWindow(string.Format("  Shape({0}).Cell(Section:{1} RowName:{2} Name:{3})", shape.Name, cell.Section, cell.RowName, cell.Name));
                    }
                    else
                    {
                        shape.AddNamedRow((short)Visio.VisSectionIndices.visSectionUser, "IsPageName",0 );
                    	shape.Cells["User.isPageName"].ResultIU = 1.0;
                    }

                    UpdatePageNameShape(shape, page.NameU);
                }
                catch (Exception ex)
                {
                    Log.Error(ex, Common.LOG_CATEGORY);
                }
            }

            VNCVisioAddIn.Common.DisplayInDebugWindow(string.Format("Visio_Shape.Add_User_IsPageName() {0}", "End"));
        }

        public static void AddColorSupportToSelection()
        {
            Visio.Application app = Globals.ThisAddIn.Application;

            Visio.Selection selection = app.ActiveWindow.Selection;

            foreach (Visio.Shape shape in selection)
            {
                Add_ColorSupport(shape);
            }
        }

        public static void AddHyperlinkToPage_FromShapeText()
        {
            Visio.Application app = Globals.ThisAddIn.Application;

            Visio.Selection selection = app.ActiveWindow.Selection;

            foreach (Visio.Shape shape in selection)
            {
                string pageName = shape.Text;
                AddHyperLink(shape, pageName);
            }
        }

        public static void AddHyperLink(Visio.Shape shape, string pageName)
        {
            try
            {
                // TODO(crhodes):
                //	Validate pageName matches an existing pageName

                Visio.Hyperlink hlink = shape.Hyperlinks.Add();
                // hlink.Name = "do we need a name?";
                hlink.SubAddress = pageName;
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }

        public static void Add_IDandTextSupport_ToSelection()
        {
            VNCVisioAddIn.Common.DisplayInDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            Visio.Application app = Globals.ThisAddIn.Application;

            Visio.Selection selection = app.ActiveWindow.Selection;

            foreach (Visio.Shape shape in selection)
            {
                Add_IDandTextSupport(shape);
            }
        }

        public static void Add_IDSupport_ToSelection()
        {
            VNCVisioAddIn.Common.DisplayInDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            Visio.Application app = Globals.ThisAddIn.Application;

            Visio.Selection selection = app.ActiveWindow.Selection;

            foreach (Visio.Shape shape in selection)
            {
                Add_IDSupport(shape);
            }
        }

        public static void Add_TextControl_ToSelection()
        {
            VNCVisioAddIn.Common.DisplayInDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            Visio.Application app = Globals.ThisAddIn.Application;

            Visio.Selection selection = app.ActiveWindow.Selection;

            foreach (Visio.Shape shape in selection)
            {
                Add_TextTransformControl(shape);
            }
        }

        public static void DisplayInfo(Visio.Shape shape)
        {
            var isPageName = shape.CellExists["User.IsPageName", 0];    // 0 is Local and Inherited, 1 is Local only 

            //VNCVisioAddIn.Common.DisplayInDebugWindow(string.Format("    Shape({0}).IsPageName({1})", shape.Name, isPageName));

            if (isPageName != 0)
            {
                Visio.Cell cell = shape.Cells["User.IsPageName"];
            }
            else
            {
                
            }

            VNCVisioAddIn.Common.DisplayInDebugWindow(string.Format("   Shape(ID:{0}  Name:{1}  Text:>{2}<)",
                shape.ID, shape.Name, shape.Text));
        }

        public static void ClearConnectionPoints(string tag)
        {
            VNCVisioAddIn.Common.DisplayInDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            Visio.Application app = Globals.ThisAddIn.Application;

            Visio.Page page = app.ActivePage;

            Visio.Selection selection = app.ActiveWindow.Selection;

            foreach (Visio.Shape shape in selection)
            {
                Clear_ConnectionPoints(shape, tag);
            }
        }

        static void Clear_ConnectionPoints(Visio.Shape shape, string tag)
        {
            // TODO(crhodes)
            // Add a remove Connection Points method to clear things out.

            short sectionConnection = (short)Visio.VisSectionIndices.visSectionConnectionPts;

            try
            {
                short exists = shape.SectionExists[sectionConnection, 0];

                if (exists != 0)
                {
                    int rows = shape.RowCount[sectionConnection];

                    Visio.Section connectionPoints = shape.Section[sectionConnection];

                    for (short row = 0; row < rows; row++)
                    {
                        Visio.Cell cnnctX = shape.CellsSRC[sectionConnection, row, (short)Visio.VisCellIndices.visCnnctX];
                        Visio.Cell cnnctY = shape.CellsSRC[sectionConnection, row, (short)Visio.VisCellIndices.visCnnctY];
                        Visio.Cell cnnctDirX = shape.CellsSRC[sectionConnection, row, (short)Visio.VisCellIndices.visCnnctDirX];
                        Visio.Cell cnnctA = shape.CellsSRC[sectionConnection, row, (short)Visio.VisCellIndices.visCnnctA];
                        Visio.Cell cnnctDirY = shape.CellsSRC[sectionConnection, row, (short)Visio.VisCellIndices.visCnnctDirY];
                        Visio.Cell cnnctB = shape.CellsSRC[sectionConnection, row, (short)Visio.VisCellIndices.visCnnctB];
                        Visio.Cell cnnctType = shape.CellsSRC[sectionConnection, row, (short)Visio.VisCellIndices.visCnnctType];
                        Visio.Cell cnnctC = shape.CellsSRC[sectionConnection, row, (short)Visio.VisCellIndices.visCnnctC];
                        Visio.Cell cnnctD = shape.CellsSRC[sectionConnection, row, (short)Visio.VisCellIndices.visCnnctD];

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
                            shape.DeleteSection((short)Visio.VisSectionIndices.visSectionConnectionPts);
                            break;

                        default:
                            MessageBox.Show($"Unknown tag: {tag}");
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }

        internal static void Add_ConnectionPoints(List<VNCVisioAddIn.Domain.ConnectionPointRow> connectionPoints)
        {
            VNCVisioAddIn.Common.DisplayInDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            Visio.Application app = Globals.ThisAddIn.Application;

            Visio.Page page = app.ActivePage;

            Visio.Selection selection = app.ActiveWindow.Selection;

            foreach (Visio.Shape shape in selection)
            {
                Add_ConnectionPoints(shape, connectionPoints);
            }
        }

        static void Add_Connection_Row(Shape shape,
            VNCVisioAddIn.Domain.ConnectionPointRow connectionPoint)
        {
            short sectionConnectionPts = (short)Visio.VisSectionIndices.visSectionConnectionPts;
            short tagConnectionPts = (short)VisRowTags.visTagCnnctPt;
            short tagConnectionNamed = (short)VisRowTags.visTagCnnctNamedABCD;

            //short rowID = shape.AddRow(sectionConnectionPts, (short)VisRowIndices.visRowLast, tagConnectionPts);
            short rowID = shape.AddNamedRow(sectionConnectionPts, connectionPoint.Name, tagConnectionNamed);

            
            //shape.CellsSRC[
            //    (short)Visio.VisSectionIndices.visSectionConnectionPts,
            //    rowID,
            //    (short)Visio.VisCellIndices.visCnnctD].FormulaU = "Top";

            //shape.CellsSRC[
            //    (short)Visio.VisSectionIndices.visSectionConnectionPts,
            //    rowID,
            //    (short)Visio.VisCellIndices.visCnnctX].RowNameU = connectionPoint.Name;

            shape.CellsSRC[
                sectionConnectionPts,
                rowID,
                (short)Visio.VisCellIndices.visCnnctX].FormulaU = connectionPoint.X;

            shape.CellsSRC[
                sectionConnectionPts,
                 rowID,
                (short)Visio.VisCellIndices.visCnnctY].FormulaU = connectionPoint.Y;

            shape.CellsSRC[
                sectionConnectionPts,
                rowID,
                (short)Visio.VisCellIndices.visCnnctDirX].FormulaU = connectionPoint.DirX;

            shape.CellsSRC[
                sectionConnectionPts,
                 rowID,
                (short)Visio.VisCellIndices.visCnnctDirY].FormulaU = connectionPoint.DirY;

            shape.CellsSRC[
                sectionConnectionPts,
                 rowID,
                (short)Visio.VisCellIndices.visCnnctType].FormulaU = connectionPoint.Type;

            shape.CellsSRC[
                sectionConnectionPts,
                 rowID,
                (short)Visio.VisCellIndices.visCnnctD].FormulaU = connectionPoint.D.WrapInDblQuotes();
        }

        static void Add_ConnectionPoints(Visio.Shape shape, List<VNCVisioAddIn.Domain.ConnectionPointRow> connectionPoints)
        {
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
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }

        public static void DisplayShapeInfo(Visio.Shape shape)
        {
            VNCVisioAddIn.Common.DisplayInDebugWindow(string.Format("{0}()  Shape(ID:{1}  Name:{2}  Text:>{3}<)",
                MethodBase.GetCurrentMethod().Name,
                shape.ID, shape.Name, shape.Text));
        }

        public static void GatherInfo()
        {
            VNCVisioAddIn.Common.DisplayInDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            Visio.Application app = Globals.ThisAddIn.Application;

            StringBuilder sb = new StringBuilder();

            Visio.Page page = app.ActivePage;

            Visio.Selection selection = app.ActiveWindow.Selection;

            foreach (Visio.Shape shape in selection)
            {
                DisplayInfo(shape);
            }

            VNCVisioAddIn.Common.DisplayInDebugWindow(string.Format("{0}() {1}",
                MethodBase.GetCurrentMethod().Name, "End"));
        }

        public static void HandleShapeAdded(Visio.Shape shape)
        {
            var isPageName = shape.CellExists["User.IsPageName", 0];    // 0 is Local and Inherited, 1 is Local only 
            //VNCVisioAddIn.Common.DisplayInDebugWindow(string.Format("{0}({1}  isPageName:{2})",
            //    MethodBase.GetCurrentMethod().Name, shape.Name, isPageName));

            //VNCVisioAddIn.Common.DisplayInDebugWindow(string.Format("  UpdatePageNameShape({0}).IsPageName({1})", shape.Name, isPageName));

            if (isPageName != 0)
            {
                Visio.Cell cell = shape.Cells["User.IsPageName"];

                //VNCVisioAddIn.Common.DisplayInDebugWindow(string.Format("    Shape({0}).Cell(Section:{1} RowName:{2} Name:{3} Value:{4})",
                //    shape.Name, cell.Section, cell.RowName, cell.Name, cell.ResultIU));

                if (cell.ResultIU > 0)
                {
                    Visio.Application app = Globals.ThisAddIn.Application;
                    Visio.Page page = app.ActivePage;
                    shape.Text = page.NameU;
                }
            }
        }

        public static void LinkShapeToPage(Visio.Application app, string doc, string page, string shape, string shapeu, String[] args)
        {
            string pageLevel = args[0];
            string separator = "";

            VNCVisioAddIn.Common.DisplayInDebugWindow(string.Format("{0}() PageLevel:{1}",
                MethodBase.GetCurrentMethod().Name,
                pageLevel));

            // Current shape contains text for new page name.

            Visio.Shape activeShape = app.ActivePage.Shapes[shape];
            VNCVisioAddIn.Common.DisplayInDebugWindow(string.Format("  Shape(Name:{0}  Text:{1}", activeShape.Name, activeShape.Text));

            // Update the current shape's hyperlink to point to the page represented by the text

            if (pageLevel.Length > 0)
            {
                separator = "-";
            }

            // shape.Text comes in as OBJ if use fields and Shape Data.   Use shape.Characters instead. 
            string pageName = string.Format("{0}{1}{2}", pageLevel, separator, activeShape.Characters.TextAsString.Replace("\n", " "));
            //string pageName = string.Format("{0}{1}{2}", pageLevel, separator, activeShape.Text.Replace("\n", " "));

            // TODO(crhodes):
            //	Not sure which of these two approaches is doing the magic.

            Visio.Hyperlink newHyperLink = activeShape.AddHyperlink();
            newHyperLink.SubAddress = pageName;

            //activeShape.CellsSRC[(short)Visio.VisSectionIndices.visSectionHyperlink, 0, 2].Formula = pageName;

        }

        public static void MakeLinkableMaster()
        {
            VNCVisioAddIn.Common.DisplayInDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            Visio.Application app = Globals.ThisAddIn.Application;

            Visio.Selection selection = app.ActiveWindow.Selection;

            foreach (Visio.Shape shape in selection)
            {
                MakeLinkableMaster(shape);
            }
        }

        private static void MakeLinkableMaster(Microsoft.Office.Interop.Visio.Shape shape)
        {
            VNCVisioAddIn.Common.DisplayInDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            try
            {
                Validate_Action_SectionExists(shape);

                Delete_Section_Row(shape, Visio.VisSectionIndices.visSectionAction, "Actions", "CreateActivityPage");
                Delete_Section_Row(shape, Visio.VisSectionIndices.visSectionAction, "Actions", "LinkShapeToPage");

                Add_ActionSection_Row(shape,
                    "CreateActivityPage",
                    "=RUNADDONWARGS(\"QueueMarkerEvent\",\"CreatePageForShape,,, Page % 20Base\")",
                    "Create Page for Shape");
                Add_ActionSection_Row(shape,
                    "LinkShapeToPage",
                    "=RUNADDONWARGS(\"QueueMarkerEvent\",\"LinkShapeToPage, \")",
                    "Link Shape to Page");

                Validate_Prop_SectionExists(shape);

                Delete_Section_Row(shape, Visio.VisSectionIndices.visSectionProp, "Prop", "HyperLink");
                Delete_Section_Row(shape, Visio.VisSectionIndices.visSectionProp, "Prop", "ReturnLink");
                Delete_Section_Row(shape, Visio.VisSectionIndices.visSectionProp, "Prop", "PageName");
                Delete_Section_Row(shape, Visio.VisSectionIndices.visSectionProp, "Prop", "ExternalLink");
                Delete_Section_Row(shape, Visio.VisSectionIndices.visSectionProp, "Prop", "HyperLinkPrefix");

                Add_Prop_Row(shape, "PageName", "PageName", (short)Visio.VisCellVals.visPropTypeString, null, "<page name>".WrapInDblQuotes(), prompt: "Enter Page Name", ask: "TRUE");
                Add_Prop_Row(shape, "HyperLink", "HyperLink", (short)Visio.VisCellVals.visPropTypeString, null, "");
                Add_Prop_Row(shape, "ReturnLink", "ReturnLink", (short)Visio.VisCellVals.visPropTypeString, null, "Page Shapes.vssx,PageLink Arrow Left".WrapInDblQuotes());
                Add_Prop_Row(shape, "ExternalLink", "ExternalLink", (short)Visio.VisCellVals.visPropTypeString, null, "");
                Add_Prop_Row(shape, "HyperLinkPrefix", "HyperLinkPrefix", (short)Visio.VisCellVals.visPropTypeString, null, "");

                // For now assume the shape does not have any hyperlinks we care about.

                shape.DeleteSection((short)Visio.VisSectionIndices.visSectionHyperlink);

                Validate_HyperLink_SectionExists(shape);

                Visio.Hyperlink newHyperLink = shape.AddHyperlink();

                // This doesn't work as the value is treated as a string.
                //newHyperLink.SubAddress = "Prop.HyperLink";

                // Need to go at it as a CellSRC

                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionHyperlink,
                    (short)Visio.VisRowIndices.visRow1stHyperlink,
                    (short)Visio.VisCellIndices.visHLinkSubAddress].FormulaU = "GUARD(Prop.HyperLink)";

                // This creates a section that we could update, but the shape.Characters also adds a TextField row
                // we don't need two.
                //Validate_TextField_SectionExists(shape);

                // Not sure how to go about this.  Macro recorder shows this

                Visio.Characters characters = shape.Characters;
                characters.AddCustomFieldU("GUARD(Prop.PageName)", (short)Visio.VisFieldFormats.visFmtNumGenNoUnits);

                // TODO(crhodes)
                // Need to protect the Text so not accidentally overridden.
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }

        public static void MoveToBackgroundLayer()
        {
            Visio.Application app = Globals.ThisAddIn.Application;
            Visio.Page currentPage = app.ActivePage;
            Visio.Layer backgroundLayer = null;
            string layerLock = "1"; // Default is to lock when moved.

            // See if layer exists.  If so, save current lock setting.

            try
            {
                backgroundLayer = currentPage.Layers["Background"];
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
                
            if (backgroundLayer == null)
            {
                backgroundLayer = Visio_Page.AddLayer(currentPage, "Background");
            }
            else
            {
                layerLock = backgroundLayer.CellsC[(short)Visio.VisCellIndices.visLayerLock].FormulaU;
            }

            Visio.Selection selection = app.ActiveWindow.Selection;

            foreach (Visio.Shape shape in selection)
            {
                backgroundLayer.Add(shape, 1);
            }

            backgroundLayer.CellsC[(short)Visio.VisCellIndices.visLayerLock].FormulaU = layerLock;
        }

        public static void Populate_Actions_Section(Visio.Shape shape, string actionName, string action, string menu, string tagName, string buttonFace, string sortKey, string isChecked, string isDisabled, string isReadOnly, string isInvisible, string beginGroup, string flyoutChild)
        {
            Add_ActionSection_Row(shape,
                actionName,
                action,
                menu,
                tagName,
                buttonFace, sortKey, isChecked, isDisabled, isReadOnly, isInvisible, beginGroup, flyoutChild);
        }

        public static void Populate_Hyperlinks_Section(Visio.Shape shape, string rowName, string description, string address, string subAddress, string extraInfo, string frame, string sortKey, string newWindow, string default1, string invisible)
        {
            Add_HyperlinkSection_Row(shape,
                rowName,
                description, 
                address, 
                subAddress, 
                extraInfo, frame, sortKey, newWindow, default1, invisible);
        }

        private static void Set_RowFill_Cell(Visio.Shape shape, Visio.VisCellIndices cellIndex, string value)
        {
            if (value != null)
            {
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionObject,
                    (short)Visio.VisRowIndices.visRowFill,
                    (short)cellIndex].FormulaU = value;
            }
        }

        private static void Set_RowXFormOut_Cell(Visio.Shape shape, Visio.VisCellIndices cellIndex, string value)
        {
            if (value != null)
            {
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionObject,
                    (short)Visio.VisRowIndices.visRowXFormOut,
                    (short)cellIndex].FormulaU = value;
            }
        }

        public static void Set_FillFormat_SectionOld(Microsoft.Office.Interop.Visio.Shape shape, 
            string fillForegnd = null, string fillForegndTrans = null, 
            string fillBkgnd = null, string fillBkgndTrans = null, string fillPattern = null,
            string shdwForegnd = null, string shdwForegndTrans = null, 
            string shdwPattern = null, string shapeShdwOffsetX = null, string shapeShdwOffsetY = null,
            string shapeShdwType = null, string shapeShdwObliqueAngle = null, string shapeShdwScaleFactor = null, 
            string shapeShdwBlur = null, string shapeShdwShow = null)
        {
            VNCVisioAddIn.Common.DisplayInDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            // This Section always exists, so just set values

            // Everything defaults to null and is in the likely order of most often changed.
            // If null, skip setting value.

            try
            {
                Set_RowFill_Cell(shape, Visio.VisCellIndices.visFillForegnd, fillForegnd);
                Set_RowFill_Cell(shape, Visio.VisCellIndices.visFillForegndTrans, fillForegndTrans);

                Set_RowFill_Cell(shape, Visio.VisCellIndices.visFillBkgnd, fillBkgnd);
                Set_RowFill_Cell(shape, Visio.VisCellIndices.visFillBkgndTrans, fillBkgndTrans);
                Set_RowFill_Cell(shape, Visio.VisCellIndices.visFillPattern, fillPattern);

                Set_RowFill_Cell(shape, Visio.VisCellIndices.visFillShdwForegnd, shdwForegnd);
                Set_RowFill_Cell(shape, Visio.VisCellIndices.visFillShdwForegndTrans, shdwForegndTrans);

                Set_RowFill_Cell(shape, Visio.VisCellIndices.visFillShdwPattern, shdwPattern);
                Set_RowFill_Cell(shape, Visio.VisCellIndices.visFillShdwOffsetX, shapeShdwOffsetX);
                Set_RowFill_Cell(shape, Visio.VisCellIndices.visFillShdwOffsetY, shapeShdwOffsetY);

                Set_RowFill_Cell(shape, Visio.VisCellIndices.visFillShdwType, shapeShdwType);
                Set_RowFill_Cell(shape, Visio.VisCellIndices.visFillShdwObliqueAngle, shapeShdwObliqueAngle);
                Set_RowFill_Cell(shape, Visio.VisCellIndices.visFillShdwScaleFactor, shapeShdwScaleFactor);
                Set_RowFill_Cell(shape, Visio.VisCellIndices.visFillShdwBlur, shapeShdwBlur);
                Set_RowFill_Cell(shape, Visio.VisCellIndices.visFillShdwShow, shapeShdwShow);
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }

        public static void Set_ShapeTransform_Section(Visio.Shape shape, 
                string width = null, string height = null, string pinX = null, string pinY = null, 
                string flipX = null, string flipY = null, string locPinX = null, string locPinY = null, 
                string angle = null, string resizeMode = null)
        {
            VNCVisioAddIn.Common.DisplayInDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            // This Section always exists, so just set values

            // Everything defaults to null and is in the likely order of most often changed.
            // If null, skip setting value.

            try
            {
                Set_RowXFormOut_Cell(shape, Visio.VisCellIndices.visXFormWidth, width);
                Set_RowXFormOut_Cell(shape, Visio.VisCellIndices.visXFormHeight, height);
                Set_RowXFormOut_Cell(shape, Visio.VisCellIndices.visXFormPinX, pinX);
                Set_RowXFormOut_Cell(shape, Visio.VisCellIndices.visXFormPinY, pinY);
                Set_RowXFormOut_Cell(shape, Visio.VisCellIndices.visXFormFlipX, flipX);
                Set_RowXFormOut_Cell(shape, Visio.VisCellIndices.visXFormFlipY, flipY);
                Set_RowXFormOut_Cell(shape, Visio.VisCellIndices.visXFormLocPinX, locPinX);
                Set_RowXFormOut_Cell(shape, Visio.VisCellIndices.visXFormLocPinY, locPinY);
                Set_RowXFormOut_Cell(shape, Visio.VisCellIndices.visXFormAngle, angle);
                Set_RowXFormOut_Cell(shape, Visio.VisCellIndices.visXFormResizeMode, resizeMode);
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }

        public static void UpdatePageNameShape(Visio.Shape shape, string pageName)
        {
            var isPageName = shape.CellExistsU["User.IsPageName", 0];    // 0 is Local and Inherited, 1 is Local only 

            VNCVisioAddIn.Common.DisplayInDebugWindow(string.Format("{0}({1}  isPageName:{2})",
                MethodBase.GetCurrentMethod().Name, shape.Name, isPageName));

            //VNCVisioAddIn.Common.DisplayInDebugWindow(string.Format("  UpdatePageNameShape({0}).IsPageName({1})", shape.Name, isPageName));

            if (isPageName != 0)
            {
                Visio.Cell cell = shape.CellsU["User.IsPageName"];

                VNCVisioAddIn.Common.DisplayInDebugWindow(string.Format("    Shape({0}).Cell(Section:{1} RowName:{2} Name:{3} Value:{4})",
                    shape.Name, cell.Section, cell.RowName, cell.Name, cell.ResultIU));

                if (cell.ResultIU > 0)
                {
                    shape.Text = pageName;
                }
            }
        }

        public static void SetMargins(string points)
        {
            VNCVisioAddIn.Common.DisplayInDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            Visio.Application app = Globals.ThisAddIn.Application;

            Visio.Selection selection = app.ActiveWindow.Selection;

            foreach (Visio.Shape shape in selection)
            {
                SetAllMargins(shape, points);
            }
        }

        #endregion

        #region Private Methods

        private static void Add_ColorSupport(Visio.Shape shape)
        {
            // Have to add these in the right order as there are some dependencies

            string value = string.Empty;

            value = "AliceBlue; AntiqueWhite; Aqua; Aquamarine; Azure; Beige; Bisque; Black; BlanchedAlmond; Blue; BlueViolet; Brown; BurlyWood; CadetBlue; Chartreuse; Chocolate; Coral; CornflowerBlue; Cornsilk; Crimson; Cyan; DarkBlue; DarkCyan; DarkGoldenrod; DarkGray; DarkGreen; DarkKhaki; DarkMagenta; DarkOliveGreen; DarkOrange; DarkOrchid; DarkRed; DarkSalmon; DarkSeaGreen; DarkSlateBlue; DarkSlateGray; DarkTurquoise; DarkViolet; DeepPink; DeepSkyBlue; DimGray; DodgerBlue; Firebrick; FloralWhite; ForestGreen; Fuchsia; Gainsboro; GhostWhite; Gold; Goldenrod; Gray; Green; GreenYellow; Honeydew; HotPink; IndianRed; Indigo; Ivory; Khaki; Lavender; LavenderBlush; LawnGreen; LemonChiffon; LightBlue; LightCoral; LightCyan; LightGoldenrodYellow; LightGreen; LightGray; LightPink; LightSalmon; LightSeaGreen; LightSkyBlue; LightSlateGray; LightSteelBlue; LightYellow; Lime; LimeGreen; Linen; Magenta; Maroon; MediumAquamarine; MediumBlue; MediumOrchid; MediumPurple; MediumSeaGreen; MediumSlateBlue; MediumSpringGreen; MediumTurquoise; MediumVioletRed; MidnightBlue; MintCream; MistyRose; Moccasin; NavajoWhite; Navy; OldLace; Olive; OliveDrab; Orange; OrangeRed; Orchid; PaleGoldenrod; PaleGreen; PaleTurquoise; PaleVioletRed; PapayaWhip; PeachPuff; Peru; Pink; Plum; PowderBlue; Purple; Red; RosyBrown; RoyalBlue; SaddleBrown; Salmon; SandyBrown; SeaGreen; SeaShell; Sienna; Silver; SkyBlue; SlateBlue; SlateGray; Snow; SpringGreen; SteelBlue; Tan; Teal; Thistle; Tomato; Turquoise; Violet; Wheat; White; WhiteSmoke; Yellow; YellowGreen";
            value = string.Format("\"{0}\"", value);
            Add_User_Row(shape, "colorNames", value);

            value = "RGB(240, 248, 255); RGB(250, 235, 215); RGB(0, 255, 255); RGB(127, 255, 212); RGB(240, 255, 255); RGB(245, 245, 220); RGB(255, 228, 196); RGB(0, 0, 0); RGB(255, 235, 205); RGB(0, 0, 255); RGB(138, 43, 226); RGB(165, 42, 42); RGB(222, 184, 135); RGB(95, 158, 160); RGB(127, 255, 0); RGB(210, 105, 30); RGB(255, 127, 80); RGB(100, 149, 237); RGB(255, 248, 220); RGB(220, 20, 60); RGB(0, 255, 255); RGB(0, 0, 139); RGB(0, 139, 139); RGB(184, 134, 11); RGB(169, 169, 169); RGB(0, 100, 0); RGB(189, 183, 107); RGB(139, 0, 139); RGB(85, 107, 47); RGB(255, 140, 0); RGB(153, 50, 204); RGB(139, 0, 0); RGB(233, 150, 122); RGB(143, 188, 139); RGB(72, 61, 139); RGB(47, 79, 79); RGB(0, 206, 209); RGB(148, 0, 211); RGB(255, 20, 147); RGB(0, 191, 255); RGB(105, 105, 105); RGB(30, 144, 255); RGB(178, 34, 34); RGB(255, 250, 240); RGB(34, 139, 34); RGB(255, 0, 255); RGB(220, 220, 220); RGB(248, 248, 255); RGB(255, 215, 0); RGB(218, 165, 32); RGB(128, 128, 128); RGB(0, 128, 0); RGB(173, 255, 47); RGB(240, 255, 240); RGB(255, 105, 180); RGB(205, 92, 92); RGB(75, 0, 130); RGB(255, 255, 240); RGB(240, 230, 140); RGB(230, 230, 250); RGB(255, 240, 245); RGB(124, 252, 0); RGB(255, 250, 205); RGB(173, 216, 230); RGB(240, 128, 128); RGB(224, 255, 255); RGB(250, 250, 210); RGB(144, 238, 144); RGB(211, 211, 211); RGB(255, 182, 193); RGB(255, 160, 122); RGB(32, 178, 170); RGB(135, 206, 250); RGB(119, 136, 153); RGB(176, 196, 222); RGB(255, 255, 224); RGB(0, 255, 0); RGB(50, 205, 50); RGB(250, 240, 230); RGB(255, 0, 255); RGB(128, 0, 0); RGB(102, 205, 170); RGB(0, 0, 205); RGB(186, 85, 211); RGB(147, 112, 219); RGB(60, 179, 113); RGB(123, 104, 238); RGB(0, 250, 154); RGB(72, 209, 204); RGB(199, 21, 133); RGB(25, 25, 112); RGB(245, 255, 250); RGB(255, 228, 225); RGB(255, 228, 181); RGB(255, 222, 173); RGB(0, 0, 128); RGB(253, 245, 230); RGB(128, 128, 0); RGB(107, 142, 35); RGB(255, 165, 0); RGB(255, 69, 0); RGB(218, 112, 214); RGB(238, 232, 170); RGB(152, 251, 152); RGB(175, 238, 238); RGB(219, 112, 147); RGB(255, 239, 213); RGB(255, 218, 185); RGB(205, 133, 63); RGB(255, 192, 203); RGB(221, 160, 221); RGB(176, 224, 230); RGB(128, 0, 128); RGB(255, 0, 0); RGB(188, 143, 143); RGB(65, 105, 225); RGB(139, 69, 19); RGB(250, 128, 114); RGB(244, 164, 96); RGB(46, 139, 87); RGB(255, 245, 238); RGB(160, 82, 45); RGB(192, 192, 192); RGB(135, 206, 235); RGB(106, 90, 205); RGB(112, 128, 144); RGB(255, 250, 250); RGB(0, 255, 127); RGB(70, 130, 180); RGB(210, 180, 140); RGB(0, 128, 128); RGB(216, 191, 216); RGB(255, 99, 71); RGB(64, 224, 208); RGB(238, 130, 238); RGB(245, 222, 179); RGB(255, 255, 255); RGB(245, 245, 245); RGB(255, 255, 0); RGB(154, 205, 50)";
            value = string.Format("\"{0}\"", value);
            Add_User_Row(shape, "colorValues", value);

            Add_Prop_Row(shape, "Color", "Color", (short)Visio.VisCellVals.visPropTypeListFix, "User.colorNames", "", "", "");

            value = "INDEX(LOOKUP(Prop.Color,User.colorNames),User.colorValues)";
            Add_User_Row(shape, "Color", value);
        }

        private static void Add_IDandTextSupport(Microsoft.Office.Interop.Visio.Shape shape)
        {
            VNCVisioAddIn.Common.DisplayInDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            Validate_Prop_SectionExists(shape);

            Add_Prop_Row(shape, rowName: "ID", label: "ID", type: (short)Visio.VisCellVals.visPropTypeString, format: null, value: "Xnnn".WrapInDblQuotes());
            Add_Prop_Row(shape, "ShowID", "Show ID", (short)Visio.VisCellVals.visPropTypeBool, null, "TRUE".WrapInDblQuotes());

            Add_Prop_Row(shape, "Text", "Text", (short)Visio.VisCellVals.visPropTypeString, null, "<text>".WrapInDblQuotes());
            Add_Prop_Row(shape, "TextLeft", "Text Left", (short)Visio.VisCellVals.visPropTypeString, null, "<left text>".WrapInDblQuotes());
            Add_Prop_Row(shape, "TextTop", "Text Top", (short)Visio.VisCellVals.visPropTypeString, null, "<top text>".WrapInDblQuotes());
            Add_Prop_Row(shape, "TextRight", "Text Right", (short)Visio.VisCellVals.visPropTypeString, null, "<right text>".WrapInDblQuotes());
            Add_Prop_Row(shape, "TextBottom", "Text Bottom", (short)Visio.VisCellVals.visPropTypeString, null, "<bottom text>".WrapInDblQuotes());

            Add_Prop_Row(shape, "ShowLeftText", "Show Left Text", (short)Visio.VisCellVals.visPropTypeBool, null, "FALSE".WrapInDblQuotes());
            Add_Prop_Row(shape, "ShowTopText", "Show Top Text", (short)Visio.VisCellVals.visPropTypeBool, null, "FALSE".WrapInDblQuotes());
            Add_Prop_Row(shape, "ShowRightText", "Show Right Text", (short)Visio.VisCellVals.visPropTypeBool, null, "FALSE".WrapInDblQuotes());
            Add_Prop_Row(shape, "ShowBottomText", "Show Bottom Text", (short)Visio.VisCellVals.visPropTypeBool, null, "FALSE".WrapInDblQuotes());

            Add_Prop_Row(shape, "SizeText", "Size Text", (short)Visio.VisCellVals.visPropTypeNumber, "0.0".WrapInDblQuotes(), "1.0");
            Add_Prop_Row(shape, "SizeLeftText", "Size Left Text", (short)Visio.VisCellVals.visPropTypeNumber, "0.0".WrapInDblQuotes(), "1.0");
            Add_Prop_Row(shape, "SizeTopText", "Size Top Text", (short)Visio.VisCellVals.visPropTypeNumber, "0.0".WrapInDblQuotes(), "1.0");
            Add_Prop_Row(shape, "SizeRightText", "Size Right Text", (short)Visio.VisCellVals.visPropTypeNumber, "0.0".WrapInDblQuotes(), "1.0");
            Add_Prop_Row(shape, "SizeBottomText", "Size Bottom Text", (short)Visio.VisCellVals.visPropTypeNumber, value: "1.0", format: "0.0".WrapInDblQuotes());
            Add_Prop_Row(shape, "SizeIDText", "Size ID Text", (short)Visio.VisCellVals.visPropTypeNumber, value: "1.0", format: "0.0".WrapInDblQuotes());

            Add_Prop_Row(shape, "GroupName", "Group Name", (short)Visio.VisCellVals.visPropTypeString, null, "<group name>".WrapInDblQuotes());
            Add_Prop_Row(shape, "TabColor", "Tab Color (RGB)", (short)Visio.VisCellVals.visPropTypeString, null, "RGB(128,128,128)");
        }
        private static void Add_IDSupport(Visio.Shape shape)
        {
            VNCVisioAddIn.Common.DisplayInDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            Validate_Prop_SectionExists(shape);

            Add_Prop_Row(shape, "ID", "ID", (short)Visio.VisCellVals.visPropTypeString, null, "000".WrapInDblQuotes());
            Add_Prop_Row(shape, "ShowID", "Show ID", (short)Visio.VisCellVals.visPropTypeBool, null, "TRUE".WrapIn("\""));
            Add_Prop_Row(shape, "Text", "Text", (short)Visio.VisCellVals.visPropTypeString, null, "<text>".WrapInDblQuotes());
            Add_Prop_Row(shape, "SizeText", "Size Text", (short)Visio.VisCellVals.visPropTypeNumber, format: "0.0".WrapInDblQuotes(), value: "1.0");
        }

        // TODO(crhodes):
        // This section should be reviewed and if appropriate lifted out into the ShapeEditor 
        // and associated configuration file

        private static void Add_TextTransformControl(Visio.Shape shape)
        {
            VNCVisioAddIn.Common.DisplayInDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            Populate_Controls_Section(shape,
                "Width*0.5",
                "Height*0.5",
                "Controls.Row_1",
                "Controls.Row_1.Y",
                "0",
                "0",
                "TRUE",
                "Reposition Text");

            Set_TextTransform_Section(shape,
                "GUARD(Width*2)",
                "GUARD(Height*2)",
                "GUARD(Controls.Row_1)",
                "GUARD(Controls.Row_1.Y)",
                "TxtWidth*0.5",
                "TxtHeight*0.5",
                "0 deg"
                );

            //Set_Paragraph_Section()

            VNCVisioAddIn.Domain.TextBlockFormat textBlock = new VNCVisioAddIn.Domain.TextBlockFormat();
            Set_TextBlockFormat_Section(shape, textBlock);
        }

        private static void SetAllMargins(Visio.Shape shape, string points)
        {
            VNCVisioAddIn.Common.DisplayInDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            Validate_TextBlockFormat_SectionExists(shape);

            Set_TextBlockMargins(shape, points, points, points, points);
        }

        private static void SetMargins(Visio.Shape shape, string leftPoints, string topPoints, string rightPoints, string bottomPoints)
        {
            VNCVisioAddIn.Common.DisplayInDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            Validate_TextBlockFormat_SectionExists(shape);

            Set_TextBlockMargins(shape, leftPoints, topPoints, rightPoints, bottomPoints);
        }

        //private static void ZeroMargins(Visio.Shape shape)
        //{
        //    VNCVisioAddIn.Common.DisplayInDebugWindow(string.Format("{0}()",
        //        MethodBase.GetCurrentMethod().Name));

        //    Validate_TextBlockFormat_SectionExists(shape);

        //    Set_TextBlockMargins(shape, "0", "0", "0", "0");
        //}

        #endregion

        #region Utility Routines

        private static void Add_ActionSection_Row(Visio.Shape shape, string rowName, 
            string action, 
            string menu,
            string tagName = "", 
            string buttonFace = "", 
            string sortKey = "",
            string isChecked = "0", 
            string isDisabled = "0", 
            string isReadOnly = "0", 
            string isInvisible = "0", 
            string beginGroup = "0", 
            string flyoutChild = "0")
        {
            //result = shape.AddRow((short)Visio.VisSectionIndices.visSectionAction, (short)Visio.VisRowIndices.visRowLast, (short)Visio.VisRowTags.visTagDefault);
            // TODO(crhodes):
            // Determine what this does if row already exists.
      
            try
            {
                var rowNumber = shape.AddNamedRow((short)Visio.VisSectionIndices.visSectionAction, rowName, (short)Visio.VisRowTags.visTagDefault);

                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionAction,
                    rowNumber,
                    (short)Visio.VisCellIndices.visActionAction].FormulaU = action;
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionAction,
                    rowNumber,
                    (short)Visio.VisCellIndices.visActionMenu].FormulaU = menu.WrapInDblQuotes();
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionAction,
                    rowNumber,
                    (short)Visio.VisCellIndices.visActionTagName].FormulaU = tagName;
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionAction,
                    rowNumber,
                    (short)Visio.VisCellIndices.visActionButtonFace].FormulaU = buttonFace;
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionAction,
                    rowNumber,
                    (short)Visio.VisCellIndices.visActionSortKey].FormulaU = sortKey;
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionAction,
                    rowNumber,
                    (short)Visio.VisCellIndices.visActionChecked].FormulaU = isChecked;
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionAction,
                    rowNumber,
                    (short)Visio.VisCellIndices.visActionDisabled].FormulaU = isDisabled;
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionAction,
                    rowNumber,
                    (short)Visio.VisCellIndices.visActionReadOnly].FormulaU = isReadOnly;
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionAction,
                    rowNumber,
                    (short)Visio.VisCellIndices.visActionInvisible].FormulaU = isInvisible;
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionAction,
                    rowNumber,
                    (short)Visio.VisCellIndices.visActionBeginGroup].FormulaU = beginGroup;
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionAction,
                    rowNumber,
                    (short)Visio.VisCellIndices.visActionFlyoutChild].FormulaU = flyoutChild;
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }

        private static void Add_HyperlinkSection_Row(Visio.Shape shape, 
            string rowName, 
            string description, 
            string address, 
            string subAddress,
            string extraInfo = "", 
            string frame = "", 
            string sortKey = "", 
            string newWindow = "0", 
            string default1 = "0", 
            string invisible = "0")
        {
            //result = shape.AddRow((short)Visio.VisSectionIndices.visSectionAction, (short)Visio.VisRowIndices.visRowLast, (short)Visio.VisRowTags.visTagDefault);
            // TODO(crhodes):
            // Determine what this does if row already exists.

            var rowNumber = shape.AddNamedRow((short)Visio.VisSectionIndices.visSectionHyperlink, rowName, (short)Visio.VisRowTags.visTagDefault);

            try
            {
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionHyperlink,
                    rowNumber,
                    (short)Visio.VisCellIndices.visHLinkDescription].FormulaU = description;
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionHyperlink,
                    rowNumber,
                    (short)Visio.VisCellIndices.visHLinkAddress].FormulaU = address;
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionHyperlink,
                    rowNumber,
                    (short)Visio.VisCellIndices.visHLinkSubAddress].FormulaU = subAddress;  // Wrapping in doubleqoutes would break entering formulas
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionHyperlink,
                    rowNumber,
                    (short)Visio.VisCellIndices.visHLinkExtraInfo].FormulaU = extraInfo;
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionHyperlink,
                    rowNumber,
                    (short)Visio.VisCellIndices.visHLinkExtraInfo].FormulaU = frame;
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionHyperlink,
                    rowNumber,
                    (short)Visio.VisCellIndices.visHLinkSortKey].FormulaU = sortKey;
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionHyperlink,
                    rowNumber,
                    (short)Visio.VisCellIndices.visHLinkNewWin].FormulaU = newWindow;
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionHyperlink,
                    rowNumber,
                    (short)Visio.VisCellIndices.visHLinkDefault].FormulaU = default1;
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionHyperlink,
                    rowNumber,
                    (short)Visio.VisCellIndices.visHLinkInvisible].FormulaU = invisible;

            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }

        }

        internal static void Add_ShapeDataSection_Row(Visio.Shape shape, string rowName, 
            string action, 
            string menu,
            string tagName = "", 
            string buttonFace = "", 
            string sortKey = "",
            string isChecked = "0", 
            string isDisabled = "0", 
            string isReadOnly = "0", 
            string isInvisible = "0", 
            string beginGroup = "0", 
            string flyoutChild = "0")
        {
            var rowNumber = shape.AddNamedRow((short)Visio.VisSectionIndices.visSectionAction, rowName, (short)Visio.VisRowTags.visTagDefault);

            try
            {
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionAction,
                    rowNumber,
                    (short)Visio.VisCellIndices.visActionAction].FormulaU = action;
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionAction,
                    rowNumber,
                    (short)Visio.VisCellIndices.visActionMenu].FormulaU = menu;
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionAction,
                    rowNumber,
                    (short)Visio.VisCellIndices.visActionTagName].FormulaU = tagName;
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionAction,
                    rowNumber,
                    (short)Visio.VisCellIndices.visActionButtonFace].FormulaU = buttonFace;
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionAction,
                    rowNumber,
                    (short)Visio.VisCellIndices.visActionSortKey].FormulaU = sortKey;
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionAction,
                    rowNumber,
                    (short)Visio.VisCellIndices.visActionChecked].FormulaU = isChecked;
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionAction,
                    rowNumber,
                    (short)Visio.VisCellIndices.visActionDisabled].FormulaU = isDisabled;
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionAction,
                    rowNumber,
                    (short)Visio.VisCellIndices.visActionReadOnly].FormulaU = isReadOnly;
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionAction,
                    rowNumber,
                    (short)Visio.VisCellIndices.visActionInvisible].FormulaU = isInvisible;
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionAction,
                    rowNumber,
                    (short)Visio.VisCellIndices.visActionBeginGroup].FormulaU = beginGroup;
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionAction,
                    rowNumber,
                    (short)Visio.VisCellIndices.visActionFlyoutChild].FormulaU = flyoutChild;
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }

        /// <summary>
        /// Add a Prop (ShapeData) section to a ShapeSheet
        /// </summary>
        /// <param name="shape"></param>
        /// <param name="rowName"></param>
        /// <param name="label"></param>
        /// <param name="type"></param>
        /// <param name="format">Values must be placed in Quotes ("") if strings</param>
        /// <param name="value">Values must be placed in Quotes ("") if strings</param>
        /// <param name="prompt"></param>
        /// <param name="sortKey"></param>
        /// <param name="invisible"></param>
        /// <param name="ask"></param>
        /// <param name="langID"></param>
        /// <param name="calendar"></param>
        internal static void Add_Prop_Row(Visio.Shape shape, 
            string rowName, 
            string label, short type, string format, string value, 
            string prompt = null, string sortKey = null,
            string invisible = null, string ask = null, string langID = null, string calendar = null)
        {
            Validate_Prop_SectionExists(shape);

            try
            {
                // Add the Row

                short rowNumber = shape.AddNamedRow(
                    (short)Visio.VisSectionIndices.visSectionProp,
                    rowName,
                    (short)Visio.VisRowTags.visTagDefault);

                // And the important cells: Label, Type, Value

                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionProp,
                    rowNumber,
                    (short)Visio.VisCellIndices.visCustPropsLabel].FormulaU = label.WrapInDblQuotes();

                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionProp,
                    rowNumber,
                    (short)Visio.VisCellIndices.visCustPropsType].FormulaU = type.ToString();

                if (format != null)
                {
                    shape.CellsSRC[
                        (short)Visio.VisSectionIndices.visSectionProp,
                        rowNumber,
                        (short)Visio.VisCellIndices.visCustPropsFormat].FormulaU = format.WrapInDblQuotes();    // Is this ever wrong?
                }

                var v1 = value;
                var v2 = value;

                if (value.Contains("\""))
                {
                    value = value.Replace("\"", "\"\"");
                }

                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionProp,
                    rowNumber,
                    (short)Visio.VisCellIndices.visCustPropsValue].FormulaU = value.WrapInDblQuotes();    // Is this ever wrong?;

                // And any optional cells

                if (! String.IsNullOrEmpty(prompt))
                {
                    shape.CellsSRC[
                       (short)Visio.VisSectionIndices.visSectionProp,
                       rowNumber,
                       (short)Visio.VisCellIndices.visCustPropsPrompt].FormulaU = prompt.WrapInDblQuotes();
                }

                //if (null != prompt)
                //{
                //    shape.CellsSRC[
                //       (short)Visio.VisSectionIndices.visSectionProp,
                //       rowNumber,
                //       (short)Visio.VisCellIndices.visCustPropsPrompt].FormulaU = prompt.WrapInDblQuotes();
                //}

                if (null != sortKey)
                {
                    shape.CellsSRC[
                        (short)Visio.VisSectionIndices.visSectionProp,
                        rowNumber,
                        (short)Visio.VisCellIndices.visCustPropsSortKey].FormulaU = sortKey.WrapInDblQuotes();
                }

                // TODO(crhodes):
                // Add support for remaining optional arguments

            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }

        internal static void Add_User_Row(Visio.Shape shape, 
            string rowName, string value, 
            string prompt="")
        {
            Validate_User_SectionExists(shape);

            try
            {
                short rowNumber = shape.AddNamedRow(
                    (short)Visio.VisSectionIndices.visSectionUser, 
                    rowName, 
                    (short)Visio.VisRowTags.visTagDefault);

                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionUser,
                    (short)(Visio.VisRowIndices.visRowControl + rowNumber),
                    (short)Visio.VisCellIndices.visUserValue].FormulaU = value;

                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionUser,
                    (short)(Visio.VisRowIndices.visRowControl + rowNumber),
                    (short)Visio.VisCellIndices.visUserPrompt].FormulaU = string.Format("\"{0}\"", prompt);
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }

        internal static void Delete_Section_Row(
            Visio.Shape shape,
            Visio.VisSectionIndices sectionIndex,
            string sectionName,
            string rowName)
        {
            try
            {
                Validate_Prop_SectionExists(shape);

                short rowIndex = -1;

                if (shape.CellExistsU[$"{sectionName}.{rowName}", 0] != 0)
                {
                    rowIndex = shape.CellsRowIndex[$"{sectionName}.{rowName}"];
                    shape.DeleteRow((short)sectionIndex, rowIndex);
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }

        internal static void Populate_Controls_Section(Visio.Shape shape, 
            string X, string Y, 
            string XDynamics, string YDynamics, 
            string XBehavior, string YBehavior, 
            string CanGlue, string Tip)
        {
            // There can be more than one Controls Row so need to think through how to handle existing rows.

            Validate_Controls_SectionExists(shape);

            short newRow = shape.AddRow(
                (short)Visio.VisSectionIndices.visSectionControls,
                (short)Visio.VisRowIndices.visRowControl,
                (short)Visio.VisRowTags.visTagDefault);

            try
            {
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionControls,
                    (short)Visio.VisRowIndices.visRowControl + 0,
                    (short)Visio.VisCellIndices.visCtlX].FormulaU = X;
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionControls,
                    (short)Visio.VisRowIndices.visRowControl + 0,
                    (short)Visio.VisCellIndices.visCtlY].FormulaU = Y;
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionControls,
                    (short)Visio.VisRowIndices.visRowControl + 0,
                    (short)Visio.VisCellIndices.visCtlXDyn].FormulaU = XDynamics;
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionControls,
                    (short)Visio.VisRowIndices.visRowControl + 0,
                    (short)Visio.VisCellIndices.visCtlYDyn].FormulaU = YDynamics;
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionControls,
                    (short)Visio.VisRowIndices.visRowControl + 0,
                    (short)Visio.VisCellIndices.visCtlXCon].FormulaU = XBehavior;
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionControls,
                    (short)Visio.VisRowIndices.visRowControl + 0,
                    (short)Visio.VisCellIndices.visCtlYCon].FormulaU = YBehavior;
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionControls,
                    (short)Visio.VisRowIndices.visRowControl + 0,
                    (short)Visio.VisCellIndices.visCtlGlue].FormulaU = CanGlue;
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionControls,
                    (short)Visio.VisRowIndices.visRowControl + 0,
                    (short)Visio.VisCellIndices.visCtlTip].FormulaU = string.Format("\"{0}\"", Tip);
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }

        internal static void Populate_Controls_Section(Visio.Shape shape, string rowName,
            string X, string Y,
            string XDynamics, string YDynamics,
            string XBehavior, string YBehavior,
            string CanGlue, string Tip)
        {
            // There can be more than one Controls Row so need to think through how to handle existing rows.

            Validate_Controls_SectionExists(shape);

            short newRow = shape.AddNamedRow(
                (short)Visio.VisSectionIndices.visSectionControls,
                rowName,
                (short)Visio.VisRowTags.visTagDefault);

            try
            {
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionControls,
                    newRow,
                    (short)Visio.VisCellIndices.visCtlX].FormulaU = X;
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionControls,
                    newRow,
                    (short)Visio.VisCellIndices.visCtlY].FormulaU = Y;
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionControls,
                    newRow,
                    (short)Visio.VisCellIndices.visCtlXDyn].FormulaU = XDynamics;
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionControls,
                    newRow,
                    (short)Visio.VisCellIndices.visCtlYDyn].FormulaU = YDynamics;
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionControls,
                    newRow,
                    (short)Visio.VisCellIndices.visCtlXCon].FormulaU = XBehavior;
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionControls,
                    newRow,
                    (short)Visio.VisCellIndices.visCtlYCon].FormulaU = YBehavior;
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionControls,
                    newRow,
                    (short)Visio.VisCellIndices.visCtlGlue].FormulaU = CanGlue;
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionControls,
                    newRow,
                    (short)Visio.VisCellIndices.visCtlTip].FormulaU = string.Format("\"{0}\"", Tip);
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }

        internal static void Set_TextBlockFormat_Section(Visio.Shape shape,
            VNCVisioAddIn.Domain.TextBlockFormat textBlockFormat = null)
        {
            Validate_TextBlockFormat_SectionExists(shape);

            try
            {
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionObject,
                    (short)Visio.VisRowIndices.visRowText,
                    (short)Visio.VisCellIndices.visTxtBlkLeftMargin].FormulaU = textBlockFormat.LeftMargin;
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionObject,
                    (short)Visio.VisRowIndices.visRowText,
                    (short)Visio.VisCellIndices.visTxtBlkTopMargin].FormulaU = textBlockFormat.TopMargin;
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionObject,
                    (short)Visio.VisRowIndices.visRowText,
                    (short)Visio.VisCellIndices.visTxtBlkRightMargin].FormulaU = textBlockFormat.RightMargin;
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionObject,
                    (short)Visio.VisRowIndices.visRowText,
                    (short)Visio.VisCellIndices.visTxtBlkBottomMargin].FormulaU = textBlockFormat.BottomMargin;
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionObject,
                    (short)Visio.VisRowIndices.visRowText,
                    (short)Visio.VisCellIndices.visTxtBlkDirection].FormulaU = textBlockFormat.TextDirection;
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionObject,
                    (short)Visio.VisRowIndices.visRowText,
                    (short)Visio.VisCellIndices.visTxtBlkVerticalAlign].FormulaU = textBlockFormat.VerticalAlign;
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionObject,
                    (short)Visio.VisRowIndices.visRowText,
                    (short)Visio.VisCellIndices.visTxtBlkBkgnd].FormulaU = textBlockFormat.TextBkgnd;
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionObject,
                    (short)Visio.VisRowIndices.visRowText,
                    (short)Visio.VisCellIndices.visTxtBlkBkgndTrans].FormulaU = textBlockFormat.TextBkgndTrans;
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionObject,
                    (short)Visio.VisRowIndices.visRowText,
                    (short)Visio.VisCellIndices.visTxtBlkDefaultTabStop].FormulaU = textBlockFormat.DefaultTabStop;

            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }

        internal static void Set_TextBlockMargins(Visio.Shape shape, 
            string LeftMargin, 
            string TopMargin, 
            string RightMargin, 
            string BottomMargin)
        {
            // TODO(crhodes):
            // Consider making some of the arguments optional with reasonable defaults

            try
            {
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionObject,
                    (short)Visio.VisRowIndices.visRowText,
                    (short)Visio.VisCellIndices.visTxtBlkLeftMargin].FormulaU = LeftMargin;
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionObject,
                    (short)Visio.VisRowIndices.visRowText,
                    (short)Visio.VisCellIndices.visTxtBlkTopMargin].FormulaU = TopMargin;
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionObject,
                    (short)Visio.VisRowIndices.visRowText,
                    (short)Visio.VisCellIndices.visTxtBlkRightMargin].FormulaU = RightMargin;
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionObject,
                    (short)Visio.VisRowIndices.visRowText,
                    (short)Visio.VisCellIndices.visTxtBlkBottomMargin].FormulaU = BottomMargin;
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }

        internal static void Set_TextTransform_Section(Visio.Shape shape,
            string Width, string Height, 
            string PinX, string PinY, 
            string LocPinX, string LocPinY, 
            string Angle)
        {
            Validate_TextXForm_SectionExists(shape);

            try
            {
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionObject, 
                    (short)Visio.VisRowIndices.visRowTextXForm, 
                    (short)Visio.VisCellIndices.visXFormWidth].FormulaU = Width;
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionObject,
                    (short)Visio.VisRowIndices.visRowTextXForm,
                    (short)Visio.VisCellIndices.visXFormHeight].FormulaU = Height;
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionObject,
                    (short)Visio.VisRowIndices.visRowTextXForm,
                    (short)Visio.VisCellIndices.visXFormPinX].FormulaU = PinX;
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionObject,
                    (short)Visio.VisRowIndices.visRowTextXForm,
                    (short)Visio.VisCellIndices.visXFormPinY].FormulaU = PinY;
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionObject,
                    (short)Visio.VisRowIndices.visRowTextXForm,
                    (short)Visio.VisCellIndices.visXFormLocPinX].FormulaU = LocPinX;
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionObject,
                    (short)Visio.VisRowIndices.visRowTextXForm,
                    (short)Visio.VisCellIndices.visXFormLocPinY].FormulaU = LocPinY;
                shape.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionObject,
                    (short)Visio.VisRowIndices.visRowTextXForm,
                    (short)Visio.VisCellIndices.visXFormAngle].FormulaU = Angle;
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }

        private static void Validate_Action_SectionExists(Visio.Shape shape)
        {
            if (0 == shape.SectionExists[(short)Visio.VisSectionIndices.visSectionAction, 0])
            {
                try
                {
                    var result = shape.AddSection((short)Visio.VisSectionIndices.visSectionAction);
                }
                catch (Exception ex)
                {
                    Log.Error(ex, Common.LOG_CATEGORY);
                }
            }
        }

        private static void Validate_Controls_SectionExists(Visio.Shape shape)
        {
            if (0 == shape.SectionExists[(short)Visio.VisSectionIndices.visSectionControls, 0])
            {
                try
                {
                    var result = shape.AddSection((short)Visio.VisSectionIndices.visSectionControls);
                    //result = shape.AddRow(
                    //    (short)Visio.VisSectionIndices.visSectionControls, 
                    //    (short)Visio.VisRowIndices.visRowControl, 
                    //    (short)Visio.VisRowTags.visTagDefault);
                }
                catch (Exception ex)
                {
                    Log.Error(ex, Common.LOG_CATEGORY);
                }
            }
        }

        private static void Validate_HyperLink_SectionExists(Visio.Shape shape)
        {
            // NB. Shape Data = visSectionProp

            if (0 == shape.SectionExists[(short)Visio.VisSectionIndices.visSectionHyperlink, 0])
            {
                try
                {
                    var result = shape.AddSection((short)Visio.VisSectionIndices.visSectionHyperlink);
                }
                catch (Exception ex)
                {
                    Log.Error(ex, Common.LOG_CATEGORY);
                }
            }
        }

        private static void Validate_Prop_SectionExists(Visio.Shape shape)
        {
            // NB. Shape Data = visSectionProp

            if (0 == shape.SectionExists[(short)Visio.VisSectionIndices.visSectionProp, 0])
            {
                try
                {
                    var result = shape.AddSection((short)Visio.VisSectionIndices.visSectionProp);
                }
                catch (Exception ex)
                {
                    Log.Error(ex, Common.LOG_CATEGORY);
                }
            }
        }

        private static void Validate_TextBlockFormat_SectionExists(Visio.Shape shape)
        {
            // TextBlockFormat exists as a row in the SectionObject!

            if (0 == shape.RowExists[
                (short)Visio.VisSectionIndices.visSectionObject,
                (short)Visio.VisRowIndices.visRowText,
                (short)Visio.VisExistsFlags.visExistsAnywhere])
            {
                try
                {
                    shape.AddRow(
                        (short)Visio.VisSectionIndices.visSectionObject,
                        (short)Visio.VisRowIndices.visRowText,
                        (short)Visio.VisRowTags.visTagDefault);
                }
                catch (Exception ex)
                {
                    Log.Error(ex, Common.LOG_CATEGORY);
                }
            }
        }

        private static void Validate_TextXForm_SectionExists(Visio.Shape shape)
        {
            // TextXForm exists as a row in the SectionObject!

            if (0 == shape.RowExists[
                (short)Visio.VisSectionIndices.visSectionObject,
                (short)Visio.VisRowIndices.visRowTextXForm,
                (short)Visio.VisExistsFlags.visExistsAnywhere])
            {
                try
                {
                    shape.AddRow(
                        (short)Visio.VisSectionIndices.visSectionObject,
                        (short)Visio.VisRowIndices.visRowTextXForm,
                        (short)Visio.VisRowTags.visTagDefault);
                }
                catch (Exception ex)
                {
                    Log.Error(ex, Common.LOG_CATEGORY);
                }
            }
        }

        private static void Validate_TextField_SectionExists(Visio.Shape shape)
        {
            if (0 == shape.RowExists[
                (short)Visio.VisSectionIndices.visSectionTextField,
                (short)Visio.VisRowIndices.visRowText,
                (short)Visio.VisExistsFlags.visExistsAnywhere])
            {
                try
                {
                    shape.AddRow(
                        (short)Visio.VisSectionIndices.visSectionTextField,
                        (short)Visio.VisRowIndices.visRowText,
                        (short)Visio.VisRowTags.visTagDefault);
                }
                catch (Exception ex)
                {
                    Log.Error(ex, Common.LOG_CATEGORY);
                }
            }
        }

        private static void Validate_User_SectionExists(Visio.Shape shape)
        {
            if (0 == shape.SectionExists[(short)Visio.VisSectionIndices.visSectionUser, 0])
            {
                try
                {
                    var result = shape.AddSection((short)Visio.VisSectionIndices.visSectionUser);
                }
                catch (Exception ex)
                {
                    Log.Error(ex, Common.LOG_CATEGORY);
                }
            }
        }

        public static void MoveToBackgroundLayer(Visio.Application app, string doc, string page, string shape, string shapeu)
        {
            

        }

        public static void UpdateTextSections(VNCVisioAddIn.Domain.TextBlockFormat textBlockFormat)
        {
            VNCVisioAddIn.Common.DisplayInDebugWindow(string.Format("{0}()",
                MethodBase.GetCurrentMethod().Name));

            Visio.Application app = Globals.ThisAddIn.Application;

            Visio.Page page = app.ActivePage;

            Visio.Selection selection = app.ActiveWindow.Selection;

            foreach (Visio.Shape shape in selection)
            {
                Set_TextBlockFormat_Section(shape, textBlockFormat);
            }
        }

        #endregion

        #region Get ShapeSheet Section

        #region Get ShapeSheet Shape Section

        #region Get ShapeSheet Shape Section Object Based

        public static VNCVisioAddIn.Domain.OneDEndPoints Get_OneDEndPoints(Shape shape)
        {
            VNCVisioAddIn.Domain.OneDEndPoints row = new VNCVisioAddIn.Domain.OneDEndPoints();

            Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionObject];
            Visio.Row sectionRow = section[(short)Visio.VisRowIndices.visRowXForm1D];

            row.BeginX = sectionRow[VisCellIndices.vis1DBeginX].FormulaU;
            row.BeginY = sectionRow[VisCellIndices.vis1DBeginY].FormulaU;
            row.EndX = sectionRow[VisCellIndices.vis1DEndX].FormulaU;
            row.EndY = sectionRow[VisCellIndices.vis1DEndY].FormulaU;

            return row;
        }

        public static VNCVisioAddIn.Domain.AdditionalEffectProperties Get_AdditionalEffectProperties(Shape shape)
        {
            VNCVisioAddIn.Domain.AdditionalEffectProperties row = new VNCVisioAddIn.Domain.AdditionalEffectProperties();

            Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionObject];
            Visio.Row sectionRow = section[(short)Visio.VisRowIndices.visRowOtherEffectProperties];

            row.ReflectionTrans = sectionRow[VisCellIndices.visReflectionTrans].FormulaU;
            row.ReflectionSize = sectionRow[VisCellIndices.visReflectionSize].FormulaU;
            row.ReflectionDist = sectionRow[VisCellIndices.visReflectionDist].FormulaU;
            row.ReflectionBlur = sectionRow[VisCellIndices.visReflectionBlur].FormulaU;
            row.SketchEnabled = sectionRow[VisCellIndices.visSketchEnabled].FormulaU;
            row.SketchSeed = sectionRow[VisCellIndices.visSketchSeed].FormulaU;
            row.SketchAmount = sectionRow[VisCellIndices.visSketchAmount].FormulaU;
            row.SketchLineWeight = sectionRow[VisCellIndices.visSketchLineWeight].FormulaU;
            row.SketchLineChange = sectionRow[VisCellIndices.visSketchLineChange].FormulaU;
            row.SketchFillChange = sectionRow[VisCellIndices.visSketchFillChange].FormulaU;
            row.GlowColor = sectionRow[VisCellIndices.visGlowColor].FormulaU;
            row.GlowColorTrans = sectionRow[VisCellIndices.visGlowColorTrans].FormulaU;
            row.GlowSize = sectionRow[VisCellIndices.visGlowSize].FormulaU;
            row.SoftEdgesSize = sectionRow[VisCellIndices.visSoftEdgesSize].FormulaU;

            return row;
        }
        
        internal static VNCVisioAddIn.Domain.ChangeShapeBehavior Get_ChangeShapeBehavior(Shape shape)
        {
            VNCVisioAddIn.Domain.ChangeShapeBehavior row = new VNCVisioAddIn.Domain.ChangeShapeBehavior();

            Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionObject];
            Visio.Row sectionRow = section[(short)Visio.VisRowIndices.visRowReplaceBehaviors];

            row.ReplaceLockShapeData = sectionRow[VisCellIndices.visReplaceLockShapeData].FormulaU;
            row.ReplaceLockText = sectionRow[VisCellIndices.visReplaceLockText].FormulaU;
            row.ReplaceLockFormat = sectionRow[VisCellIndices.visReplaceLockFormat].FormulaU;
            row.ReplaceCopyCells = sectionRow[VisCellIndices.visReplaceCopyCells].FormulaU;

            return row;
        }

        public static VNCVisioAddIn.Domain.Events Get_Events(Shape shape)
        {
            VNCVisioAddIn.Domain.Events row = new VNCVisioAddIn.Domain.Events();

            Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionObject];
            Visio.Row sectionRow = section[(short)Visio.VisRowIndices.visRowEvent];

            //row.TheData = sectionRow[VisCellIndices.???].FormulaU;
            row.EventDblClick = sectionRow[VisCellIndices.visEvtCellDblClick].FormulaU;
            row.EventDrop = sectionRow[VisCellIndices.visEvtCellDrop].FormulaU;
            row.TheText = sectionRow[VisCellIndices.visEvtCellTheText].FormulaU;
            row.EventXFMod = sectionRow[VisCellIndices.visEvtCellXFMod].FormulaU;
            row.EventMultiDrop = sectionRow[VisCellIndices.visEvtCellMultiDrop].FormulaU;

            return row;
        }

        public static VNCVisioAddIn.Domain.FillFormat Get_FillFormat(Shape shape)
        {
            VNCVisioAddIn.Domain.FillFormat row = new VNCVisioAddIn.Domain.FillFormat();

            Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionObject];
            Visio.Row sectionRow = section[(short)Visio.VisRowIndices.visRowFill];

            row.FillForegnd = sectionRow[VisCellIndices.visFillForegnd].FormulaU;
            row.ShdwForegnd = sectionRow[VisCellIndices.visFillShdwForegnd].FormulaU;
            row.ShapeShdwType = sectionRow[VisCellIndices.visFillShdwType].FormulaU;
            row.FillForegndTrans = sectionRow[VisCellIndices.visFillForegndTrans].FormulaU;
            row.ShdwForegndTrans = sectionRow[VisCellIndices.visFillShdwForegndTrans].FormulaU;
            row.ShapeShdwObliqueAngle = sectionRow[VisCellIndices.visFillShdwObliqueAngle].FormulaU;
            row.FillBkgnd = sectionRow[VisCellIndices.visFillBkgnd].FormulaU;
            row.ShdwPattern = sectionRow[VisCellIndices.visFillShdwPattern].FormulaU;
            row.ShapeShdwScaleFactor = sectionRow[VisCellIndices.visFillShdwScaleFactor].FormulaU;
            row.FillBkgndTrans = sectionRow[VisCellIndices.visFillBkgndTrans].FormulaU;
            row.ShapeShdwOffsetX = sectionRow[VisCellIndices.visFillShdwOffsetX].FormulaU;
            row.ShapeShdwOffsetY = sectionRow[VisCellIndices.visFillShdwOffsetY].FormulaU;
            row.ShapeShdwBlur = sectionRow[VisCellIndices.visFillShdwShow].FormulaU;
            row.FillPattern = sectionRow[VisCellIndices.visFillShdwPattern].FormulaU;
            row.ShapeShdwShow = sectionRow[VisCellIndices.visFillShdwShow].FormulaU;

            return row;
        }

        public static VNCVisioAddIn.Domain.GlueInfo Get_GlueInfo(Shape shape)
        {
            VNCVisioAddIn.Domain.GlueInfo row = new VNCVisioAddIn.Domain.GlueInfo();

            Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionObject];
            Visio.Row sectionRow = section[(short)Visio.VisRowIndices.visRowMisc];

            row.BegTrigger = sectionRow[VisCellIndices.visBegTrigger].FormulaU;
            row.EndTrigger = sectionRow[VisCellIndices.visEndTrigger].FormulaU;
            row.GlueType = sectionRow[VisCellIndices.visGlueType].FormulaU;
            row.WalkPreference = sectionRow[VisCellIndices.visWalkPref].FormulaU;

            return row;
        }

        public static VNCVisioAddIn.Domain.GradientProperties Get_GradientProperties(Shape shape)
        {
            VNCVisioAddIn.Domain.GradientProperties row = new VNCVisioAddIn.Domain.GradientProperties();

            Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionObject];
            Visio.Row sectionRow = section[(short)Visio.VisRowIndices.visRowGradientProperties];

            row.LineGradientDir = sectionRow[VisCellIndices.visLineGradientDir].FormulaU;
            row.LineGradientAngle = sectionRow[VisCellIndices.visLineGradientAngle].FormulaU;
            row.FillGradientDir = sectionRow[VisCellIndices.visFillGradientDir].FormulaU;
            row.FillGradientAngle = sectionRow[VisCellIndices.visFillGradientAngle].FormulaU;
            row.LineGradientEnabled = sectionRow[VisCellIndices.visLineGradientEnabled].FormulaU;
            row.FillGradientEnabled = sectionRow[VisCellIndices.visFillGradientEnabled].FormulaU;
            row.RotateGradientWithShape = sectionRow[VisCellIndices.visRotateGradientWithShape].FormulaU;
            row.UseGroupGradient = sectionRow[VisCellIndices.visUseGroupGradient].FormulaU;

            return row;
        }

        public static VNCVisioAddIn.Domain.GroupProperties Get_GroupProperties(Shape shape)
        {
            VNCVisioAddIn.Domain.GroupProperties row = new VNCVisioAddIn.Domain.GroupProperties();

            Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionObject];
            Visio.Row sectionRow = section[(short)Visio.VisRowIndices.visRowGroup];

            row.SelectMode = sectionRow[VisCellIndices.visGroupSelectMode].FormulaU;
            row.IsTextEditTarget = sectionRow[VisCellIndices.visGroupIsTextEditTarget].FormulaU;
            row.IsDropTarget = sectionRow[VisCellIndices.visGroupIsDropTarget].FormulaU;
            row.DisplayMode = sectionRow[VisCellIndices.visGroupDisplayMode].FormulaU;
            row.IsSnapTarget = sectionRow[VisCellIndices.visGroupIsSnapTarget].FormulaU;
            row.DontMoveChildren = sectionRow[VisCellIndices.visGroupDontMoveChildren].FormulaU;

            return row;
        }

        public static VNCVisioAddIn.Domain.ImageProperties Get_ImageProperties(Shape shape)
        {
            VNCVisioAddIn.Domain.ImageProperties row = new VNCVisioAddIn.Domain.ImageProperties();

            Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionObject];
            Visio.Row sectionRow = section[(short)Visio.VisRowIndices.visRowImage];

            row.Contrast = sectionRow[VisCellIndices.visImageContrast].FormulaU;
            row.Gamma = sectionRow[VisCellIndices.visImageGamma].FormulaU;
            row.Sharpen = sectionRow[VisCellIndices.visImageSharpen].FormulaU;
            row.Brightness = sectionRow[VisCellIndices.visImageBrightness].FormulaU;
            row.Blur = sectionRow[VisCellIndices.visImageBlur].FormulaU;
            row.Denoise = sectionRow[VisCellIndices.visImageDenoise].FormulaU;
            row.Transparency = sectionRow[VisCellIndices.visImageTransparency].FormulaU;

            return row;
        }

        internal static VNCVisioAddIn.Domain.LayerMembership Get_LayerMembership(Shape shape)
        {
            VNCVisioAddIn.Domain.LayerMembership row = new VNCVisioAddIn.Domain.LayerMembership();

            Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionObject];
            Visio.Row sectionRow = section[(short)Visio.VisRowIndices.visRowLayerMem];

              row.Name = sectionRow[VisCellIndices.visLayerMember].FormulaU;

            return row;
        }

        public static VNCVisioAddIn.Domain.LineFormat Get_LineFormat(Shape shape)
        {
            VNCVisioAddIn.Domain.LineFormat row = new VNCVisioAddIn.Domain.LineFormat();

            Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionObject];
            Visio.Row sectionRow = section[(short)Visio.VisRowIndices.visRowLine];

            row.LinePattern = sectionRow[VisCellIndices.visLinePattern].FormulaU;
            row.LineWeight = sectionRow[VisCellIndices.visLineWeight].FormulaU;
            row.LineColor = sectionRow[VisCellIndices.visLineColor].FormulaU;
            row.LineCap = sectionRow[VisCellIndices.visLineEndCap].FormulaU;
            row.BeginArrow = sectionRow[VisCellIndices.visLineBeginArrow].FormulaU;
            row.EndArrow = sectionRow[VisCellIndices.visLineEndArrow].FormulaU;
            row.LineColorTrans = sectionRow[VisCellIndices.visLineColorTrans].FormulaU;
            row.CompoundType = sectionRow[VisCellIndices.visCompoundType].FormulaU;
            row.BeginArrowSize = sectionRow[VisCellIndices.visLineBeginArrowSize].FormulaU;
            row.EndArrowSize = sectionRow[VisCellIndices.visLineEndArrowSize].FormulaU;
            row.Rounding = sectionRow[VisCellIndices.visLineRounding].FormulaU;

            return row;
        }

        internal static VNCVisioAddIn.Domain.Miscellaneous Get_Miscellaneous(Shape shape)
        {
            VNCVisioAddIn.Domain.Miscellaneous row = new VNCVisioAddIn.Domain.Miscellaneous();

            Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionObject];
            Visio.Row sectionRow = section[(short)Visio.VisRowIndices.visRowMisc];

            row.NoObjHandles = sectionRow[VisCellIndices.visNoObjHandles].FormulaU;
            row.NoCtlHandle = sectionRow[VisCellIndices.visNoCtlHandles].FormulaU;
            row.NoAlignBox = sectionRow[VisCellIndices.visNoAlignBox].FormulaU;
            row.NonPrinting = sectionRow[VisCellIndices.visNonPrinting].FormulaU;
            row.LangID = sectionRow[VisCellIndices.visObjLangID].FormulaU;
            row.HideText = sectionRow[VisCellIndices.visHideText].FormulaU;
            row.UpdateAlignBox = sectionRow[VisCellIndices.visUpdateAlignBox].FormulaU;
            row.DynFeedback = sectionRow[VisCellIndices.visDynFeedback].FormulaU;
            row.NoLiveDynamics = sectionRow[VisCellIndices.visNoLiveDynamics].FormulaU;
            row.Calendar = sectionRow[VisCellIndices.visObjCalendar].FormulaU;
            row.ObjType = sectionRow[VisCellIndices.visLOFlags].FormulaU;
            row.IsDropSource = sectionRow[VisCellIndices.visDropSource].FormulaU;
            row.Comment = sectionRow[VisCellIndices.visComment].FormulaU;
            row.DropOnPageScale = sectionRow[VisCellIndices.visObjDropOnPageScale].FormulaU;
            row.LocalizeMerge = sectionRow[VisCellIndices.visObjLocalizeMerge].FormulaU;
            row.NoProofing = sectionRow[VisCellIndices.visObjNoProofing].FormulaU;

            return row;
        }

        public static VNCVisioAddIn.Domain.Protection Get_Protection(Shape shape)
        {
            VNCVisioAddIn.Domain.Protection row = new VNCVisioAddIn.Domain.Protection();

            Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionObject];
            Visio.Row sectionRow = section[(short)Visio.VisRowIndices.visRowLock];

            row.LockWidth = sectionRow[VisCellIndices.visLockWidth].FormulaU;
            row.LockHeight = sectionRow[VisCellIndices.visLockHeight].FormulaU;
            row.LockAspect = sectionRow[VisCellIndices.visLockAspect].FormulaU;
            row.LockMoveX = sectionRow[VisCellIndices.visLockMoveX].FormulaU;
            row.LockMoveY = sectionRow[VisCellIndices.visLockMoveY].FormulaU;
            row.LockRotate = sectionRow[VisCellIndices.visLockRotate].FormulaU;
            row.LockBegin = sectionRow[VisCellIndices.visLockBegin].FormulaU;
            row.LockReplace = sectionRow[VisCellIndices.visLockReplace].FormulaU;
            row.LockEnd = sectionRow[VisCellIndices.visLockEnd].FormulaU;
            row.LockDelete = sectionRow[VisCellIndices.visLockDelete].FormulaU;
            row.LockSelect = sectionRow[VisCellIndices.visLockSelect].FormulaU;
            row.LockFormat = sectionRow[VisCellIndices.visLockFormat].FormulaU;
            row.LockCustProp = sectionRow[VisCellIndices.visLockCustProp].FormulaU;
            row.LockTextEdit = sectionRow[VisCellIndices.visLockTextEdit].FormulaU;
            row.LockVtxEdit = sectionRow[VisCellIndices.visLockVtxEdit].FormulaU;
            row.LockThemeIndex = sectionRow[VisCellIndices.visLockThemeIndex].FormulaU;
            row.LockCrop = sectionRow[VisCellIndices.visLockCrop].FormulaU;
            row.LockGroup = sectionRow[VisCellIndices.visLockGroup].FormulaU;
            row.LockCalcWH = sectionRow[VisCellIndices.visLockCalcWH].FormulaU;
            row.LockFromGroupFormat = sectionRow[VisCellIndices.visLockFromGroupFormat].FormulaU;
            row.LockThemeColors = sectionRow[VisCellIndices.visLockThemeColors].FormulaU;
            row.LockThemeEffects = sectionRow[VisCellIndices.visLockThemeEffects].FormulaU;
            row.LockThemeConnectors = sectionRow[VisCellIndices.visLockThemeConnectors].FormulaU;
            row.LockThemeFonts = sectionRow[VisCellIndices.visLockThemeFonts].FormulaU;
            row.LockVariation = sectionRow[VisCellIndices.visLockVariation].FormulaU;

            return row;
        }

        internal static VNCVisioAddIn.Domain.QuickStyle Get_QuickStyle(Shape shape)
        {
            VNCVisioAddIn.Domain.QuickStyle row = new VNCVisioAddIn.Domain.QuickStyle();

            Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionObject];
            Visio.Row sectionRow = section[(short)Visio.VisRowIndices.visRowQuickStyleProperties];

            row.QuickStyleLineMatrix = sectionRow[VisCellIndices.visQuickStyleLineMatrix].FormulaU;
            row.QuickStyleLineColor = sectionRow[VisCellIndices.visQuickStyleLineColor].FormulaU;
            row.QuickStyleFontColor = sectionRow[VisCellIndices.visQuickStyleFontColor].FormulaU;
            row.QuickStyleVariation = sectionRow[VisCellIndices.visQuickStyleVariation].FormulaU;
            row.QuickStyleFillMatrix = sectionRow[VisCellIndices.visQuickStyleFillMatrix].FormulaU;
            row.QuickStyleFontMatrix = sectionRow[VisCellIndices.visQuickStyleFontMatrix].FormulaU;
            row.QuickStyleEffectsMatrix = sectionRow[VisCellIndices.visQuickStyleEffectsMatrix].FormulaU;
            row.QuickStyleShadowColor = sectionRow[VisCellIndices.visQuickStyleShadowColor].FormulaU;
            row.QuickStyleType = sectionRow[VisCellIndices.visQuickStyleType].FormulaU;

            return row;
        }

        public static VNCVisioAddIn.Domain.ShapeLayout Get_ShapeLayout(Shape shape)
        {
            VNCVisioAddIn.Domain.ShapeLayout row = new VNCVisioAddIn.Domain.ShapeLayout();

            Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionObject];
            Visio.Row sectionRow = section[(short)Visio.VisRowIndices.visRowShapeLayout];

            row.ShapePermeableX = sectionRow[VisCellIndices.visSLOPermX].FormulaU;
            row.ShapePermeableY = sectionRow[VisCellIndices.visSLOPermY].FormulaU;
            row.ShapeFixedCode = sectionRow[VisCellIndices.visSLOFixedCode].FormulaU;
            row.ConLineJumpDirX = sectionRow[VisCellIndices.visSLOJumpDirX].FormulaU;
            row.ConLineJumpDirY = sectionRow[VisCellIndices.visSLOJumpDirY].FormulaU;
            row.ConLineJumpCode = sectionRow[VisCellIndices.visSLOJumpCode].FormulaU;
            row.ShapePlaceFlip = sectionRow[VisCellIndices.visSLOPlaceFlip].FormulaU;
            row.ShapePlaceStyle = sectionRow[VisCellIndices.visSLOPlaceStyle].FormulaU;
            row.ShapePlowCode = sectionRow[VisCellIndices.visSLOPlowCode].FormulaU;
            row.ConLineJumpStyle = sectionRow[VisCellIndices.visSLOJumpStyle].FormulaU;
            row.ConLineRouteExt = sectionRow[VisCellIndices.visSLOLineRouteExt].FormulaU;
            row.DisplayLevel = sectionRow[VisCellIndices.visSLODisplayLevel].FormulaU;
            row.ShapeRouteStyle = sectionRow[VisCellIndices.visSLORouteStyle].FormulaU;
            row.ConFixedCode = sectionRow[VisCellIndices.visSLOConFixedCode].FormulaU;
            row.ShapeSplit = sectionRow[VisCellIndices.visSLOSplit].FormulaU;
            row.ShapeSplittable = sectionRow[VisCellIndices.visSLOSplittable].FormulaU;
            row.Relationships = sectionRow[VisCellIndices.visSLORelationships].FormulaU;

            return row;
        }

        public static VNCVisioAddIn.Domain.TextTransform Get_TextTransform(Shape shape)
        {
            VNCVisioAddIn.Domain.TextTransform row = new VNCVisioAddIn.Domain.TextTransform();

            Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionObject];
            Visio.Row sectionRow = section[(short)Visio.VisRowIndices.visRowTextXForm];

            row.TxtWidth = sectionRow[VisCellIndices.visXFormWidth].FormulaU;
            row.TxtHeight = sectionRow[VisCellIndices.visXFormHeight].FormulaU;
            row.TxtAngle = sectionRow[VisCellIndices.visXFormAngle].FormulaU;
            row.TxtPinX = sectionRow[VisCellIndices.visXFormPinX].FormulaU;
            row.TxtPinY = sectionRow[VisCellIndices.visXFormPinY].FormulaU;
            row.TxtLocPinX = sectionRow[VisCellIndices.visXFormLocPinX].FormulaU;
            row.TxtLocPinY = sectionRow[VisCellIndices.visXFormLocPinY].FormulaU;

            return row;
        }

        internal static VNCVisioAddIn.Domain.ThemeProperties Get_ThemeProperties(Shape shape)
        {
            VNCVisioAddIn.Domain.ThemeProperties row = new VNCVisioAddIn.Domain.ThemeProperties();

            Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionObject];
            Visio.Row sectionRow = section[(short)Visio.VisRowIndices.visRowThemeProperties];

            row.ConnectorSchemeIndex = sectionRow[VisCellIndices.visConnectorSchemeIndex].FormulaU;
            row.EffectSchemeIndex = sectionRow[VisCellIndices.visEffectSchemeIndex].FormulaU;
            row.ColorSchemeIndex = sectionRow[VisCellIndices.visColorSchemeIndex].FormulaU;
            row.FontSchemeIndex = sectionRow[VisCellIndices.visFontSchemeIndex].FormulaU;
            row.ThemeIndex = sectionRow[VisCellIndices.visThemeIndex].FormulaU;
            row.VariationColorIndex = sectionRow[VisCellIndices.visVariationColorIndex].FormulaU;
            row.VariationStyleIndex = sectionRow[VisCellIndices.visVariationStyleIndex].FormulaU;
            row.EmbellishmentIndex = sectionRow[VisCellIndices.visEmbellishmentIndex].FormulaU;

            return row;
        }

        internal static VNCVisioAddIn.Domain.ThreeDRotationProperties Get_ThreeDRotationProperties(Shape shape)
        {
            VNCVisioAddIn.Domain.ThreeDRotationProperties row = new VNCVisioAddIn.Domain.ThreeDRotationProperties();

            Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionObject];
            Visio.Row sectionRow = section[(short)Visio.VisRowIndices.visRow3DRotationProperties];

            row.RotationXAngle = sectionRow[VisCellIndices.visRotationXAngle].FormulaU;
            row.RotationYAngle = sectionRow[VisCellIndices.visRotationYAngle].FormulaU;
            row.RotationZAngle = sectionRow[VisCellIndices.visRotationZAngle].FormulaU;
            row.RotationType = sectionRow[VisCellIndices.visRotationType].FormulaU;
            row.Perspective = sectionRow[VisCellIndices.visPerspective].FormulaU;
            row.DistanceFromGround = sectionRow[VisCellIndices.visDistanceFromGround].FormulaU;
            row.KeepTextFlat = sectionRow[VisCellIndices.visKeepTextFlat].FormulaU;

            return row;
        }

        #endregion

        #region Get ShapeSheet Shape Section Row Based

        public static VNCVisioAddIn.Domain.ActionRow Get_ActionRow(Shape shape)
        {
            VNCVisioAddIn.Domain.ActionRow row = new VNCVisioAddIn.Domain.ActionRow();

            Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionAction];
            Visio.Row sectionRow = section[0];

            // TODO(crhodes)
            // Handle multiple rows

            row.Action = sectionRow[VisCellIndices.visActionAction].FormulaU;

            return row;
        }

        public static VNCVisioAddIn.Domain.ActionTagRow Get_ActionTagRow(Shape shape)
        {
            VNCVisioAddIn.Domain.ActionTagRow row = new VNCVisioAddIn.Domain.ActionTagRow();

            // TODO(crhodes)
            // Can't find Section Index

            //Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.];
            //Visio.Row sectionRow = section[0];

            //// TODO(crhodes)
            //// Handle multiple rows

            //row.Action = sectionRow[VisCellIndices.visActionAction].FormulaU;

            return row;
        }

        public static VNCVisioAddIn.Domain.CharacterRow Get_CharacterRow(Shape shape)
        {
            VNCVisioAddIn.Domain.CharacterRow row = new VNCVisioAddIn.Domain.CharacterRow();

            Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionCharacter];
            Visio.Row sectionRow = section[0];

            // TODO(crhodes)
            // Handle multiple rows

            row.AsianFont = sectionRow[VisCellIndices.visCharacterAsianFont].FormulaU;

            return row;
        }

        public static VNCVisioAddIn.Domain.ConnectionPointRow Get_ConnectionPointRow(Shape shape)
        {
            VNCVisioAddIn.Domain.ConnectionPointRow row = new VNCVisioAddIn.Domain.ConnectionPointRow();

            Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionConnectionPts];
            Visio.Row sectionRow = section[0];

            // TODO(crhodes)
            // Handle multiple rows
            row.Name = sectionRow.Name;

            row.X = sectionRow[VisCellIndices.visCnnctX].FormulaU;

            return row;
        }

        public static VNCVisioAddIn.Domain.ControlsRow Get_ControlsRow(Shape shape)
        {
            VNCVisioAddIn.Domain.ControlsRow controlRow = new VNCVisioAddIn.Domain.ControlsRow();

            Visio.Section controlRowsSection = shape.Section[(short)Visio.VisSectionIndices.visSectionControls];
            Visio.Row firstControlRow = controlRowsSection[0];

            // TODO(crhodes)
            // Handle multiple ControlRows

            controlRow.Name = firstControlRow.Name;

            controlRow.X = firstControlRow[VisCellIndices.visCtlX].FormulaU;
            controlRow.Y = firstControlRow[VisCellIndices.visCtlY].FormulaU;
            controlRow.XDynamics = firstControlRow[VisCellIndices.visCtlXDyn].FormulaU;
            controlRow.YDynamics = firstControlRow[VisCellIndices.visCtlYDyn].FormulaU;
            controlRow.XBehavior = firstControlRow[VisCellIndices.visCtlXCon].FormulaU;
            controlRow.YBehavior = firstControlRow[VisCellIndices.visCtlYCon].FormulaU;
            controlRow.CanGlue = firstControlRow[VisCellIndices.visCtlGlue].FormulaU;
            controlRow.Tip = firstControlRow[VisCellIndices.visCtlTip].FormulaU;

            return controlRow;
        }

        internal static VNCVisioAddIn.Domain.FillGradientStopRow Get_FillGradientStopRow(Shape shape)
        {
            VNCVisioAddIn.Domain.FillGradientStopRow row = new VNCVisioAddIn.Domain.FillGradientStopRow();

            // Shape Transform Section is part of object
            Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionObject];
            Visio.Row sectionRow = section[(short)Visio.VisRowIndices.visRowGradientStop];

            // TODO(crhodes)
            // Handle multiple rows

            //row.NoObjHandles = sectionRow[VisCellIndices.visNoObjHandles].FormulaU;

            return row;
        }

        //public static HyperlinkRow Get_HyperlinkRow(Shape shape)
        //{
        //    HyperlinkRow row = new HyperlinkRow();

        //    Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionHyperlink];
        //    Visio.Row sectionRow = section[0];

        //    // TODO(crhodes)
        //    // Handle multiple rows

        //    row.Name = sectionRow.Name;

        //    row.Address = sectionRow[VisCellIndices.visHLinkAddress].FormulaU;

        //    return row;
        //}

        internal static VNCVisioAddIn.Domain.LayerRow Get_LayerRow(Shape shape)
        {
            VNCVisioAddIn.Domain.LayerRow row = new VNCVisioAddIn.Domain.LayerRow();

            Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionObject];
            Visio.Row sectionRow = section[(short)Visio.VisRowIndices.visRowLayer];

            // TODO(crhodes)
            // Handle multiple rows

            row.Name = sectionRow.Name;
            row.Visible = sectionRow[VisCellIndices.visLayerVisible].FormulaU;
            row.Print = sectionRow[VisCellIndices.visLayerPrint].FormulaU;
            row.Active = sectionRow[VisCellIndices.visLayerActive].FormulaU;
            row.Lock = sectionRow[VisCellIndices.visLayerLock].FormulaU;
            row.Snap = sectionRow[VisCellIndices.visLayerLock].FormulaU;
            row.Glue = sectionRow[VisCellIndices.visLayerGlue].FormulaU;
            row.Color = sectionRow[VisCellIndices.visLayerColor].FormulaU;
            row.Transparency = sectionRow[VisCellIndices.visLayerColorTrans].FormulaU;

            return row;
        }

        internal static ObservableCollection<VNCVisioAddIn.Domain.LayerRow> Get_LayerRows(Shape shape)
        {
            var rows = new ObservableCollection<VNCVisioAddIn.Domain.LayerRow>();

            Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionLayer];

            var rowCount = section.Count;

            for (short i = 0; i < rowCount; i++)
            {
                VNCVisioAddIn.Domain.LayerRow layerRow = new VNCVisioAddIn.Domain.LayerRow();

                var row = section[i];

                layerRow.Name = row[(short)VisCellIndices.visLayerName].FormulaU;
                layerRow.Visible = row[(short)VisCellIndices.visLayerVisible].FormulaU;
                layerRow.Print = row[(short)VisCellIndices.visLayerPrint].FormulaU;
                layerRow.Active = row[(short)VisCellIndices.visLayerActive].FormulaU;
                layerRow.Lock = row[(short)VisCellIndices.visLayerLock].FormulaU;
                layerRow.Snap = row[(short)VisCellIndices.visLayerSnap].FormulaU;
                layerRow.Glue = row[(short)VisCellIndices.visLayerGlue].FormulaU;
                layerRow.Color = row[(short)VisCellIndices.visLayerColor].FormulaU;
                layerRow.Transparency = row[(short)VisCellIndices.visLayerColorTrans].FormulaU;

                // NOTE(crhodes)
                // There are a few more VisCellIndices.  See what they do
                //VisCellIndices.visLayerMember
                //VisCellIndices.visLayerStatus

                rows.Add(layerRow);
            }

            return rows;
        }

        public static VNCVisioAddIn.Domain.LineGradientStopRow Get_LineGradientStopRow(Shape shape)
        {
            VNCVisioAddIn.Domain.LineGradientStopRow row = new VNCVisioAddIn.Domain.LineGradientStopRow();

            Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionLineGradientStops];
            Visio.Row sectionRow = section[0];

            //// TODO(crhodes)
            //// Handle multiple rows

            //row. = sectionRow[VisCellIndices.visLineGradientEnabled].FormulaU;

            return row;
        }

        internal static ObservableCollection<VNCVisioAddIn.Domain.ActionRow> Get_ActionsRows(Shape shape)
        {
            var rows = new ObservableCollection<VNCVisioAddIn.Domain.ActionRow>();

            Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionAction];

            var rowCount = section.Count;

            for (short i = 0; i < rowCount; i++)
            {
                VNCVisioAddIn.Domain.ActionRow actionRow = new VNCVisioAddIn.Domain.ActionRow();

                var row = section[i];

                actionRow.Name = row.NameU;

                actionRow.Action = row[(short)VisCellIndices.visActionAction].FormulaU;
                actionRow.Menu = row[(short)VisCellIndices.visActionMenu].FormulaU;
                actionRow.TagName = row[(short)VisCellIndices.visActionTagName].FormulaU;
                actionRow.ButtonFace = row[(short)VisCellIndices.visActionButtonFace].FormulaU;
                actionRow.SortKey = row[(short)VisCellIndices.visActionSortKey].FormulaU;
                actionRow.Checked = row[(short)VisCellIndices.visActionChecked].FormulaU;
                actionRow.Disabled = row[(short)VisCellIndices.visActionDisabled].FormulaU;
                actionRow.ReadOnly = row[(short)VisCellIndices.visActionReadOnly].FormulaU;
                actionRow.Invisible = row[(short)VisCellIndices.visActionInvisible].FormulaU;
                actionRow.BeginGroup = row[(short)VisCellIndices.visActionBeginGroup].FormulaU;
                actionRow.FlyoutChild = row[(short)VisCellIndices.visActionFlyoutChild].FormulaU;

                rows.Add(actionRow);
            }

            return rows;
        }

        internal static ObservableCollection<VNCVisioAddIn.Domain.ActionTagRow> Get_ActionTagRows(Shape shape)
        {
            var rows = new ObservableCollection<VNCVisioAddIn.Domain.ActionTagRow>();

            Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionSmartTag];

            var rowCount = section.Count;

            for (short i = 0; i < rowCount; i++)
            {
                VNCVisioAddIn.Domain.ActionTagRow actionTagRow = new VNCVisioAddIn.Domain.ActionTagRow();

                var row = section[i];

                actionTagRow.Name = row.NameU;

                actionTagRow.X = row[(short)VisCellIndices.visSmartTagX].FormulaU;
                actionTagRow.Y = row[(short)VisCellIndices.visSmartTagY].FormulaU;
                actionTagRow.TagName = row[(short)VisCellIndices.visSmartTagName].FormulaU;
                actionTagRow.XJustify = row[(short)VisCellIndices.visSmartTagXJustify].FormulaU;
                actionTagRow.YJustify = row[(short)VisCellIndices.visSmartTagYJustify].FormulaU;
                actionTagRow.DisplayMode = row[(short)VisCellIndices.visSmartTagDisplayMode].FormulaU;
                actionTagRow.ButtonFace = row[(short)VisCellIndices.visSmartTagButtonFace].FormulaU;
                actionTagRow.Description = row[(short)VisCellIndices.visSmartTagDescription].FormulaU;
                actionTagRow.Disabled = row[(short)VisCellIndices.visSmartTagDisabled].FormulaU;

                rows.Add(actionTagRow);
            }

            return rows;
        }

        internal static ObservableCollection<VNCVisioAddIn.Domain.ConnectionPointRow> Get_ConnectionPointRows(Shape shape)
        {
            var rows = new ObservableCollection<VNCVisioAddIn.Domain.ConnectionPointRow>();

            Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionConnectionPts];

            var rowCount = section.Count;

            for (short i = 0; i < rowCount; i++)
            {
                VNCVisioAddIn.Domain.ConnectionPointRow connectionPointRow = new VNCVisioAddIn.Domain.ConnectionPointRow();

                var row = section[i];

                connectionPointRow.Name = row.NameU;

                connectionPointRow.X = row[(short)VisCellIndices.visCnnctX].FormulaU;
                connectionPointRow.Y = row[(short)VisCellIndices.visCnnctY].FormulaU;
                connectionPointRow.DirX = row[(short)VisCellIndices.visCnnctDirX].FormulaU;
                connectionPointRow.A = row[(short)VisCellIndices.visCnnctA].FormulaU;
                connectionPointRow.DirY = row[(short)VisCellIndices.visCnnctDirY].FormulaU;
                connectionPointRow.B = row[(short)VisCellIndices.visCnnctB].FormulaU;
                connectionPointRow.Type = row[(short)VisCellIndices.visCnnctType].FormulaU;
                connectionPointRow.C = row[(short)VisCellIndices.visCnnctC].FormulaU;
                connectionPointRow.D = row[(short)VisCellIndices.visCnnctD].FormulaU;

                rows.Add(connectionPointRow);
            }

            return rows;
        }

        internal static ObservableCollection<VNCVisioAddIn.Domain.ControlsRow> Get_ControlsRows(Shape shape)
        {
            var rows = new ObservableCollection<VNCVisioAddIn.Domain.ControlsRow>();

            Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionControls];

            var rowCount = section.Count;

            for (short i = 0; i < rowCount; i++)
            {
                VNCVisioAddIn.Domain.ControlsRow controlsRow = new VNCVisioAddIn.Domain.ControlsRow();

                var row = section[i];

                controlsRow.Name = row.NameU;

                controlsRow.X = row[VisCellIndices.visCtlX].FormulaU;
                controlsRow.Y = row[VisCellIndices.visCtlY].FormulaU;
                controlsRow.XDynamics = row[VisCellIndices.visCtlXDyn].FormulaU;
                controlsRow.YDynamics = row[VisCellIndices.visCtlYDyn].FormulaU;
                controlsRow.XBehavior = row[VisCellIndices.visCtlXCon].FormulaU;
                controlsRow.YBehavior = row[VisCellIndices.visCtlYCon].FormulaU;
                controlsRow.CanGlue = row[VisCellIndices.visCtlGlue].FormulaU;
                controlsRow.Tip = row[VisCellIndices.visCtlTip].FormulaU;

                rows.Add(controlsRow);
            }

            return rows;
        }

        internal static ObservableCollection<VNCVisioAddIn.Domain.HyperlinkRow> Get_HyperlinksRows(Shape shape)
        {
            var rows = new ObservableCollection<VNCVisioAddIn.Domain.HyperlinkRow>();

            Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionHyperlink];

            var rowCount = section.Count;

            for (short i = 0; i < rowCount; i++)
            {
                VNCVisioAddIn.Domain.HyperlinkRow hyperlinkRow = new VNCVisioAddIn.Domain.HyperlinkRow();

                var row = section[i];

                hyperlinkRow.Name = row.NameU;

                hyperlinkRow.Description = row[(short)VisCellIndices.visHLinkDescription].FormulaU;
                hyperlinkRow.Address = row[(short)VisCellIndices.visHLinkAddress].FormulaU;
                hyperlinkRow.SubAddress = row[(short)VisCellIndices.visHLinkSubAddress].FormulaU;
                hyperlinkRow.ExtraInfo = row[(short)VisCellIndices.visHLinkExtraInfo].FormulaU;
                hyperlinkRow.Frame = row[(short)VisCellIndices.visHLinkFrame].FormulaU;
                hyperlinkRow.SortKey = row[(short)VisCellIndices.visHLinkSortKey].FormulaU;
                hyperlinkRow.NewWindow = row[(short)VisCellIndices.visHLinkNewWin].FormulaU;
                hyperlinkRow.Default = row[(short)VisCellIndices.visHLinkDefault].FormulaU;
                hyperlinkRow.Invisible = row[(short)VisCellIndices.visHLinkInvisible].FormulaU;

                rows.Add(hyperlinkRow);
            }

            return rows;
        }

        public static VNCVisioAddIn.Domain.ScratchRow Get_ScratchRow(Shape shape, short rowNumber)
        {
            VNCVisioAddIn.Domain.ScratchRow row = new VNCVisioAddIn.Domain.ScratchRow();

            Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionScratch];
            Visio.Row sectionRow = section[rowNumber];

            return row;
        }

        internal static ObservableCollection<VNCVisioAddIn.Domain.ScratchRow> Get_ScratchRows(Shape shape)
        {
            var rows = new ObservableCollection<VNCVisioAddIn.Domain.ScratchRow>();

            Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionScratch];

            var rowCount = section.Count;

            for (short i = 0; i < rowCount; i++)
            {
                VNCVisioAddIn.Domain.ScratchRow scratchRow = new VNCVisioAddIn.Domain.ScratchRow();

                var row = section[i];

                scratchRow.Row = $"{i}";

                scratchRow.X = row[(short)VisCellIndices.visScratchX].FormulaU;
                scratchRow.Y = row[(short)VisCellIndices.visScratchY].FormulaU;
                scratchRow.A = row[(short)VisCellIndices.visScratchA].FormulaU;
                scratchRow.B = row[(short)VisCellIndices.visScratchB].FormulaU;
                scratchRow.C = row[(short)VisCellIndices.visScratchC].FormulaU;
                scratchRow.D = row[(short)VisCellIndices.visScratchD].FormulaU;

                rows.Add(scratchRow);
            }

            return rows;
        }

        public static VNCVisioAddIn.Domain.ParagraphRow Get_ParagraphSection(Visio.Shape shape)
        {
            VNCVisioAddIn.Domain.ParagraphRow paragraph = new VNCVisioAddIn.Domain.ParagraphRow();

            Visio.Section paragraphSection = shape.Section[(short)Visio.VisSectionIndices.visSectionParagraph];
            Visio.Row paragraphRow = paragraphSection[0];

            paragraph.IndFirst = paragraphRow[VisCellIndices.visIndentFirst].FormulaU;
            paragraph.IndLeft = paragraphRow[VisCellIndices.visIndentLeft].FormulaU;
            paragraph.IndRight = paragraphRow[VisCellIndices.visIndentRight].FormulaU;
            paragraph.SpLine = paragraphRow[VisCellIndices.visSpaceLine].FormulaU;
            paragraph.SpBefore = paragraphRow[VisCellIndices.visSpaceBefore].FormulaU;
            paragraph.SpAfter = paragraphRow[VisCellIndices.visSpaceAfter].FormulaU;
            paragraph.HAlign = paragraphRow[VisCellIndices.visHorzAlign].FormulaU;
            paragraph.Bullet = paragraphRow[VisCellIndices.visBulletIndex].FormulaU;
            paragraph.BulletString = paragraphRow[VisCellIndices.visBulletString].FormulaU;
            paragraph.BulletFont = paragraphRow[VisCellIndices.visBulletFont].FormulaU;
            paragraph.TextPosAfterBullet = paragraphRow[VisCellIndices.visTextPosAfterBullet].FormulaU;
            paragraph.BulletSize = paragraphRow[VisCellIndices.visBulletFontSize].FormulaU;
            paragraph.Flags = paragraphRow[VisCellIndices.visFlags].FormulaU;

            return paragraph;
        }

        #endregion

        #endregion

        #region Get ShapeSheet Page Section

        internal static VNCVisioAddIn.Domain.PageProperties Get_PageProperties(Shape shape)
        {
            VNCVisioAddIn.Domain.PageProperties row = new VNCVisioAddIn.Domain.PageProperties();

            Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionObject];
            Visio.Row sectionRow = section[(short)Visio.VisRowIndices.visRowPage];

            row.PageWidth = sectionRow[VisCellIndices.visPageWidth].FormulaU;
            row.PageHeight = sectionRow[VisCellIndices.visPageHeight].FormulaU;
            row.PageScale = sectionRow[VisCellIndices.visPageScale].FormulaU;
            row.DrawingScale = sectionRow[VisCellIndices.visPageDrawingScale].FormulaU;
            row.DrawingSizeType = sectionRow[VisCellIndices.visPageDrawSizeType].FormulaU;
            row.DrawingResizeType = sectionRow[VisCellIndices.visPageDrawResizeType].FormulaU;
            row.DrawingScaleType = sectionRow[VisCellIndices.visPageDrawScaleType].FormulaU;
            row.InhibitSnap = sectionRow[VisCellIndices.visPageInhibitSnap].FormulaU;
            row.UIVisibility = sectionRow[VisCellIndices.visPageUIVisibility].FormulaU;
            row.PageLockReplace = sectionRow[VisCellIndices.visPageLockReplace].FormulaU;
            row.PageLockDuplicate = sectionRow[VisCellIndices.visPageLockDuplicate].FormulaU;

            return row;
        }

        internal static VNCVisioAddIn.Domain.PageLayout Get_PageLayout(Shape shape)
        {
            VNCVisioAddIn.Domain.PageLayout row = new VNCVisioAddIn.Domain.PageLayout();

            Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionObject];
            Visio.Row sectionRow = section[(short)Visio.VisRowIndices.visRowPageLayout];

            row.PlaceStyle = sectionRow[VisCellIndices.visPLOPlaceStyle].FormulaU;
            row.PlaceDepth = sectionRow[VisCellIndices.visPLOPlaceDepth].FormulaU;
            row.PlowCode = sectionRow[VisCellIndices.visPLOPlowCode].FormulaU;
            row.ResizePage = sectionRow[VisCellIndices.visPLOResizePage].FormulaU;
            row.DynamicsOff = sectionRow[VisCellIndices.visPLODynamicsOff].FormulaU;
            row.EnableGrid = sectionRow[VisCellIndices.visPLOEnableGrid].FormulaU;
            row.CtrlAsInput = sectionRow[VisCellIndices.visPLOCtrlAsInput].FormulaU;
            row.LineAdjustFrom = sectionRow[VisCellIndices.visPLOLineAdjustFrom].FormulaU;
            row.PlaceFlip = sectionRow[VisCellIndices.visPLOPlaceFlip].FormulaU;
            row.AvoidPageBreaks = sectionRow[VisCellIndices.visPLOAvoidPageBreaks].FormulaU;
            row.BlockSizeX = sectionRow[VisCellIndices.visPLOBlockSizeX].FormulaU;
            row.BlockSizeY = sectionRow[VisCellIndices.visPLOBlockSizeY].FormulaU;
            row.AvenueSizeX = sectionRow[VisCellIndices.visPLOAvenueSizeX].FormulaU;
            row.AvenueSizeY = sectionRow[VisCellIndices.visPLOAvenueSizeY].FormulaU;
            row.RouteStyle = sectionRow[VisCellIndices.visPLORouteStyle].FormulaU;
            row.PageLineJumpDirX = sectionRow[VisCellIndices.visPLOJumpDirX].FormulaU;
            row.PageLineJumpDirY = sectionRow[VisCellIndices.visPLOJumpDirY].FormulaU;
            row.LineAdjustTo = sectionRow[VisCellIndices.visPLOLineAdjustTo].FormulaU;
            row.LineRouteExt = sectionRow[VisCellIndices.visPLOLineRouteExt].FormulaU;
            row.LineToNodeX = sectionRow[VisCellIndices.visPLOLineToNodeX].FormulaU;
            row.LineToNodeY = sectionRow[VisCellIndices.visPLOLineToNodeY].FormulaU;
            row.LineToLineX = sectionRow[VisCellIndices.visPLOLineToLineX].FormulaU;
            row.LineToLineY = sectionRow[VisCellIndices.visPLOLineToLineY].FormulaU;
            row.LineJumpFactorX = sectionRow[VisCellIndices.visPLOJumpFactorX].FormulaU;
            row.LineJumpFactorY = sectionRow[VisCellIndices.visPLOJumpFactorY].FormulaU;
            row.LineJumpCode = sectionRow[VisCellIndices.visPLOJumpCode].FormulaU;
            row.LineJumpStyle = sectionRow[VisCellIndices.visPLOJumpStyle].FormulaU;
            row.PageShapeSplit = sectionRow[VisCellIndices.visPLOSplit].FormulaU;

            return row;
        }

        public static VNCVisioAddIn.Domain.PrintProperties Get_PrintProperties(Shape shape)
        {
            VNCVisioAddIn.Domain.PrintProperties row = new VNCVisioAddIn.Domain.PrintProperties();

            Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionObject];
            Visio.Row sectionRow = section[(short)Visio.VisRowIndices.visRowPrintProperties];

            row.PageLeftMargin = sectionRow[VisCellIndices.visPrintPropertiesLeftMargin].FormulaU;
            row.PageTopMargin = sectionRow[VisCellIndices.visPrintPropertiesTopMargin].FormulaU;
            row.PageRightMargin = sectionRow[VisCellIndices.visPrintPropertiesRightMargin].FormulaU;
            row.PageBottomMargin = sectionRow[VisCellIndices.visPrintPropertiesBottomMargin].FormulaU;
            row.ScaleX = sectionRow[VisCellIndices.visPrintPropertiesScaleX].FormulaU;
            row.ScaleY = sectionRow[VisCellIndices.visPrintPropertiesScaleY].FormulaU;
            row.PagesX= sectionRow[VisCellIndices.visPrintPropertiesPagesX].FormulaU;
            row.PagesY = sectionRow[VisCellIndices.visPrintPropertiesPagesY].FormulaU;
            row.CenterX = sectionRow[VisCellIndices.visPrintPropertiesCenterX].FormulaU;
            row.CenterY = sectionRow[VisCellIndices.visPrintPropertiesCenterY].FormulaU;
            row.OnPage = sectionRow[VisCellIndices.visPrintPropertiesOnPage].FormulaU;
            row.PrintGrid = sectionRow[VisCellIndices.visPrintPropertiesPrintGrid].FormulaU;
            row.PrintPageOrientation = sectionRow[VisCellIndices.visPrintPropertiesPageOrientation].FormulaU;
            row.PaperKind = sectionRow[VisCellIndices.visPrintPropertiesPaperKind].FormulaU;
            row.PaperSource = sectionRow[VisCellIndices.visPrintPropertiesPaperSource].FormulaU;

            return row;
        }

        public static VNCVisioAddIn.Domain.RulerAndGrid Get_RulerAndGrid(Shape shape)
        {
            VNCVisioAddIn.Domain.RulerAndGrid row = new VNCVisioAddIn.Domain.RulerAndGrid();

            Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionObject];
            Visio.Row sectionRow = section[(short)Visio.VisRowIndices.visRowRulerGrid];

            row.XRulerOrigin = sectionRow[VisCellIndices.visXRulerOrigin].FormulaU;
            row.YRulerOrigin = sectionRow[VisCellIndices.visYRulerOrigin].FormulaU;
            row.XRulerDensity = sectionRow[VisCellIndices.visXRulerDensity].FormulaU;
            row.YRulerDensity = sectionRow[VisCellIndices.visYRulerDensity].FormulaU;
            row.XGridOrigin = sectionRow[VisCellIndices.visXGridOrigin].FormulaU;
            row.YGridOrigin = sectionRow[VisCellIndices.visYGridOrigin].FormulaU;
            row.XGridDensity = sectionRow[VisCellIndices.visXGridDensity].FormulaU;
            row.YGridDensity = sectionRow[VisCellIndices.visYGridDensity].FormulaU;
            row.XGridSpacing = sectionRow[VisCellIndices.visXGridSpacing].FormulaU;
            row.YGridSpacing = sectionRow[VisCellIndices.visYGridSpacing].FormulaU;

            return row;
        }

        #endregion

        #region Get ShapeSheet Document Section

        internal static VNCVisioAddIn.Domain.DocumentProperties Get_DocumentProperties(Shape shape)
        {
            VNCVisioAddIn.Domain.DocumentProperties row = new VNCVisioAddIn.Domain.DocumentProperties();

            Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionObject];
            Visio.Row sectionRow = section[(short)Visio.VisRowIndices.visRowDoc];

            row.PreviewQuality = sectionRow[VisCellIndices.visDocPreviewQuality].FormulaU;
            row.OutputFormat = sectionRow[VisCellIndices.visDocOutputFormat].FormulaU;
            row.PreviewScope = sectionRow[VisCellIndices.visDocPreviewScope].FormulaU;
            row.LockPreview = sectionRow[VisCellIndices.visDocLockPreview].FormulaU;
            row.AddMarkup = sectionRow[VisCellIndices.visDocAddMarkup].FormulaU;
            row.ViewMarkup = sectionRow[VisCellIndices.visDocViewMarkup].FormulaU;
            row.DocLangID = sectionRow[VisCellIndices.visDocLangID].FormulaU;
            row.DocLockReplace = sectionRow[VisCellIndices.visDocLockReplace].FormulaU;
            row.NoCoauth = sectionRow[VisCellIndices.visDocNoCoauth].FormulaU;
            row.DocLockDuplicatePage = sectionRow[VisCellIndices.visDocLockDuplicatePage].FormulaU;

            return row;
        }

        #endregion

        public static VNCVisioAddIn.Domain.GeometryRow Get_GeometryRow(Shape shape)
        {
            VNCVisioAddIn.Domain.GeometryRow row = new VNCVisioAddIn.Domain.GeometryRow();

            // TODO(crhodes)
            // Can't find Section Index

            //Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.v];
            //Visio.Row sectionRow = section[0];

            //// TODO(crhodes)
            //// Handle multiple rows

            //row.Action = sectionRow[VisCellIndices.visActionAction].FormulaU;

            return row;
        }

        public static bool HasTextTransformSection(Shape shape)
        {
            bool result = false;

            Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionObject];
            Visio.Row sectionRow = section[(short)Visio.VisRowIndices.visRowTextXForm];

            return result;
        }

        public static VNCVisioAddIn.Domain.BevelProperties Get_BevelProperties(Shape shape)
        {
            VNCVisioAddIn.Domain.BevelProperties row = new VNCVisioAddIn.Domain.BevelProperties();

            Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionObject];
            Visio.Row sectionRow = section[(short)Visio.VisRowIndices.visRowBevelProperties];

            row.BevelTopType = sectionRow[VisCellIndices.visBevelTopType].FormulaU;
            row.BevelTopWidth = sectionRow[VisCellIndices.visBevelTopWidth].FormulaU;
            row.BevelTopHeight = sectionRow[VisCellIndices.visBevelTopHeight].FormulaU;
            row.BevelBottomType = sectionRow[VisCellIndices.visBevelBottomType].FormulaU;
            row.BevelBottomWidth = sectionRow[VisCellIndices.visBevelBottomWidth].FormulaU;
            row.BevelBottomHeight = sectionRow[VisCellIndices.visBevelBottomHeight].FormulaU;
            row.BevelDepthColor = sectionRow[VisCellIndices.visBevelDepthColor].FormulaU;
            row.BevelDepthSize = sectionRow[VisCellIndices.visBevelDepthSize].FormulaU;
            row.BevelContourColor = sectionRow[VisCellIndices.visBevelContourColor].FormulaU;
            row.BevelContourSize = sectionRow[VisCellIndices.visBevelContourSize].FormulaU;
            row.BevelMaterialType = sectionRow[VisCellIndices.visBevelMaterialType].FormulaU;
            row.BevelLightingType = sectionRow[VisCellIndices.visBevelLightingType].FormulaU;
            row.BevelLightingAngle = sectionRow[VisCellIndices.visBevelLightingAngle].FormulaU;

            return row;
        }

        public static VNCVisioAddIn.Domain.ShapeTransform Get_ShapeTransform(Shape shape)
        {
            VNCVisioAddIn.Domain.ShapeTransform row = new VNCVisioAddIn.Domain.ShapeTransform();

            Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionObject];
            Visio.Row sectionRow = section[(short)Visio.VisRowIndices.visRowXFormOut];

            row.Width = sectionRow[VisCellIndices.visXFormWidth].FormulaU;
            row.Height = sectionRow[VisCellIndices.visXFormHeight].FormulaU;
            row.Angle = sectionRow[VisCellIndices.visXFormAngle].FormulaU;
            row.PinX = sectionRow[VisCellIndices.visXFormPinX].FormulaU;
            row.PinY = sectionRow[VisCellIndices.visXFormPinY].FormulaU;
            row.LocPinX = sectionRow[VisCellIndices.visXFormLocPinX].FormulaU;
            row.LocPinY = sectionRow[VisCellIndices.visXFormLocPinY].FormulaU;
            row.FlipX = sectionRow[VisCellIndices.visXFormFlipX].FormulaU;
            row.FlipY = sectionRow[VisCellIndices.visXFormFlipY].FormulaU;
            row.ResizeMode = sectionRow[VisCellIndices.visXFormResizeMode].FormulaU;

            return row;
        }

        public static VNCVisioAddIn.Domain.TextBlockFormat Get_TextBlockFormat(Shape shape)
        {
            VNCVisioAddIn.Domain.TextBlockFormat row = new VNCVisioAddIn.Domain.TextBlockFormat();

            // Shape Transform Section is part of object
            Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionObject];
            Visio.Row sectionRow = section[(short)Visio.VisRowIndices.visRowText];

            // TODO(crhodes)
            // Handle multiple rows

            row.LeftMargin = sectionRow[VisCellIndices.visTxtBlkLeftMargin].FormulaU;
            row.TopMargin = sectionRow[VisCellIndices.visTxtBlkTopMargin].FormulaU;
            row.RightMargin = sectionRow[VisCellIndices.visTxtBlkRightMargin].FormulaU;
            row.BottomMargin = sectionRow[VisCellIndices.visTxtBlkBottomMargin].FormulaU;
            row.TextBkgnd = sectionRow[VisCellIndices.visTxtBlkBkgnd].FormulaU;
            row.TextBkgndTrans = sectionRow[VisCellIndices.visTxtBlkBkgndTrans].FormulaU;
            row.TextDirection = sectionRow[VisCellIndices.visTxtBlkDirection].FormulaU;
            row.VerticalAlign = sectionRow[VisCellIndices.visTxtBlkVerticalAlign].FormulaU;
            row.DefaultTabStop = sectionRow[VisCellIndices.visTxtBlkDefaultTabStop].FormulaU;

            return row;
        }

        public static VNCVisioAddIn.Domain.StyleProperties Get_StyleProperties(Shape shape)
        {
            VNCVisioAddIn.Domain.StyleProperties row = new VNCVisioAddIn.Domain.StyleProperties();

            // Shape Transform Section is part of object
            Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionObject];
            Visio.Row sectionRow = section[(short)Visio.VisRowIndices.visRowStyle];

            //row.LeftMargin = sectionRow[VisCellIndices.visTxtBlkLeftMargin].FormulaU;

            return row;
        }

        public static VNCVisioAddIn.Domain.TabRow Get_TabRow(Shape shape)
        {
            throw new NotImplementedException();
        }

        public static ObservableCollection<VNCVisioAddIn.Domain.ShapeDataRow> Get_ShapeDataRows(Shape shape)
        {
            var rows = new ObservableCollection<VNCVisioAddIn.Domain.ShapeDataRow>();

            Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionProp];

            var rowCount = section.Count;

            for (short i = 0; i < rowCount; i++)
            {
                VNCVisioAddIn.Domain.ShapeDataRow shapeDataRow = new VNCVisioAddIn.Domain.ShapeDataRow();

                var row = section[i];

                shapeDataRow.Name = row.NameU;

                // HACK(crhodes)
                // Trying to find a way to determine if there is a formula in the cell
                // Nothing obvious
                var fooProp = row[(short)VisCellIndices.visCustPropsPrompt];
                var fooStat = row[(short)VisCellIndices.visCustPropsPrompt].Stat;

                var fooResult = row[(short)VisCellIndices.visCustPropsPrompt].Result[VisUnitCodes.visUnitsInval];
                var fooFormula = row[(short)VisCellIndices.visXFormPinY].Formula;
                var fooFormulaU = row[(short)VisCellIndices.visCustPropsPrompt].FormulaU;
                var fooResultStrU = row[(short)VisCellIndices.visCustPropsPrompt].ResultStrU[VisUnitCodes.visUnitsString];

                var fooUnits = row[(short)VisCellIndices.visCustPropsPrompt].Units;

                //shapeDataRow.Label = row[(short)VisCellIndices.visCustPropsLabel].FormulaU;
                //shapeDataRow.Prompt = row[(short)VisCellIndices.visCustPropsPrompt].FormulaU;
                //shapeDataRow.Type = row[(short)VisCellIndices.visCustPropsType].FormulaU;
                //shapeDataRow.Format = row[(short)VisCellIndices.visCustPropsFormat].FormulaU;
                //shapeDataRow.Value = row[(short)VisCellIndices.visCustPropsValue].FormulaU;
                //shapeDataRow.SortKey = row[(short)VisCellIndices.visCustPropsSortKey].FormulaU;
                //shapeDataRow.Invisible = row[(short)VisCellIndices.visCustPropsInvis].FormulaU;
                //shapeDataRow.Ask = row[(short)VisCellIndices.visCustPropsAsk].FormulaU;
                //shapeDataRow.LangID = row[(short)VisCellIndices.visCustPropsLangID].FormulaU;
                //shapeDataRow.Calendar = row[(short)VisCellIndices.visCustPropsCalendar].FormulaU;

                shapeDataRow.Label = row[(short)VisCellIndices.visCustPropsLabel].ResultStrU[VisUnitCodes.visUnitsString];


                shapeDataRow.Prompt = row[(short)VisCellIndices.visCustPropsPrompt].ResultStrU[VisUnitCodes.visUnitsString];
                shapeDataRow.Type = row[(short)VisCellIndices.visCustPropsType].ResultStrU[VisUnitCodes.visUnitsString];
                shapeDataRow.Format = row[(short)VisCellIndices.visCustPropsFormat].ResultStrU[VisUnitCodes.visUnitsString];
                shapeDataRow.Value = row[(short)VisCellIndices.visCustPropsValue].ResultStrU[VisUnitCodes.visUnitsString];
                shapeDataRow.SortKey = row[(short)VisCellIndices.visCustPropsSortKey].ResultStrU[VisUnitCodes.visUnitsString];
                shapeDataRow.Invisible = row[(short)VisCellIndices.visCustPropsInvis].ResultStrU[VisUnitCodes.visUnitsString];
                shapeDataRow.Ask = row[(short)VisCellIndices.visCustPropsAsk].ResultStrU[VisUnitCodes.visUnitsString];
                shapeDataRow.LangID = row[(short)VisCellIndices.visCustPropsLangID].ResultStrU[VisUnitCodes.visUnitsString];
                shapeDataRow.Calendar = row[(short)VisCellIndices.visCustPropsCalendar].ResultStrU[VisUnitCodes.visUnitsString];

                rows.Add(shapeDataRow);
            }

            return rows;
        }

        public static ObservableCollection<VNCVisioAddIn.Domain.UserDefinedCellRow> Get_UserDefinedCellsRows(Shape shape)
        {
            var rows = new ObservableCollection<VNCVisioAddIn.Domain.UserDefinedCellRow>();

            Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionUser];

            var rowCount = section.Count;

            for (short i = 0; i < rowCount; i++)
            {
                var userRow = new VNCVisioAddIn.Domain.UserDefinedCellRow();

                var row = section[i];

                userRow.Name = row.NameU;

                userRow.Value = row[(short)VisCellIndices.visUserValue].FormulaU;
                userRow.Prompt = row[(short)VisCellIndices.visUserPrompt].FormulaU;

                rows.Add(userRow);
            }

            return rows;
        }

        public static VNCVisioAddIn.Domain.TextFieldRow Get_TextFieldRow(Shape shape)
        {
            throw new NotImplementedException();
        }

        #endregion

        #region Set ShapeSheet Section

        #region Document Section
        public static void Set_DocumentProperties_Section(Shape shape, VNCVisioAddIn.Domain.DocumentProperties documentProperties)
        {
            try
            {
                Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionObject];
                Visio.Row sectionRow = section[(short)Visio.VisRowIndices.visRowDoc];

                sectionRow[VisCellIndices.visDocPreviewQuality].FormulaU = documentProperties.PreviewQuality;
                sectionRow[VisCellIndices.visDocOutputFormat].FormulaU = documentProperties.OutputFormat;
                sectionRow[VisCellIndices.visDocPreviewScope].FormulaU = documentProperties.PreviewScope;
                sectionRow[VisCellIndices.visDocLockPreview].FormulaU = documentProperties.LockPreview;
                sectionRow[VisCellIndices.visDocAddMarkup].FormulaU = documentProperties.AddMarkup;
                sectionRow[VisCellIndices.visDocViewMarkup].FormulaU = documentProperties.ViewMarkup;
                sectionRow[VisCellIndices.visDocLangID].FormulaU = documentProperties.DocLangID;
                sectionRow[VisCellIndices.visDocLockReplace].FormulaU = documentProperties.DocLockReplace;
                sectionRow[VisCellIndices.visDocNoCoauth].FormulaU = documentProperties.NoCoauth;
                sectionRow[VisCellIndices.visDocLockDuplicatePage].FormulaU = documentProperties.DocLockDuplicatePage;
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }

        #endregion

        #region Page Section

        public static void Set_PageProperties_Section(Shape shape, VNCVisioAddIn.Domain.PageProperties pageProperties)
        {
            try
            {
                Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionObject];
                Visio.Row sectionRow = section[(short)Visio.VisRowIndices.visRowPage];

                sectionRow[VisCellIndices.visPageWidth].FormulaU = pageProperties.PageWidth;
                sectionRow[VisCellIndices.visPageHeight].FormulaU = pageProperties.PageHeight;
                sectionRow[VisCellIndices.visPageScale].FormulaU = pageProperties.PageScale;
                sectionRow[VisCellIndices.visPageDrawingScale].FormulaU = pageProperties.DrawingScale;
                sectionRow[VisCellIndices.visPageDrawSizeType].FormulaU = pageProperties.DrawingSizeType;
                sectionRow[VisCellIndices.visPageDrawScaleType].FormulaU = pageProperties.DrawingScaleType;
                sectionRow[VisCellIndices.visPageInhibitSnap].FormulaU = pageProperties.InhibitSnap;
                sectionRow[VisCellIndices.visPageUIVisibility].FormulaU = pageProperties.UIVisibility;
                sectionRow[VisCellIndices.visPageLockReplace].FormulaU = pageProperties.PageLockReplace;
                sectionRow[VisCellIndices.visPageLockDuplicate].FormulaU = pageProperties.PageLockDuplicate;
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }

        #endregion

        #region Shape Section


        public static void Set_AdditionalEffectProperties_Section(Shape shape, VNCVisioAddIn.Domain.AdditionalEffectProperties additionalEffectProperties)
        {
            try
            {
                Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionObject];
                Visio.Row sectionRow = section[(short)Visio.VisRowIndices.visRowOtherEffectProperties];

                sectionRow[VisCellIndices.visReflectionTrans].FormulaU = additionalEffectProperties.ReflectionTrans;
                sectionRow[VisCellIndices.visReflectionSize].FormulaU = additionalEffectProperties.ReflectionSize;
                sectionRow[VisCellIndices.visReflectionDist].FormulaU = additionalEffectProperties.ReflectionDist;
                sectionRow[VisCellIndices.visReflectionBlur].FormulaU = additionalEffectProperties.ReflectionBlur;
                sectionRow[VisCellIndices.visSketchEnabled].FormulaU = additionalEffectProperties.SketchEnabled;
                sectionRow[VisCellIndices.visSketchSeed].FormulaU = additionalEffectProperties.SketchSeed;
                sectionRow[VisCellIndices.visSketchAmount].FormulaU = additionalEffectProperties.SketchAmount;
                sectionRow[VisCellIndices.visSketchLineWeight].FormulaU = additionalEffectProperties.SketchLineWeight;
                sectionRow[VisCellIndices.visSketchLineChange].FormulaU = additionalEffectProperties.SketchLineChange;
                sectionRow[VisCellIndices.visSketchFillChange].FormulaU = additionalEffectProperties.SketchFillChange;
                sectionRow[VisCellIndices.visGlowColor].FormulaU = additionalEffectProperties.GlowColor;
                sectionRow[VisCellIndices.visGlowColorTrans].FormulaU = additionalEffectProperties.GlowColorTrans;
                sectionRow[VisCellIndices.visGlowSize].FormulaU = additionalEffectProperties.GlowSize;
                sectionRow[VisCellIndices.visSoftEdgesSize].FormulaU = additionalEffectProperties.SoftEdgesSize;
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }
        
        public static void Set_BevelPropertiesWrapper_Section(Shape shape, VNCVisioAddIn.Domain.BevelProperties bevelProperties)
        {
            try
            {
                Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionObject];
                Visio.Row sectionRow = section[(short)Visio.VisRowIndices.visRowBevelProperties];

                sectionRow[VisCellIndices.visBevelTopType].FormulaU = bevelProperties.BevelTopType;
                sectionRow[VisCellIndices.visBevelTopWidth].FormulaU = bevelProperties.BevelTopWidth;
                sectionRow[VisCellIndices.visBevelTopHeight].FormulaU = bevelProperties.BevelTopHeight;
                sectionRow[VisCellIndices.visBevelBottomType].FormulaU = bevelProperties.BevelBottomType;
                sectionRow[VisCellIndices.visBevelBottomWidth].FormulaU = bevelProperties.BevelBottomWidth;
                sectionRow[VisCellIndices.visBevelBottomHeight].FormulaU = bevelProperties.BevelBottomHeight;
                sectionRow[VisCellIndices.visBevelDepthColor].FormulaU = bevelProperties.BevelDepthColor;
                sectionRow[VisCellIndices.visBevelDepthSize].FormulaU = bevelProperties.BevelDepthSize;
                sectionRow[VisCellIndices.visBevelContourColor].FormulaU = bevelProperties.BevelContourColor;
                sectionRow[VisCellIndices.visBevelContourSize].FormulaU = bevelProperties.BevelContourSize;
                sectionRow[VisCellIndices.visBevelMaterialType].FormulaU = bevelProperties.BevelMaterialType;
                sectionRow[VisCellIndices.visBevelLightingType].FormulaU = bevelProperties.BevelLightingType;
                sectionRow[VisCellIndices.visBevelLightingAngle].FormulaU = bevelProperties.BevelLightingAngle;
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }

        public static void Set_ChangeShapeBehavior_Section(Shape shape, VNCVisioAddIn.Domain.ChangeShapeBehavior changeShapeBehavior)
        {
            try
            {
                Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionObject];
                Visio.Row sectionRow = section[(short)Visio.VisRowIndices.visRowReplaceBehaviors];

                sectionRow[VisCellIndices.visReplaceLockShapeData].FormulaU = changeShapeBehavior.ReplaceLockShapeData;
                sectionRow[VisCellIndices.visReplaceLockText].FormulaU = changeShapeBehavior.ReplaceLockText;
                sectionRow[VisCellIndices.visReplaceLockFormat].FormulaU = changeShapeBehavior.ReplaceLockFormat;
                sectionRow[VisCellIndices.visReplaceCopyCells].FormulaU = changeShapeBehavior.ReplaceCopyCells;
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }

        public static void Set_FillFormat_Section(Shape shape, VNCVisioAddIn.Domain.FillFormat fillFormat)
        {
            Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionObject];
            Visio.Row sectionRow = section[(short)Visio.VisRowIndices.visRowFill];

            sectionRow[VisCellIndices.visFillForegnd].FormulaU = fillFormat.FillForegnd;
            sectionRow[VisCellIndices.visFillShdwForegnd].FormulaU = fillFormat.ShdwForegnd;
            sectionRow[VisCellIndices.visFillShdwType].FormulaU = fillFormat.ShapeShdwType;
            sectionRow[VisCellIndices.visFillForegndTrans].FormulaU = fillFormat.FillForegndTrans;
            sectionRow[VisCellIndices.visFillShdwForegndTrans].FormulaU = fillFormat.ShdwForegndTrans;
            sectionRow[VisCellIndices.visFillShdwObliqueAngle].FormulaU = fillFormat.ShapeShdwObliqueAngle;
            sectionRow[VisCellIndices.visFillBkgnd].FormulaU = fillFormat.FillBkgnd;
            sectionRow[VisCellIndices.visFillShdwPattern].FormulaU = fillFormat.ShdwPattern;
            sectionRow[VisCellIndices.visFillShdwScaleFactor].FormulaU = fillFormat.ShapeShdwScaleFactor;
            sectionRow[VisCellIndices.visFillBkgndTrans].FormulaU = fillFormat.FillBkgndTrans;
            sectionRow[VisCellIndices.visFillShdwOffsetX].FormulaU = fillFormat.ShapeShdwOffsetX;
            sectionRow[VisCellIndices.visFillShdwOffsetY].FormulaU = fillFormat.ShapeShdwOffsetY;
            sectionRow[VisCellIndices.visFillShdwShow].FormulaU = fillFormat.ShapeShdwBlur;
            sectionRow[VisCellIndices.visFillShdwPattern].FormulaU = fillFormat.FillPattern;
            sectionRow[VisCellIndices.visFillShdwShow].FormulaU = fillFormat.ShapeShdwShow;
        }

        public static void Set_GlueInfo_Section(Shape shape, VNCVisioAddIn.Domain.GlueInfo glueInfo)
        {
            try
            {
                Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionObject];
                Visio.Row sectionRow = section[(short)Visio.VisRowIndices.visRowMisc];

                sectionRow[VisCellIndices.visBegTrigger].FormulaU = glueInfo.BegTrigger;
                sectionRow[VisCellIndices.visEndTrigger].FormulaU = glueInfo.EndTrigger;
                sectionRow[VisCellIndices.visGlueType].FormulaU = glueInfo.GlueType;
                sectionRow[VisCellIndices.visWalkPref].FormulaU = glueInfo.WalkPreference;
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }

        public static void Set_GradientProperties_Section(Shape shape, VNCVisioAddIn.Domain.GradientProperties gradientProperties)
        {
            try
            {
                Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionObject];
                Visio.Row sectionRow = section[(short)Visio.VisRowIndices.visRowGradientProperties];

                sectionRow[VisCellIndices.visLineGradientDir].FormulaU = gradientProperties.LineGradientDir;
                sectionRow[VisCellIndices.visLineGradientAngle].FormulaU = gradientProperties.LineGradientAngle;
                sectionRow[VisCellIndices.visFillGradientDir].FormulaU = gradientProperties.FillGradientDir;
                sectionRow[VisCellIndices.visFillGradientAngle].FormulaU = gradientProperties.FillGradientAngle;
                sectionRow[VisCellIndices.visLineGradientEnabled].FormulaU = gradientProperties.LineGradientEnabled;
                sectionRow[VisCellIndices.visFillGradientEnabled].FormulaU = gradientProperties.FillGradientEnabled;
                sectionRow[VisCellIndices.visRotateGradientWithShape].FormulaU = gradientProperties.RotateGradientWithShape;
                sectionRow[VisCellIndices.visUseGroupGradient].FormulaU = gradientProperties.UseGroupGradient;
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }

        public static void Set_GroupProperties_Section(Shape shape, VNCVisioAddIn.Domain.GroupProperties groupProperties)
        {
            Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionObject];
            Visio.Row sectionRow = section[(short)Visio.VisRowIndices.visRowGroup];

            groupProperties.SelectMode = sectionRow[VisCellIndices.visGroupSelectMode].FormulaU;
            groupProperties.IsTextEditTarget = sectionRow[VisCellIndices.visGroupIsTextEditTarget].FormulaU;
            groupProperties.IsDropTarget = sectionRow[VisCellIndices.visGroupIsDropTarget].FormulaU;
            groupProperties.DisplayMode = sectionRow[VisCellIndices.visGroupDisplayMode].FormulaU;
            groupProperties.IsSnapTarget = sectionRow[VisCellIndices.visGroupIsSnapTarget].FormulaU;
            groupProperties.DontMoveChildren = sectionRow[VisCellIndices.visGroupDontMoveChildren].FormulaU;
        }

        public static void Set_ImageProperties_Section(Shape shape, VNCVisioAddIn.Domain.ImageProperties imageProperties)
        {
            try
            {
                Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionObject];
                Visio.Row sectionRow = section[(short)Visio.VisRowIndices.visRowImage];

                sectionRow[VisCellIndices.visImageContrast].FormulaU = imageProperties.Contrast;
                sectionRow[VisCellIndices.visImageGamma].FormulaU = imageProperties.Gamma;
                sectionRow[VisCellIndices.visImageSharpen].FormulaU = imageProperties.Sharpen;
                sectionRow[VisCellIndices.visImageBrightness].FormulaU = imageProperties.Brightness;
                sectionRow[VisCellIndices.visImageBlur].FormulaU = imageProperties.Blur;
                sectionRow[VisCellIndices.visImageDenoise].FormulaU = imageProperties.Denoise;
                imageProperties.Transparency = sectionRow[VisCellIndices.visImageTransparency].FormulaU;
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }

        public static void Set_LayerMembership_Section(Shape shape, VNCVisioAddIn.Domain.LayerMembership layerMembership)
        {
            try
            {
                Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionObject];
                Visio.Row sectionRow = section[(short)Visio.VisRowIndices.visRowLayerMem];

                sectionRow[VisCellIndices.visLayerMember].FormulaU = layerMembership.Name;
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }

        public static void Set_LineFormat_Section(Shape shape, VNCVisioAddIn.Domain.LineFormat lineFormat)
        {
            try
            {
                Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionObject];
                Visio.Row sectionRow = section[(short)Visio.VisRowIndices.visRowLine];

                sectionRow[VisCellIndices.visLinePattern].FormulaU = lineFormat.LinePattern;
                sectionRow[VisCellIndices.visLineWeight].FormulaU = lineFormat.LineWeight;
                sectionRow[VisCellIndices.visLineColor].FormulaU = lineFormat.LineColor;
                sectionRow[VisCellIndices.visLineEndCap].FormulaU = lineFormat.LineCap;
                sectionRow[VisCellIndices.visLineBeginArrow].FormulaU = lineFormat.BeginArrow;
                sectionRow[VisCellIndices.visLineEndArrow].FormulaU = lineFormat.EndArrow;
                sectionRow[VisCellIndices.visLineColorTrans].FormulaU = lineFormat.LineColorTrans;
                sectionRow[VisCellIndices.visCompoundType].FormulaU = lineFormat.CompoundType;
                sectionRow[VisCellIndices.visLineBeginArrowSize].FormulaU = lineFormat.BeginArrowSize;
                sectionRow[VisCellIndices.visLineEndArrowSize].FormulaU = lineFormat.EndArrowSize;
                sectionRow[VisCellIndices.visLineRounding].FormulaU = lineFormat.Rounding;
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }

        internal static void Set_Miscellaneous_Section(Shape shape, VNCVisioAddIn.Domain.Miscellaneous miscellaneous)
        {
            try
            {
                Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionObject];
                Visio.Row sectionRow = section[(short)Visio.VisRowIndices.visRowMisc];

                sectionRow[VisCellIndices.visNoObjHandles].FormulaU = miscellaneous.NoObjHandles;
                sectionRow[VisCellIndices.visNoCtlHandles].FormulaU = miscellaneous.NoCtlHandle;
                sectionRow[VisCellIndices.visNoAlignBox].FormulaU = miscellaneous.NoAlignBox;
                sectionRow[VisCellIndices.visNonPrinting].FormulaU = miscellaneous.NonPrinting;
                sectionRow[VisCellIndices.visObjLangID].FormulaU = miscellaneous.LangID;
                sectionRow[VisCellIndices.visHideText].FormulaU = miscellaneous.HideText;
                sectionRow[VisCellIndices.visUpdateAlignBox].FormulaU = miscellaneous.UpdateAlignBox;
                sectionRow[VisCellIndices.visDynFeedback].FormulaU = miscellaneous.DynFeedback;
                sectionRow[VisCellIndices.visNoLiveDynamics].FormulaU = miscellaneous.NoLiveDynamics;
                sectionRow[VisCellIndices.visObjCalendar].FormulaU = miscellaneous.Calendar;
                sectionRow[VisCellIndices.visLOFlags].FormulaU = miscellaneous.ObjType;
                sectionRow[VisCellIndices.visDropSource].FormulaU = miscellaneous.IsDropSource;
                sectionRow[VisCellIndices.visComment].FormulaU = miscellaneous.Comment;
                sectionRow[VisCellIndices.visObjDropOnPageScale].FormulaU = miscellaneous.DropOnPageScale;
                sectionRow[VisCellIndices.visObjLocalizeMerge].FormulaU = miscellaneous.LocalizeMerge;
                sectionRow[VisCellIndices.visObjNoProofing].FormulaU = miscellaneous.NoProofing;
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }

        public static void Set_OneDEndPoints_Section(Shape shape, VNCVisioAddIn.Domain.OneDEndPoints oneDEndPoints)
        {
            try
            {
                Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionObject];
                Visio.Row sectionRow = section[(short)Visio.VisRowIndices.visRowXForm1D];

                sectionRow[VisCellIndices.vis1DBeginX].FormulaU = oneDEndPoints.BeginX;
                sectionRow[VisCellIndices.vis1DBeginY].FormulaU = oneDEndPoints.BeginY;
                sectionRow[VisCellIndices.vis1DEndX].FormulaU = oneDEndPoints.EndX;
                sectionRow[VisCellIndices.vis1DEndY].FormulaU = oneDEndPoints.EndY;
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }

        public static void Set_PageLayout_Section(Shape shape, VNCVisioAddIn.Domain.PageLayout pageLayout)
        {
            try
            {
                Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionObject];
                Visio.Row sectionRow = section[(short)Visio.VisRowIndices.visRowPageLayout];

                sectionRow[VisCellIndices.visPLOPlaceStyle].FormulaU = pageLayout.PlaceStyle;
                sectionRow[VisCellIndices.visPLOPlaceDepth].FormulaU = pageLayout.PlaceDepth;
                sectionRow[VisCellIndices.visPLOPlowCode].FormulaU = pageLayout.PlowCode;
                sectionRow[VisCellIndices.visPLOResizePage].FormulaU = pageLayout.ResizePage;
                sectionRow[VisCellIndices.visPLODynamicsOff].FormulaU = pageLayout.DynamicsOff;
                sectionRow[VisCellIndices.visPLOEnableGrid].FormulaU = pageLayout.EnableGrid;
                sectionRow[VisCellIndices.visPLOCtrlAsInput].FormulaU = pageLayout.CtrlAsInput;
                sectionRow[VisCellIndices.visPLOLineAdjustFrom].FormulaU = pageLayout.LineAdjustFrom;
                sectionRow[VisCellIndices.visPLOPlaceFlip].FormulaU = pageLayout.PlaceFlip;
                sectionRow[VisCellIndices.visPLOAvoidPageBreaks].FormulaU = pageLayout.AvoidPageBreaks;
                sectionRow[VisCellIndices.visPLOBlockSizeX].FormulaU = pageLayout.BlockSizeX;
                sectionRow[VisCellIndices.visPLOBlockSizeY].FormulaU = pageLayout.BlockSizeY;
                sectionRow[VisCellIndices.visPLOAvenueSizeX].FormulaU = pageLayout.AvenueSizeX;
                sectionRow[VisCellIndices.visPLOAvenueSizeY].FormulaU = pageLayout.AvenueSizeY;
                sectionRow[VisCellIndices.visPLORouteStyle].FormulaU = pageLayout.RouteStyle;
                sectionRow[VisCellIndices.visPLOJumpDirX].FormulaU = pageLayout.PageLineJumpDirX;
                sectionRow[VisCellIndices.visPLOJumpDirY].FormulaU = pageLayout.PageLineJumpDirY;
                sectionRow[VisCellIndices.visPLOLineAdjustTo].FormulaU = pageLayout.LineAdjustTo;
                sectionRow[VisCellIndices.visPLOLineRouteExt].FormulaU = pageLayout.LineRouteExt;
                sectionRow[VisCellIndices.visPLOLineToNodeX].FormulaU = pageLayout.LineToNodeX;
                sectionRow[VisCellIndices.visPLOLineToNodeY].FormulaU = pageLayout.LineToNodeY;
                sectionRow[VisCellIndices.visPLOLineToLineX].FormulaU = pageLayout.LineToLineX;
                sectionRow[VisCellIndices.visPLOLineToLineY].FormulaU = pageLayout.LineToLineY;
                sectionRow[VisCellIndices.visPLOJumpFactorX].FormulaU = pageLayout.LineJumpFactorX;
                sectionRow[VisCellIndices.visPLOJumpFactorY].FormulaU = pageLayout.LineJumpFactorY;
                sectionRow[VisCellIndices.visPLOJumpCode].FormulaU = pageLayout.LineJumpCode;
                sectionRow[VisCellIndices.visPLOJumpStyle].FormulaU = pageLayout.LineJumpStyle;
                sectionRow[VisCellIndices.visPLOSplit].FormulaU = pageLayout.PageShapeSplit;
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }

        public static void Set_Paragraph_Section(Visio.Shape shape, VNCVisioAddIn.Domain.ParagraphRow paragraph)
        {
            try
            {
                Visio.Section paragraphSection = shape.Section[(short)Visio.VisSectionIndices.visSectionParagraph];
                Visio.Row paragraphRow = paragraphSection[0];

                paragraphRow[VisCellIndices.visIndentFirst].FormulaForceU = paragraph.IndFirst;

                paragraphRow[VisCellIndices.visIndentLeft].FormulaForceU = paragraph.IndLeft;
                paragraphRow[VisCellIndices.visIndentRight].FormulaForceU = paragraph.IndRight;
                paragraphRow[VisCellIndices.visSpaceLine].FormulaForceU = paragraph.SpLine;
                paragraphRow[VisCellIndices.visSpaceBefore].FormulaForceU = paragraph.SpBefore;
                paragraphRow[VisCellIndices.visSpaceAfter].FormulaForceU = paragraph.SpAfter;
                paragraphRow[VisCellIndices.visHorzAlign].FormulaForceU = paragraph.HAlign;
                paragraphRow[VisCellIndices.visBulletIndex].FormulaForceU = paragraph.Bullet;
                paragraphRow[VisCellIndices.visBulletString].FormulaForceU = paragraph.BulletString;
                paragraphRow[VisCellIndices.visBulletFont].FormulaForceU = paragraph.BulletFont;
                paragraphRow[VisCellIndices.visTextPosAfterBullet].FormulaForceU = paragraph.TextPosAfterBullet;
                paragraphRow[VisCellIndices.visBulletFontSize].FormulaForceU = paragraph.BulletSize;
                paragraphRow[VisCellIndices.visFlags].FormulaForceU = paragraph.Flags;
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }

        public static void Set_PrintProperties_Section(Shape shape, VNCVisioAddIn.Domain.PrintProperties printProperties)
        {
            try
            {
                Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionObject];
                Visio.Row sectionRow = section[(short)Visio.VisRowIndices.visRowPrintProperties];

                sectionRow[VisCellIndices.visPrintPropertiesLeftMargin].FormulaU = printProperties.PageLeftMargin;
                sectionRow[VisCellIndices.visPrintPropertiesTopMargin].FormulaU = printProperties.PageTopMargin;
                sectionRow[VisCellIndices.visPrintPropertiesRightMargin].FormulaU = printProperties.PageRightMargin;
                sectionRow[VisCellIndices.visPrintPropertiesBottomMargin].FormulaU = printProperties.PageBottomMargin;
                sectionRow[VisCellIndices.visPrintPropertiesScaleX].FormulaU = printProperties.ScaleX;
                sectionRow[VisCellIndices.visPrintPropertiesScaleY].FormulaU = printProperties.ScaleY;
                sectionRow[VisCellIndices.visPrintPropertiesPagesX].FormulaU = printProperties.PagesX;
                sectionRow[VisCellIndices.visPrintPropertiesPagesY].FormulaU = printProperties.PagesY;
                sectionRow[VisCellIndices.visPrintPropertiesCenterX].FormulaU = printProperties.CenterX;
                sectionRow[VisCellIndices.visPrintPropertiesCenterY].FormulaU = printProperties.CenterY;
                sectionRow[VisCellIndices.visPrintPropertiesOnPage].FormulaU = printProperties.OnPage;
                sectionRow[VisCellIndices.visPrintPropertiesPrintGrid].FormulaU = printProperties.PrintGrid;
                sectionRow[VisCellIndices.visPrintPropertiesPageOrientation].FormulaU = printProperties.PrintPageOrientation;
                sectionRow[VisCellIndices.visPrintPropertiesPaperKind].FormulaU = printProperties.PaperKind;
                sectionRow[VisCellIndices.visPrintPropertiesPaperSource].FormulaU = printProperties.PaperSource;
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }

        public static void Set_Protection_Section(Shape shape, VNCVisioAddIn.Domain.Protection protection)
        {
            try
            {
                Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionObject];
                Visio.Row sectionRow = section[(short)Visio.VisRowIndices.visRowLock];

                sectionRow[VisCellIndices.visLockWidth].FormulaU = protection.LockWidth;
                sectionRow[VisCellIndices.visLockHeight].FormulaU = protection.LockHeight;
                sectionRow[VisCellIndices.visLockAspect].FormulaU = protection.LockAspect;
                sectionRow[VisCellIndices.visLockMoveX].FormulaU = protection.LockMoveX;
                sectionRow[VisCellIndices.visLockMoveY].FormulaU = protection.LockMoveY;
                sectionRow[VisCellIndices.visLockRotate].FormulaU = protection.LockRotate;
                sectionRow[VisCellIndices.visLockBegin].FormulaU = protection.LockBegin;
                sectionRow[VisCellIndices.visLockReplace].FormulaU = protection.LockReplace;
                sectionRow[VisCellIndices.visLockEnd].FormulaU = protection.LockEnd;
                sectionRow[VisCellIndices.visLockDelete].FormulaU = protection.LockDelete;
                sectionRow[VisCellIndices.visLockSelect].FormulaU = protection.LockSelect;
                sectionRow[VisCellIndices.visLockFormat].FormulaU = protection.LockFormat;
                sectionRow[VisCellIndices.visLockCustProp].FormulaU = protection.LockCustProp;
                sectionRow[VisCellIndices.visLockTextEdit].FormulaU = protection.LockTextEdit;
                sectionRow[VisCellIndices.visLockVtxEdit].FormulaU = protection.LockVtxEdit;
                sectionRow[VisCellIndices.visLockThemeIndex].FormulaU = protection.LockThemeIndex;
                sectionRow[VisCellIndices.visLockCrop].FormulaU = protection.LockCrop;
                sectionRow[VisCellIndices.visLockGroup].FormulaU = protection.LockGroup;
                sectionRow[VisCellIndices.visLockCalcWH].FormulaU = protection.LockCalcWH;
                sectionRow[VisCellIndices.visLockFromGroupFormat].FormulaU = protection.LockFromGroupFormat;
                sectionRow[VisCellIndices.visLockThemeColors].FormulaU = protection.LockThemeColors;
                sectionRow[VisCellIndices.visLockThemeEffects].FormulaU = protection.LockThemeEffects;
                sectionRow[VisCellIndices.visLockThemeConnectors].FormulaU = protection.LockThemeConnectors;
                sectionRow[VisCellIndices.visLockThemeFonts].FormulaU = protection.LockThemeFonts;
                sectionRow[VisCellIndices.visLockVariation].FormulaU = protection.LockVariation;
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }

        public static void Set_QuickStyle_Section(Shape shape, VNCVisioAddIn.Domain.QuickStyle quickStyle)
        {
            try
            {
                Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionObject];
                Visio.Row sectionRow = section[(short)Visio.VisRowIndices.visRowQuickStyleProperties];

                sectionRow[VisCellIndices.visQuickStyleLineMatrix].FormulaU = quickStyle.QuickStyleLineMatrix;
                sectionRow[VisCellIndices.visQuickStyleLineColor].FormulaU = quickStyle.QuickStyleLineColor;
                sectionRow[VisCellIndices.visQuickStyleFontColor].FormulaU = quickStyle.QuickStyleFontColor;
                sectionRow[VisCellIndices.visQuickStyleVariation].FormulaU = quickStyle.QuickStyleVariation;
                sectionRow[VisCellIndices.visQuickStyleFillMatrix].FormulaU = quickStyle.QuickStyleFillMatrix;
                sectionRow[VisCellIndices.visQuickStyleFontMatrix].FormulaU = quickStyle.QuickStyleFontMatrix;
                sectionRow[VisCellIndices.visQuickStyleEffectsMatrix].FormulaU = quickStyle.QuickStyleEffectsMatrix;
                sectionRow[VisCellIndices.visQuickStyleShadowColor].FormulaU = quickStyle.QuickStyleShadowColor;
                sectionRow[VisCellIndices.visQuickStyleType].FormulaU = quickStyle.QuickStyleType;
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }

        public static void Set_RulerAndGrid_Section(Shape shape, VNCVisioAddIn.Domain.RulerAndGrid rulerAndGrid)
        {
            try
            {
                Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionObject];
                Visio.Row sectionRow = section[(short)Visio.VisRowIndices.visRowRulerGrid];

                sectionRow[VisCellIndices.visXRulerOrigin].FormulaU = rulerAndGrid.XRulerOrigin;
                sectionRow[VisCellIndices.visYRulerOrigin].FormulaU = rulerAndGrid.YRulerOrigin;
                sectionRow[VisCellIndices.visXRulerDensity].FormulaU = rulerAndGrid.XRulerDensity;
                sectionRow[VisCellIndices.visYRulerDensity].FormulaU = rulerAndGrid.YRulerDensity;
                sectionRow[VisCellIndices.visXGridOrigin].FormulaU = rulerAndGrid.XGridOrigin;
                sectionRow[VisCellIndices.visYGridOrigin].FormulaU = rulerAndGrid.YGridOrigin;
                sectionRow[VisCellIndices.visXGridDensity].FormulaU = rulerAndGrid.XGridDensity;
                sectionRow[VisCellIndices.visYGridDensity].FormulaU = rulerAndGrid.YGridDensity;
                sectionRow[VisCellIndices.visXGridSpacing].FormulaU = rulerAndGrid.XGridSpacing;
                sectionRow[VisCellIndices.visYGridSpacing].FormulaU = rulerAndGrid.YGridSpacing;
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }

        public static void Set_ShapeLayout_Section(Shape shape, VNCVisioAddIn.Domain.ShapeLayout shapeLayout)
        {
            try
            {
                Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionObject];
                Visio.Row sectionRow = section[(short)Visio.VisRowIndices.visRowShapeLayout];

                sectionRow[VisCellIndices.visSLOPermX].FormulaU = shapeLayout.ShapePermeableX;
                sectionRow[VisCellIndices.visSLOPermY].FormulaU = shapeLayout.ShapePermeableY;
                sectionRow[VisCellIndices.visSLOFixedCode].FormulaU = shapeLayout.ShapeFixedCode;
                sectionRow[VisCellIndices.visSLOJumpDirX].FormulaU = shapeLayout.ConLineJumpDirX;
                sectionRow[VisCellIndices.visSLOJumpDirY].FormulaU = shapeLayout.ConLineJumpDirY;
                sectionRow[VisCellIndices.visSLOJumpCode].FormulaU = shapeLayout.ConLineJumpCode;
                sectionRow[VisCellIndices.visSLOPlaceFlip].FormulaU = shapeLayout.ShapePlaceFlip;
                sectionRow[VisCellIndices.visSLOPlaceStyle].FormulaU = shapeLayout.ShapePlaceStyle;
                sectionRow[VisCellIndices.visSLOPlowCode].FormulaU = shapeLayout.ShapePlowCode;
                sectionRow[VisCellIndices.visSLOJumpStyle].FormulaU = shapeLayout.ConLineJumpStyle;
                sectionRow[VisCellIndices.visSLOLineRouteExt].FormulaU = shapeLayout.ConLineRouteExt;
                sectionRow[VisCellIndices.visSLODisplayLevel].FormulaU = shapeLayout.DisplayLevel;
                sectionRow[VisCellIndices.visSLORouteStyle].FormulaU = shapeLayout.ShapeRouteStyle;
                sectionRow[VisCellIndices.visSLOConFixedCode].FormulaU = shapeLayout.ConFixedCode;
                sectionRow[VisCellIndices.visSLOSplit].FormulaU = shapeLayout.ShapeSplit;
                sectionRow[VisCellIndices.visSLOSplittable].FormulaU = shapeLayout.ShapeSplittable;
                sectionRow[VisCellIndices.visSLORelationships].FormulaU = shapeLayout.Relationships;
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }

        public static void Set_ShapeTransform_Section(Shape shape, VNCVisioAddIn.Domain.ShapeTransform shapeTransform)
        {
            try
            {
                Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionObject];
                Visio.Row sectionRow = section[(short)Visio.VisRowIndices.visRowXFormOut];

                sectionRow[VisCellIndices.visXFormWidth].FormulaU = shapeTransform.Width;
                sectionRow[VisCellIndices.visXFormHeight].FormulaU = shapeTransform.Height;
                sectionRow[VisCellIndices.visXFormAngle].FormulaU = shapeTransform.Angle;
                sectionRow[VisCellIndices.visXFormPinX].FormulaU = shapeTransform.PinX;
                sectionRow[VisCellIndices.visXFormPinY].FormulaU = shapeTransform.PinY;
                sectionRow[VisCellIndices.visXFormLocPinX].FormulaU = shapeTransform.LocPinX;
                sectionRow[VisCellIndices.visXFormLocPinY].FormulaU = shapeTransform.LocPinY;
                sectionRow[VisCellIndices.visXFormFlipX].FormulaU = shapeTransform.FlipX;
                sectionRow[VisCellIndices.visXFormFlipY].FormulaU = shapeTransform.FlipY;
                sectionRow[VisCellIndices.visXFormResizeMode].FormulaU = shapeTransform.ResizeMode;
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }

        public static void Set_TextTransform_Section(Shape shape, VNCVisioAddIn.Domain.TextTransform textTransform)
        {
            try
            {
                Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionObject];
                Visio.Row sectionRow = section[(short)Visio.VisRowIndices.visRowTextXForm];

                sectionRow[VisCellIndices.visXFormWidth].FormulaForceU = textTransform.TxtWidth;
                sectionRow[VisCellIndices.visXFormHeight].FormulaForceU = textTransform.TxtHeight;
                sectionRow[VisCellIndices.visXFormAngle].FormulaForceU = textTransform.TxtAngle;
                sectionRow[VisCellIndices.visXFormPinX].FormulaForceU = textTransform.TxtPinX;
                sectionRow[VisCellIndices.visXFormPinY].FormulaForceU = textTransform.TxtPinY;
                sectionRow[VisCellIndices.visXFormLocPinX].FormulaForceU = textTransform.TxtLocPinX;
                sectionRow[VisCellIndices.visXFormLocPinY].FormulaForceU = textTransform.TxtLocPinY;
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }

        public static void Set_ThemeProperties_Section(Shape shape, VNCVisioAddIn.Domain.ThemeProperties themeProperties)
        {
            try
            {
                Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionObject];
                Visio.Row sectionRow = section[(short)Visio.VisRowIndices.visRowThemeProperties];

                sectionRow[VisCellIndices.visConnectorSchemeIndex].FormulaU = themeProperties.ConnectorSchemeIndex;
                sectionRow[VisCellIndices.visEffectSchemeIndex].FormulaU = themeProperties.EffectSchemeIndex;
                sectionRow[VisCellIndices.visColorSchemeIndex].FormulaU = themeProperties.ColorSchemeIndex;
                sectionRow[VisCellIndices.visFontSchemeIndex].FormulaU = themeProperties.FontSchemeIndex;
                sectionRow[VisCellIndices.visThemeIndex].FormulaU = themeProperties.ThemeIndex;
                sectionRow[VisCellIndices.visVariationColorIndex].FormulaU = themeProperties.VariationColorIndex;
                sectionRow[VisCellIndices.visVariationStyleIndex].FormulaU = themeProperties.VariationStyleIndex;
                sectionRow[VisCellIndices.visEmbellishmentIndex].FormulaU = themeProperties.EmbellishmentIndex;
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }

        public static void Set_ThreeDRotationProperties_Section(Shape shape, VNCVisioAddIn.Domain.ThreeDRotationProperties threeDRotationProperties)
        {
            try
            {
                Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.visSectionObject];
                Visio.Row sectionRow = section[(short)Visio.VisRowIndices.visRow3DRotationProperties];

                sectionRow[VisCellIndices.visRotationXAngle].FormulaU = threeDRotationProperties.RotationXAngle;
                sectionRow[VisCellIndices.visRotationYAngle].FormulaU = threeDRotationProperties.RotationYAngle;
                sectionRow[VisCellIndices.visRotationZAngle].FormulaU = threeDRotationProperties.RotationZAngle;
                sectionRow[VisCellIndices.visRotationType].FormulaU = threeDRotationProperties.RotationType;
                sectionRow[VisCellIndices.visPerspective].FormulaU = threeDRotationProperties.Perspective;
                sectionRow[VisCellIndices.visDistanceFromGround].FormulaU = threeDRotationProperties.DistanceFromGround;
                sectionRow[VisCellIndices.visKeepTextFlat].FormulaU = threeDRotationProperties.KeepTextFlat;
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }

        #endregion


        #endregion

    }
}
