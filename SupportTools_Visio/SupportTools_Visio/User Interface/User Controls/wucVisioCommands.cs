using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Xml.Linq;
using SupportTools_Visio.Domain;
using ExcelHlp = VNC.ExcelHelper;
using LTE = LinqToExcel;
using Visio = Microsoft.Office.Interop.Visio;
using VisioHlp = VNC.Visio.VSTOAddIn;
using XL = Microsoft.Office.Interop.Excel;
using VNCVisioAddIn = VNC.Visio.VSTOAddIn;

namespace SupportTools_Visio.User_Interface.User_Controls
{
    /// <summary>
    /// Interaction logic for VisioCommands.xaml
    /// </summary>
    public partial class wucVisioCommands : UserControl
    {
        public class TestData
        {
            public string Col1 { get; set; }
            public string Col2 { get; set; }
            public string Col3 { get; set; }
            public string Col4 { get; set; }
            public string Col5 { get; set; }
        }

        #region Event Handlers

        private void btnExecuteCommand_Click(object sender, RoutedEventArgs e)
        {       
            ParseCommand( XElement.Parse(teCommandElements.Text));
        }

        private void btnFive_Click(object sender, RoutedEventArgs e)
        {
            UseLinqToExcel();
        }

        private void btnFour_Click(object sender, RoutedEventArgs e)
        {
            LoadExcelTable();
        }

        private void btnThree_Click(object sender, RoutedEventArgs e)
        {
            UseExcelDataReader();
        }

        private void btnTwo_Click(object sender, RoutedEventArgs e)
        {
            LoadExcelFile();
        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Hello there");
        }

        private void visioCommand_Picker_ControlChanged()
        {
            //var command = visioCommand_Picker.Command;
            //teCommandElements.Text = command.ToString();
            //ParseCommand(command);
        }



        #endregion

        #region Constructors and Load

        public wucVisioCommands()
        {
            InitializeComponent();
            LoadControlContents();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {

        }

        private void UserControl_Unloaded(object sender, RoutedEventArgs e)
        {

        }

        #endregion

        #region Private Methods

        private void LoadControlContents()
        {
            try
            {
                //visioCommand_Picker.PopulateControlFromFile(Common.cCONFIG_FILE);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        #region XML Commands

        private void ParseCommand(System.Xml.Linq.XElement commandElement)
        {
            // Wrap a big, OMG, what have I done ???, undo around the whole thing !!!

            int undoScope = Globals.ThisAddIn.Application.BeginUndoScope("ParseCommand");

            Visio.Application app = Globals.ThisAddIn.Application;

            Visio.Document doc = app.ActiveDocument;

            Visio.Page page = app.ActivePage;

            // These are the top level elements that can appear in a command.
            // A command can contain more than one.

            if (commandElement.Elements("Documents").Any())
            {
                ProcessCommand_Documents(commandElement.Element("Documents").Elements());
            }

            if (commandElement.Elements("Layers").Any())
            {
                ProcessCommand_Layers(page, commandElement.Element("Layers").Elements());
            }

            if (commandElement.Elements("Pages").Any())
            {
                ProcessCommand_Pages(doc, commandElement.Element("Pages").Elements());
            }

            if (commandElement.Elements("Shapes").Any())
            {
                ProcessCommand_Shapes(commandElement.Element("Shapes").Elements());
            }

            Globals.ThisAddIn.Application.EndUndoScope(undoScope, true);
        }

        #region Commands - Documents

        private void ProcessCommand_Documents(IEnumerable<XElement> documentsElement)
        {
            // <Documents>
            //    <Add />
            //    <ActiveDocument>
            //        <Layers>
            //            <DeleteAll /> - Not sure if can do this
            //            <Delete Name = "" />
            //            <Add Name = "Layer1" IsVisible = "true" IsPrint = "true" IsActive = "true" IsLock = "true" IsSnap = "true" IsGlue = "true" Color = "" />
            //         </Layers>                          
            //    </ActiveDocument>                          
            // </Documents>
            VNCVisioAddIn.Common.WriteToWatchWindow(string.Format("{0}()",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));

            Visio.Application app = Globals.ThisAddIn.Application;

            Visio.Document doc = app.ActiveDocument;

            foreach (XElement element in documentsElement)
            {
                VNCVisioAddIn.Common.WriteToWatchWindow(element.ToString());
                VNCVisioAddIn.Common.WriteToWatchWindow(element.Name.LocalName);

                switch (element.Name.LocalName)
                {
                    case "Add":
                        ProcessCommand_Document_Add(element);
                        break;

                    case "ActiveDocument":
                        ProcessCommand_ActiveDocument(element);
                        break;

                    case "Document":
                        ProcessCommand_Document(element);
                        break;

                    case "Layers":
                        foreach (Visio.Page page in doc.Pages)
                        {
                            ProcessCommand_Layers(page, element.Elements());                           
                        }

                        break;
                        
                    default:
                        VNCVisioAddIn.Common.WriteToWatchWindow(string.Format("Element >{0}< not supported", element.Name.LocalName));
                        break;
                }
            }
        }

        private void ProcessCommand_Document(XElement documentElement)
        {
            VNCVisioAddIn.Common.WriteToWatchWindow(string.Format("{0}()",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));

            Visio.Application app = Globals.ThisAddIn.Application;

            // TODO(crhodes):
            // Add some error handling here.

            Visio.Document doc = app.Documents[documentElement.Attribute("Name").Value];

            if (documentElement.Elements("Layers").Any())
            {
                foreach (Visio.Page page in doc.Pages)
                {
                    ProcessCommand_Layers(page, documentElement.Element("Layers").Elements());                    
                }
            }

            if (documentElement.Elements("ShapeSheet").Any())
            {
                ProcessCommand_ShapeSheet(doc, documentElement.Element("ShapeSheet"));
            }
        }

        private void ProcessCommand_ActiveDocument(XElement activeDocumentElement)
        {
            VNCVisioAddIn.Common.WriteToWatchWindow(string.Format("{0}()",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));

            Visio.Application app = Globals.ThisAddIn.Application;

            Visio.Document doc = app.ActiveDocument;

            if (activeDocumentElement.Elements("Layers").Any())
            {
                foreach (Visio.Page page in doc.Pages)
                {
                    ProcessCommand_Layers(page, activeDocumentElement.Element("Layers").Elements());                  
                }
            }

            if (activeDocumentElement.Elements("ShapeSheet").Any())
            {
                ProcessCommand_ShapeSheet(doc, activeDocumentElement.Element("ShapeSheet"));
            }
        }

        private void ProcessCommand_Document_Add(XElement documentElement)
        {
            VNCVisioAddIn.Common.WriteToWatchWindow(string.Format("{0}()",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));

            try
            {
                Visio.Application app = Globals.ThisAddIn.Application;

                string documentName = documentElement.Attribute("Name").Value;

                app.Documents.Add(documentName);
            }
            catch (Exception ex)
            {
                VNCVisioAddIn.Common.WriteToWatchWindow(ex.ToString());
            }
        }
        #endregion

        #region Commands - Pages

        private void ProcessCommand_Pages(Visio.Document doc, IEnumerable<XElement> pagesElement)
        {
            VNCVisioAddIn.Common.WriteToWatchWindow(string.Format("{0}()",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));

            foreach (XElement element in pagesElement)
            {
                VNCVisioAddIn.Common.WriteToWatchWindow(element.ToString());
                VNCVisioAddIn.Common.WriteToWatchWindow(element.Name.LocalName);

                switch (element.Name.LocalName)
                {
                    case "Add":
                        ProcessCommand_Page_Add(doc, element);
                        break;

                    case "Delete":
                        ProcessCommand_Page_Delete(doc, element);
                        break;

                    case "DeleteAll":
                        ProcessCommand_Page_DeleteAll(element);
                        break;

                    case "Page":
                        ProcessCommand_Page(doc, element);
                        break;

                    case "Layers":
                        foreach (Visio.Page page in doc.Pages)
                        {
                            ProcessCommand_Layers(page, element.Elements());
                        }

                        break;

                    default:
                        VNCVisioAddIn.Common.WriteToWatchWindow(string.Format("Element >{0}< not supported", element.Name.LocalName));
                        break;
                }
            }
        }

        private void ProcessCommand_Page(Visio.Document doc, XElement pageElement)
        {
            VNCVisioAddIn.Common.WriteToWatchWindow(string.Format("{0}()",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));

            try
            {
                string pageName = pageElement.Attribute("Name") != null ? pageElement.Attribute("Name").Value : "";
                string backgroundPageName = pageElement.Attribute("BackgroundPageName") != null ? pageElement.Attribute("BackgroundPageName").Value : "";
                string isBackground = pageElement.Attribute("IsBackground") != null ? pageElement.Attribute("IsBackground").Value : "0";

                if ("" != backgroundPageName)
                {
                    doc.Pages[pageName].BackPage = backgroundPageName;
                }

                doc.Pages[pageName].Background = short.Parse(isBackground);
            }
            catch (Exception ex)
            {
                VNCVisioAddIn.Common.WriteToWatchWindow(ex.ToString());
            }
        }

        private void ProcessCommand_Page_Add(Visio.Document document, XElement pageElement)
        {
            VNCVisioAddIn.Common.WriteToWatchWindow(string.Format("{0}()",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));

            try
            {
                string newPageName = pageElement.Attribute("Name") != null ? pageElement.Attribute("Name").Value : "";
                string backgroundPageName = pageElement.Attribute("BackgroundPageName") != null ? pageElement.Attribute("BackgroundPageName").Value : "";
                string isBackground = pageElement.Attribute("IsBackground") != null ? pageElement.Attribute("IsBackground").Value : "0";

                // TODO(crhodes):
                // Need to pass in doc so this can work for any document

                Visio.Application app = Globals.ThisAddIn.Application;

                Visio.Document doc = app.ActiveDocument;

                Visio.Page newPage = Actions.Visio_Page.CreatePage(newPageName, backgroundPageName, short.Parse(isBackground));
            }
            catch (Exception ex)
            {
                VNCVisioAddIn.Common.WriteToWatchWindow(ex.ToString());
            }
        }

        private void ProcessCommand_Page_DeleteAll(XElement deleteAllElement)
        {
            VNCVisioAddIn.Common.WriteToWatchWindow(string.Format("{0}()",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));

        }

        private void ProcessCommand_Page_Delete(Visio.Document doc, XElement pageElement)
        {
            VNCVisioAddIn.Common.WriteToWatchWindow(string.Format("{0}()",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));

            try
            {
                string pageName = pageElement.Attribute("Name") != null ? pageElement.Attribute("Name").Value : "";

                if (pageName != "")
                {
                    Visio.Page deletePage = doc.Pages[pageName];

                    if (deletePage != null)
                    {
                        deletePage.Delete(0);
                    }
                    else
                    {
                        VNCVisioAddIn.Common.WriteToWatchWindow(string.Format("Page ({0}) not found", pageName));
                    }
                }
                else
                {
                    VNCVisioAddIn.Common.WriteToWatchWindow("Missing \"Name\" attribute");
                }

            }
            catch (Exception ex)
            {
                VNCVisioAddIn.Common.WriteToWatchWindow(ex.ToString());
            }
        }

        #endregion

        #region Commands - Layers

        private void ProcessCommand_Layers(Visio.Page page, IEnumerable<XElement> layersElement)
        {
            VNCVisioAddIn.Common.WriteToWatchWindow(string.Format("{0}()",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));

            // TODO(crhodes):
            // Need to pass in doc and page so this can work for any document

            foreach (XElement element in layersElement)
            {
                VNCVisioAddIn.Common.WriteToWatchWindow(element.ToString());
                VNCVisioAddIn.Common.WriteToWatchWindow(element.Name.LocalName);

                switch (element.Name.LocalName)
                {
                    case "Add":
                        ProcessCommand_Layer_Add(page, element);
                        break;

                    case "Delete":
                        ProcessCommand_Layer_Delete(page, element);
                        break;

                    case "DeleteAll":
                        ProcessCommand_Layer_DeleteAll(page, element);
                        break;

                    case "Layer":
                        ProcessCommand_Layer(page, element);
                        break;

                    default:
                        VNCVisioAddIn.Common.WriteToWatchWindow(string.Format("Element >{0}< not supported", element.Name.LocalName));
                        break;
                }
            }
        }

        private void ProcessCommand_Layer(Visio.Page page, XElement layerElement)
        {
            VNCVisioAddIn.Common.WriteToWatchWindow(string.Format("{0}()",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));

            try
            {
                string layerName = layerElement.Attribute("Name").Value;

                Visio.Layer targetLayer = page.Layers[layerName];

                if (targetLayer != null)
                {
                    string newName = (string)layerElement.Attribute("NewName");
                    string layerVisible = (string)layerElement.Attribute("IsVisible");
                    string layerPrint = (string)layerElement.Attribute("IsPrint");
                    string layerActive = (string)layerElement.Attribute("IsActive");
                    string layerLock = (string)layerElement.Attribute("IsLock");
                    string layerSnap = (string)layerElement.Attribute("IsSnap");
                    string layerGlue = (string)layerElement.Attribute("IsGlue");

                    if (newName != null && newName != "")
                    {
                        targetLayer.Name = newName;
                    }

                    if (layerVisible != null && layerVisible != "")
                    {
                        targetLayer.CellsC[(short)Visio.VisCellIndices.visLayerVisible].FormulaU = layerVisible;
                    }

                    if (layerPrint != null && layerPrint != "")
                    {
                        targetLayer.CellsC[(short)Visio.VisCellIndices.visLayerPrint].FormulaU = layerPrint;
                    }

                    if (layerActive != null && layerActive != "")
                    {
                        targetLayer.CellsC[(short)Visio.VisCellIndices.visLayerActive].FormulaU = layerActive;
                    }

                    if (layerLock != null && layerLock != "")
                    {
                        targetLayer.CellsC[(short)Visio.VisCellIndices.visLayerLock].FormulaU = layerLock;
                    }

                    if (layerSnap != null && layerSnap != "")
                    {
                        targetLayer.CellsC[(short)Visio.VisCellIndices.visLayerSnap].FormulaU = layerSnap;
                    }

                    if (layerGlue != null && layerGlue != "")
                    {
                        targetLayer.CellsC[(short)Visio.VisCellIndices.visLayerGlue].FormulaU = layerGlue;
                    }
                }
            }
            catch (Exception ex)
            {
                VNCVisioAddIn.Common.WriteToWatchWindow(ex.ToString());
            }
        }

        private void ProcessCommand_Layer_Add(Visio.Page page, XElement addElement)
        {
            VNCVisioAddIn.Common.WriteToWatchWindow(string.Format("{0}()",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));

            try
            {
                string layerName = addElement.Attribute("Name").Value;

                string layerVisible = addElement.Attribute("IsVisible") != null ? addElement.Attribute("IsVisible").Value : "1";
                string layerPrint = addElement.Attribute("IsPrint") != null ? addElement.Attribute("IsPrint").Value : "1";
                string layerActive = addElement.Attribute("IsActive") != null ? addElement.Attribute("IsActive").Value : "0";
                string layerLock = addElement.Attribute("IsLock") != null ? addElement.Attribute("IsLock").Value : "0";
                string layerSnap = addElement.Attribute("IsSnap") != null ? addElement.Attribute("IsSnap").Value : "1";
                string layerGlue = addElement.Attribute("IsGlue") != null ? addElement.Attribute("IsGlue").Value : "1";

                Actions.Visio_Page.AddLayer(page, layerName, layerVisible, layerPrint, layerActive, layerLock, layerSnap, layerGlue);

            }
            catch (Exception ex)
            {
                VNCVisioAddIn.Common.WriteToWatchWindow(ex.ToString());
            }
        }

        private void ProcessCommand_Layer_Delete(Visio.Page page, XElement deleteElement)
        {
            VNCVisioAddIn.Common.WriteToWatchWindow(string.Format("{0}()",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));

            try
            {
                string layerName = deleteElement.Attribute("Name").Value;
                string deleteShapes = deleteElement.Attribute("DeleteShapes").Value;

                Actions.Visio_Page.DeleteLayer(page, layerName, short.Parse(deleteShapes));
            }
            catch (Exception ex)
            {
                VNCVisioAddIn.Common.WriteToWatchWindow(ex.ToString());
            }
        }

        private void ProcessCommand_Layer_DeleteAll(Visio.Page page, XElement deleteAllElement)
        {
            VNCVisioAddIn.Common.WriteToWatchWindow(string.Format("{0}()",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));

            try
            {
                foreach (Visio.Layer layer in page.Layers)
                {
                    layer.Delete(0);
                }
            }
            catch (Exception ex)
            {
                VNCVisioAddIn.Common.WriteToWatchWindow(ex.ToString());
            }
        }

        #endregion
      
        #region Commands - Shapes

        /// <summary>
        /// For each Shape in Selection, apply Commands
        /// </summary>
        /// <param name="shapesElement"></param>
        private void ProcessCommand_Shapes(IEnumerable<XElement> shapesElement)
        {
            VNCVisioAddIn.Common.WriteToWatchWindow(string.Format("{0}()",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));

            Visio.Application app = Globals.ThisAddIn.Application;

            Visio.Selection selection = app.ActiveWindow.Selection;

            foreach (XElement element in shapesElement)
            {
                VNCVisioAddIn.Common.WriteToWatchWindow(element.ToString());
                VNCVisioAddIn.Common.WriteToWatchWindow(element.Name.LocalName);

                switch (element.Name.LocalName)
                {
                    case "Shape":
                        foreach (Visio.Shape shape in selection)
                        {
                            ProcessCommand_Shape(shape, element);
                        }

                        break;

                    default:
                        VNCVisioAddIn.Common.WriteToWatchWindow(string.Format("Element >{0}< not supported", element.Name.LocalName));
                        break;
                }
            }
        }

        private void ProcessCommand_Shape(Visio.Shape shape, XElement shapeElement)
        {
            foreach (XElement element in shapeElement.Elements())
            {
                VNCVisioAddIn.Common.WriteToWatchWindow(element.ToString());
                VNCVisioAddIn.Common.WriteToWatchWindow(element.Name.LocalName);

                switch (element.Name.LocalName)
                {
                    case "ShapeSheet":
                        ProcessCommand_ShapeSheet(shape, element);
                        break;

                    default:
                        VNCVisioAddIn.Common.WriteToWatchWindow(string.Format("Element >{0}< not supported", element.Name.LocalName));
                        break;
                }
            }
        }

        private void ProcessCommand_ShapeSheet(Visio.Document document, XElement shapeSheetElement)
        {
            VNCVisioAddIn.Common.WriteToWatchWindow(string.Format("{0}()",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));

            foreach (XElement element in shapeSheetElement.Elements())
            {
                VNCVisioAddIn.Common.WriteToWatchWindow(element.ToString());
                VNCVisioAddIn.Common.WriteToWatchWindow(element.Name.LocalName);

                switch (element.Name.LocalName)
                {
                    //case "AddPropRow":
                    //    ProcessCommand_ShapeSheet_AddPropRow(shape, element);
                    //    break;

                    //case "AddUserRow":
                    //    ProcessCommand_ShapeSheet_AddUserRow(shape, element);
                    //    break;

                    default:
                        VNCVisioAddIn.Common.WriteToWatchWindow(string.Format("Element >{0}< not supported", element.Name.LocalName));
                        break;
                }
            }
        }

        private void ProcessCommand_ShapeSheet(Visio.Page page, XElement shapeSheetElement)
        {
            VNCVisioAddIn.Common.WriteToWatchWindow(string.Format("{0}()",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));

            foreach (XElement element in shapeSheetElement.Elements())
            {
                VNCVisioAddIn.Common.WriteToWatchWindow(element.ToString());
                VNCVisioAddIn.Common.WriteToWatchWindow(element.Name.LocalName);

                switch (element.Name.LocalName)
                {
                    //case "AddPropRow":
                    //    ProcessCommand_ShapeSheet_AddPropRow(shape, element);
                    //    break;

                    //case "AddUserRow":
                    //    ProcessCommand_ShapeSheet_AddUserRow(shape, element);
                    //    break;

                    default:
                        VNCVisioAddIn.Common.WriteToWatchWindow(string.Format("Element >{0}< not supported", element.Name.LocalName));
                        break;
                }
            }
        }


        private void ProcessCommand_ShapeSheet(Visio.Shape shape, XElement shapeSheetElement)
        {
            VNCVisioAddIn.Common.WriteToWatchWindow(string.Format("{0}()",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));

            foreach (XElement element in shapeSheetElement.Elements())
            {
                VNCVisioAddIn.Common.WriteToWatchWindow(element.ToString());
                VNCVisioAddIn.Common.WriteToWatchWindow(element.Name.LocalName);

                switch (element.Name.LocalName)
                {
                    case "AddActionsRow":
                        ProcessCommand_ShapeSheet_ActionsRow(shape, element);
                        break;

                    case "AddControlsRow":
                        ProcessCommand_ShapeSheet_AddControlsRow(shape, element);
                        break;

                    case "AddHyperlinksRow":
                        ProcessCommand_ShapeSheet_AddHyperlinksRow(shape, element);
                        break;

                    case "AddPropRow":
                        ProcessCommand_ShapeSheet_AddPropRow(shape, element);
                        break;

                    case "AddUserRow":
                        ProcessCommand_ShapeSheet_AddUserRow(shape, element);
                        break;

                    case "SetFillFormat":
                        ProcessCommand_ShapeSheet_SetFillFormat(shape, element);
                        break;

                    case "SetShapeTransform":
                        ProcessCommand_ShapeSheet_ShapeTransform(shape, element);
                        break;

                    case "SetTextBlockFormat":
                        ProcessCommand_ShapeSheet_SetTextBlockFormat(shape, element);
                        break;

                    case "SetTextTransform":
                        ProcessCommand_ShapeSheet_SetTextTransform(shape, element);
                        break;

                    default:
                        VNCVisioAddIn.Common.WriteToWatchWindow(string.Format("Element >{0}< not supported", element.Name.LocalName));
                        break;
                }
            }
        }

        #region ShapeSheet Commands

        private void ProcessCommand_ShapeSheet_ActionsRow(Visio.Shape shape, XElement addActionsElement)
        {
            VNCVisioAddIn.Common.WriteToWatchWindow(string.Format("{0}()",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));

            try
            {
                // Required attributes

                string rowName = addActionsElement.Attribute("Name").Value;
                string action = addActionsElement.Attribute("Action").Value;
                string menu = addActionsElement.Attribute("Menu").Value;
                string tagName = addActionsElement.Attribute("TagName").Value;
                string buttonFace = addActionsElement.Attribute("ButtonFace").Value;
                string sortKey = addActionsElement.Attribute("SortKey").Value;
                string isChecked = addActionsElement.Attribute("Checked").Value;
                string disabled = addActionsElement.Attribute("Disabled").Value;
                string readOnly = addActionsElement.Attribute("ReadOnly").Value;
                string invisible = addActionsElement.Attribute("Invisible").Value;
                string beginGroup = addActionsElement.Attribute("BeginGroup").Value;
                string flyoutChild = addActionsElement.Attribute("FlyoutChild").Value;

                Actions.Visio_Shape.Populate_Actions_Section(shape, rowName, action, menu, tagName, buttonFace, sortKey, isChecked, disabled, readOnly, invisible, beginGroup, flyoutChild);
            }
            catch (Exception ex)
            {
                VNCVisioAddIn.Common.WriteToWatchWindow(ex.ToString());
            }
        }

        private void ProcessCommand_ShapeSheet_AddHyperlinksRow(Visio.Shape shape, XElement addHyperlinksElement)
        {
            VNCVisioAddIn.Common.WriteToWatchWindow(string.Format("{0}()",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));

            try
            {
                // Required attributes
                string rowName = string.Empty;
                string address = string.Empty;
                string subAddress = string.Empty;

                if (addHyperlinksElement.Attribute("Name") != null ) rowName = addHyperlinksElement.Attribute("Name").Value;
                if (addHyperlinksElement.Attribute("Address") != null) address = addHyperlinksElement.Attribute("Address").Value;
                if (addHyperlinksElement.Attribute("SubAddress") != null) subAddress = addHyperlinksElement.Attribute("SubAddress").Value;

                // TODO(crhodes):
                // Maybe some logic if neither Address nor SubAddress set

                // Optional fields

                string description = string.Empty;
                string extraInfo = string.Empty;
                string frame = string.Empty;
                string sortKey = string.Empty;
                string newWindow = "0";
                string default1 = "0";
                string invisible = "0";

                if (addHyperlinksElement.Attribute("Description") != null) extraInfo = addHyperlinksElement.Attribute("Description").Value;
                if (addHyperlinksElement.Attribute("ExtraInfo") != null) extraInfo = addHyperlinksElement.Attribute("ExtraInfo").Value;
                if (addHyperlinksElement.Attribute("Frame") != null) frame = addHyperlinksElement.Attribute("Frame").Value;
                if (addHyperlinksElement.Attribute("SortKey") != null) sortKey = addHyperlinksElement.Attribute("SortKey").Value;
                if (addHyperlinksElement.Attribute("NewWindow") != null) newWindow = addHyperlinksElement.Attribute("NewWindow").Value;
                if (addHyperlinksElement.Attribute("Default") != null) default1 = addHyperlinksElement.Attribute("Default").Value;
                if (addHyperlinksElement.Attribute("Invisible") != null) invisible = addHyperlinksElement.Attribute("Invisible").Value;


                Actions.Visio_Shape.Populate_Hyperlinks_Section(shape, rowName, description, address, subAddress, extraInfo, frame, sortKey, newWindow, default1, invisible);
            }
            catch (Exception ex)
            {
                VNCVisioAddIn.Common.WriteToWatchWindow(ex.ToString());
            }

        }

        private void ProcessCommand_ShapeSheet_AddControlsRow(Visio.Shape shape, XElement addControlsRowElement)
        {
            VNCVisioAddIn.Common.WriteToWatchWindow(string.Format("{0}()",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));

            try
            {
                // Required attributesin

                string rowName = addControlsRowElement.Attribute("Name").Value;
                string X = addControlsRowElement.Attribute("X").Value;
                string Y = addControlsRowElement.Attribute("Y").Value;
                string XDynamics = addControlsRowElement.Attribute("XDynamics").Value;
                string YDynamics = addControlsRowElement.Attribute("YDynamics").Value;
                string XBehavior = addControlsRowElement.Attribute("XBehavior").Value;
                string YBehavior = addControlsRowElement.Attribute("YBehavior").Value;
                string canGlue = addControlsRowElement.Attribute("CanGlue").Value;
                string tip = addControlsRowElement.Attribute("Tip").Value;

                Actions.Visio_Shape.Populate_Controls_Section(shape, rowName, X, Y, XDynamics, YDynamics, XBehavior, YBehavior, canGlue, tip);
            }
            catch (Exception ex)
            {
                VNCVisioAddIn.Common.WriteToWatchWindow(ex.ToString());
            }
        }

        private void ProcessCommand_ShapeSheet_SetTextBlockFormat(Visio.Shape shape, XElement setTextBlockFormatElement)
        {
            VNCVisioAddIn.Common.WriteToWatchWindow(string.Format("{0}()",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));

            try
            {
                VNCVisioAddIn.Domain.TextBlockFormat textBlockFormat = new VNCVisioAddIn.Domain.TextBlockFormat();

                textBlockFormat.LeftMargin = setTextBlockFormatElement.Attribute("LeftMargin").Value;
                textBlockFormat.TopMargin = setTextBlockFormatElement.Attribute("TopMargin").Value;
                textBlockFormat.RightMargin = setTextBlockFormatElement.Attribute("RightMargin").Value;
                textBlockFormat.BottomMargin = setTextBlockFormatElement.Attribute("BottomMargin").Value;
                textBlockFormat.TextDirection = setTextBlockFormatElement.Attribute("TextDirection").Value;
                textBlockFormat.VerticalAlign = setTextBlockFormatElement.Attribute("VerticalAlign").Value;
                textBlockFormat.TextBkgnd = setTextBlockFormatElement.Attribute("TextBkgnd").Value;
                textBlockFormat.TextBkgndTrans = setTextBlockFormatElement.Attribute("TextBkgndTrans").Value;
                textBlockFormat.DefaultTabStop = setTextBlockFormatElement.Attribute("DefaultTabStop").Value;

                Actions.Visio_Shape.Set_TextBlockFormat_Section(shape, textBlockFormat);
            }
            catch (Exception ex)
            {
                VNCVisioAddIn.Common.WriteToWatchWindow(ex.ToString());
            }
        }

        private void ProcessCommand_ShapeSheet_SetTextTransform(Visio.Shape shape, XElement setTextTransformElement)
        {
            VNCVisioAddIn.Common.WriteToWatchWindow(string.Format("{0}()",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));

            try
            {
                string width = setTextTransformElement.Attribute("Width").Value;
                string height = setTextTransformElement.Attribute("Height").Value;
                string pinX = setTextTransformElement.Attribute("PinX").Value;
                string pinY = setTextTransformElement.Attribute("PinY").Value;
                string locPinX = setTextTransformElement.Attribute("LocPinX").Value;
                string locPinY = setTextTransformElement.Attribute("LocPinY").Value;
                string angle = setTextTransformElement.Attribute("Angle").Value;


                Actions.Visio_Shape.Set_TextTransform_Section(shape, width, height, pinX, pinY, locPinX, locPinY, angle);
            }
            catch (Exception ex)
            {
                VNCVisioAddIn.Common.WriteToWatchWindow(ex.ToString());
            }
        }

        private string GetAttributeValueOrNull(XAttribute attribute)
        {
            string value = null;

            if (attribute != null)
            {
                value = attribute.Value;
            }

            return value;
        }

        private void ProcessCommand_ShapeSheet_SetFillFormat(Visio.Shape shape, XElement setFillFormatElement)
        {
            VNCVisioAddIn.Common.WriteToWatchWindow(string.Format("{0}()",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));

            try
            {
                string fillForegnd = GetAttributeValueOrNull(setFillFormatElement.Attribute("FillForegnd"));

                string fillForegndTrans = GetAttributeValueOrNull(setFillFormatElement.Attribute("FillForegndTrans"));
                string fillBkgnd = GetAttributeValueOrNull(setFillFormatElement.Attribute("FillBkgnd"));
                string fillBkgndTrans = GetAttributeValueOrNull(setFillFormatElement.Attribute("FillBkgndTrans"));
                string fillPattern = GetAttributeValueOrNull(setFillFormatElement.Attribute("FillPattern"));

                string shdwForegnd = GetAttributeValueOrNull(setFillFormatElement.Attribute("ShdwForegnd"));
                string shdwForegndTrans = GetAttributeValueOrNull(setFillFormatElement.Attribute("ShdwForegndTrans"));
                string shdwPattern = GetAttributeValueOrNull(setFillFormatElement.Attribute("ShdwPattern"));
                string shapeShdwOffsetX = GetAttributeValueOrNull(setFillFormatElement.Attribute("ShapeShdwOffsetX"));
                string shapeShdwOffsetY = GetAttributeValueOrNull(setFillFormatElement.Attribute("ShapeShdwOffsetY"));

                string shapeShdwType = GetAttributeValueOrNull(setFillFormatElement.Attribute("ShapeShdwType"));
                string shapeShdwObliqueAngle = GetAttributeValueOrNull(setFillFormatElement.Attribute("ShapeShdwObliqueAngle"));
                string shapeShdwScaleFactor = GetAttributeValueOrNull(setFillFormatElement.Attribute("ShapeShdwScaleFactor"));

                string shapeShdwBlur = GetAttributeValueOrNull(setFillFormatElement.Attribute("ShapeShdwBlur"));
                string shapeShdwShow = GetAttributeValueOrNull(setFillFormatElement.Attribute("ShapeShdwShow"));

                //attribute = setFillFormatElement.Attribute("FillForegnd"); 
                //if (attribute != null) fillForegnd = setFillFormatElement.Attribute("FillForegnd").Value;


                //string fillForegndTrans = setFillFormatElement.Attribute("FillForegndTrans").Value;
                //string fillBkgnd = setFillFormatElement.Attribute("FillBkgnd").Value;
                //string fillBkgndTrans = setFillFormatElement.Attribute("FillBkgndTrans").Value;
                //string fillPattern = setFillFormatElement.Attribute("FillPattern").Value;

                //string shdwForegnd = setFillFormatElement.Attribute("ShdwForegnd").Value;
                //string shdwForegndTrans = setFillFormatElement.Attribute("ShdwForegndTrans").Value;
                //string shdwPattern = setFillFormatElement.Attribute("ShdwPattern").Value;
                //string shapeShdwOffsetX = setFillFormatElement.Attribute("ShapeShdwOffsetX").Value;
                //string shapeShdwOffsetY = setFillFormatElement.Attribute("ShapeShdwOffsetY").Value;

                //string shapeShdwType = setFillFormatElement.Attribute("ShapeShdwType").Value;
                //string shapeShdwObliqueAngle = setFillFormatElement.Attribute("ShapeShdwObliqueAngle").Value;
                //string shapeShdwScaleFactor = setFillFormatElement.Attribute("ShapeShdwScaleFactor").Value;

                //string shapeShdwBlur = setFillFormatElement.Attribute("ShapeShdwBlur").Value;
                //string shapeShdwShow = setFillFormatElement.Attribute("ShapeShdwShow").Value;


                Actions.Visio_Shape.Set_FillFormat_SectionOld(shape, 
                    fillForegnd, fillForegndTrans, fillBkgnd, fillBkgndTrans, fillPattern, 
                    shdwForegnd, shdwForegndTrans, shdwPattern, shapeShdwOffsetX, shapeShdwOffsetY, 
                    shapeShdwType, shapeShdwObliqueAngle, shapeShdwScaleFactor, shapeShdwBlur, shapeShdwShow);
            }
            catch (Exception ex)
            {
                VNCVisioAddIn.Common.WriteToWatchWindow(ex.ToString());
            }
        }

        private void ProcessCommand_ShapeSheet_ShapeTransform(Visio.Shape shape, XElement shapeTransformElement)
        {
            VNCVisioAddIn.Common.WriteToWatchWindow(string.Format("{0}()",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));

            try
            {
                string width = GetAttributeValueOrNull(shapeTransformElement.Attribute("Width"));
                string height = GetAttributeValueOrNull(shapeTransformElement.Attribute("Height"));
                string pinX = GetAttributeValueOrNull(shapeTransformElement.Attribute("PinX"));
                string pinY = GetAttributeValueOrNull(shapeTransformElement.Attribute("PinY"));
                string flipX = GetAttributeValueOrNull(shapeTransformElement.Attribute("FlipX"));
                string flipY = GetAttributeValueOrNull(shapeTransformElement.Attribute("FlipY"));
                string locPinX = GetAttributeValueOrNull(shapeTransformElement.Attribute("LocPinX"));
                string locPinY = GetAttributeValueOrNull(shapeTransformElement.Attribute("LocPinY"));
                string angle = GetAttributeValueOrNull(shapeTransformElement.Attribute("Angle"));
                string resizeMode = GetAttributeValueOrNull(shapeTransformElement.Attribute("Angle"));

                Actions.Visio_Shape.Set_ShapeTransform_Section(shape, width, height, pinX, pinY, flipX, flipY, locPinX, locPinY, angle, resizeMode);
            }
            catch (Exception ex)
            {
                VNCVisioAddIn.Common.WriteToWatchWindow(ex.ToString());
            }
        }

        private short GetVisPropType(string value)
        {
            short visPropType = 0;

            switch (value)
            {
                case "VisCellVals.visPropTypeBool":
                    visPropType = (short)Visio.VisCellVals.visPropTypeBool;
                    break;

                case "VisCellVals.visPropTypeCurrency":
                    visPropType = (short)Visio.VisCellVals.visPropTypeCurrency;
                    break;

                case "VisCellVals.visPropTypeDate":
                    visPropType = (short)Visio.VisCellVals.visPropTypeDate;
                    break;

                case "VisCellVals.visPropTypeDuration":
                    visPropType = (short)Visio.VisCellVals.visPropTypeDuration;
                    break;

                case "VisCellVals.visPropTypeListFix":
                    visPropType = (short)Visio.VisCellVals.visPropTypeListFix;
                    break;

                case "VisCellVals.visPropTypeListVar":
                    visPropType = (short)Visio.VisCellVals.visPropTypeListVar;
                    break;

                case "VisCellVals.visPropTypeNumber":
                    visPropType = (short)Visio.VisCellVals.visPropTypeNumber;
                    break;

                case "VisCellVals.visPropTypeString":
                    visPropType = (short)Visio.VisCellVals.visPropTypeString;
                    break;

                default:
                    VNCVisioAddIn.Common.WriteToWatchWindow(string.Format("Unrecognized VisPropType >{0}<", value));
                    
                    break;
            }

            return visPropType;
        }

        private void ProcessCommand_ShapeSheet_AddPropRow(Visio.Shape shape, XElement addPropRowElement)
        {
            VNCVisioAddIn.Common.WriteToWatchWindow(string.Format("{0}()",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));

            try
            {
                // Required attributes

                string row = addPropRowElement.Attribute("Row").Value;
                string label = addPropRowElement.Attribute("Label").Value;
                short type = GetVisPropType(addPropRowElement.Attribute("Type").Value);
                string format = null;   // Used for numeric values
                string value = null;

                // Sometimes we want a string and sometimes we may want to enter a formula.

                if (addPropRowElement.Attribute("Format") != null) format = addPropRowElement.Attribute("Format").Value;

                if (addPropRowElement.Attribute("Format") != null)
                {
                    format = addPropRowElement.Attribute("Format").Value;
                }
                else if (addPropRowElement.Attribute("FormatQuoted") != null)
                {
                    format = string.Format("\"{0}\"", addPropRowElement.Attribute("FormatQuoted").Value);
                }

                if (addPropRowElement.Attribute("Value") != null)
                {
                    value = addPropRowElement.Attribute("Value").Value;
                }
                else if (addPropRowElement.Attribute("ValueQuoted") != null)
                {
                    value = string.Format("\"{0}\"", addPropRowElement.Attribute("ValueQuoted").Value);
                }
                // Optional attributes

                string prompt = null;
                string sortKey = null;

                if (addPropRowElement.Attribute("Prompt") != null) prompt = addPropRowElement.Attribute("Prompt").Value;
                if (addPropRowElement.Attribute("SortKey") != null) sortKey = addPropRowElement.Attribute("SortKey").Value;

                // Rarely used attributes - Not supported for now
                // TODO(crhodes):
                // This is going to take some thinking on how to handle.  Probably do the same as above for Prompt and SortKey.

                XAttribute invisibleAttribute = addPropRowElement.Attribute("Invisible");
                XAttribute askAttribute = addPropRowElement.Attribute("Ask");
                XAttribute langIDtAttribute = addPropRowElement.Attribute("LangID");
                XAttribute calendarAttribute = addPropRowElement.Attribute("Calendar");

                if (prompt != null || sortKey != null)
                {
                    Actions.Visio_Shape.Add_Prop_Row(shape, row, label, type, format, value, prompt, sortKey);
                }
                else
                {
                    Actions.Visio_Shape.Add_Prop_Row(shape, row, label, type, format, value);
                }
            }
            catch (Exception ex)
            {
                VNCVisioAddIn.Common.WriteToWatchWindow(ex.ToString());
            }
        }

        private void ProcessCommand_ShapeSheet_AddUserRow(Visio.Shape shape, XElement addUserRowElement)
        {
            VNCVisioAddIn.Common.WriteToWatchWindow(string.Format("{0}()",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));

            try
            {
                string row = addUserRowElement.Attribute("Row").Value;
                string value = null;

                if (addUserRowElement.Attribute("Value") != null)
                {
                    value = addUserRowElement.Attribute("Value").Value;
                }
                else if (addUserRowElement.Attribute("ValueQuoted") != null)
                {
                    value = string.Format("\"{0}\"", addUserRowElement.Attribute("ValueQuoted").Value);
                }

                XAttribute promptAttribute = addUserRowElement.Attribute("Prompt");

                if (promptAttribute != null)
                {
                    Actions.Visio_Shape.Add_User_Row(shape, row, value, promptAttribute.Value);                    
                }
                else
                {
                    Actions.Visio_Shape.Add_User_Row(shape, row, value);                    
                }
            }
            catch (Exception ex)
            {
                VNCVisioAddIn.Common.WriteToWatchWindow(ex.ToString());
            }
        }        

        #endregion

        #endregion

        #endregion

        private void LoadExcelFile()
        {
            XL.Application xlApp = new XL.Application();

            XL.Workbook wb = xlApp.Workbooks.Open(@"B:\Publish\SupportTools_Visio\TestData.xlsx");
            XL.Worksheet ws = wb.Sheets[1];
            XL.Range rng = ws.UsedRange;

            int rows = rng.Rows.Count;
            int cols = rng.Columns.Count;

            for (int i = 1; i <= rows; i++)
            {
                for (int j = 1; j <= cols; j++)
                {
                    VNCVisioAddIn.Common.WriteToWatchWindow(rng.Cells[i, j].Value2.ToString());
                }
            }

            wb.Close();
        }

        private void LoadExcelTable()
        {
            string workBookName = @"B:\Publish\SupportTools_Visio\TestData.xlsx";
            string workSheetName = "Sheet2";
            string tableName = "tbl_Data";
            XL.Application xlApp = new XL.Application();

            XL.Workbook wb = xlApp.Workbooks.Open(workBookName);
            XL.Worksheet ws = wb.Sheets[workSheetName];
            XL.ListObject lo = ws.ListObjects[tableName];
            XL.ListColumns listColumns = lo.ListColumns;
            XL.ListRows listRows = lo.ListRows;

            VNCVisioAddIn.Common.WriteToWatchWindow(string.Format("{0}\n", tableName));

            foreach (XL.ListColumn col in listColumns)
            {
                VNCVisioAddIn.Common.WriteToWatchWindow(col.Name);
            }

            tableName = "tbl_Data2";

            lo = ws.ListObjects[tableName];
            listColumns = lo.ListColumns;
            listRows = lo.ListRows;

            VNCVisioAddIn.Common.WriteToWatchWindow(string.Format("{0}\n", tableName));

            foreach (XL.ListColumn col in listColumns)
            {
                VNCVisioAddIn.Common.WriteToWatchWindow(col.Name);
            }

            foreach (XL.ListRow row in listRows)
            {
                VNCVisioAddIn.Common.WriteToWatchWindow(row.ToString());
            }
            wb.Close();
        }

        private void UseExcelDataReader()
        {
            string path = @"B:\Publish\SupportTools_Visio\TestData.xlsx";

            var excelData = new ExcelHlp.Domain.XlData(path);
            var sheets = excelData.GetWorkSheetNames();

            foreach (var sheet in sheets)
            {
                VNCVisioAddIn.Common.WriteToWatchWindow(sheet);
            }

            VNCVisioAddIn.Common.WriteToWatchWindow("Has Header Row\n");

            var info = excelData.GetData("Sheet1");

            foreach (var row in info)
            {
                VNCVisioAddIn.Common.WriteToWatchWindow("NewRow\n");

                for (int i = 0; i <= row.ItemArray.GetUpperBound(0); i++)
                {
                    VNCVisioAddIn.Common.WriteToWatchWindow(row[i].ToString());
                }
            }

            VNCVisioAddIn.Common.WriteToWatchWindow("Has No Header Row\n");

            info = excelData.GetData("Sheet1", false);

            foreach (var row in info)
            {
                VNCVisioAddIn.Common.WriteToWatchWindow("NewRow\n");

                for (int i = 0; i <= row.ItemArray.GetUpperBound(0); i++)
                {
                    VNCVisioAddIn.Common.WriteToWatchWindow(row[i].ToString());
                }
            }

            List<TestData> testData = new List<TestData>();

            info = excelData.GetData("Sheet1");

            foreach (var row in info)
            {
                var testDataRow = new TestData()
                {
                    Col1 = row["Col1"].ToString(),
                    Col2 = row["Col2"].ToString(),
                    Col3 = row["Col3"].ToString(),
                    Col4 = row["Col4"].ToString(),
                    Col5 = row["Col5"].ToString()
                };

                testData.Add(testDataRow);
            }

            foreach (var item in testData)
            {
                VNCVisioAddIn.Common.WriteToWatchWindow(
                    string.Format("Col1:{0} Col2:{1} Col3:{2} Col4:{3} Col5:{4}", item.Col1, item.Col2, item.Col3, item.Col4, item.Col5)
                    );
            }
        }

        private void UseLinqToExcel()
        {
            string path = @"B:\Publish\SupportTools_Visio\TestData.xlsx";
            var excel = new LTE.ExcelQueryFactory(path);

            var stuff = from c in excel.Worksheet<TestData>()
                        select c;

            foreach (var item in stuff)
            {
                VNCVisioAddIn.Common.WriteToWatchWindow(
                    string.Format("Col1:{0} Col2:{1} Col3:{2} Col4:{3} Col5:{4}", item.Col1, item.Col2, item.Col3, item.Col4, item.Col5)
                    );
            }
        }

        #endregion


    }
}
