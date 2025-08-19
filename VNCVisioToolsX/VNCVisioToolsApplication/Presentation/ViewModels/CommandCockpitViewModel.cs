using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

using System.Windows.Input;
using System.Xml;
using System.Xml.Linq;

using Microsoft.Win32;

using Prism.Commands;
using Prism.Events;
using Prism.Services.Dialogs;

using VNC;
using VNC.Core.Mvvm;

using MSVisio = Microsoft.Office.Interop.Visio;
using VNCVisioAddIn = VNC.Visio.VSTOAddIn;

namespace VNCVisioToolsApplication.Presentation.ViewModels
{

    public class CommandCockpitViewModel : EventViewModelBase, ICommandCockpitViewModel, IInstanceCountVM
    {
        #region Constructors, Initialization, and Load

        public CommandCockpitViewModel(
            IEventAggregator eventAggregator,
            DialogService dialogService) : base(eventAggregator, dialogService)
        {
            Int64 startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            // TODO(crhodes)
            // Save constructor parameters here

            InitializeViewModel();

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void InitializeViewModel()
        {
            Int64 startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            InstanceCountVM++;

            // TODO(crhodes)
            //

            PopulateControlsFromXmlFile(Common.cCONFIG_FILE);

            SayHelloCommand = new DelegateCommand(
                SayHello, SayHelloCanExecute);

            ExecuteCommand = new DelegateCommand(
                Execute, ExecuteCanExecute);

            ReloadXmlCommand = new DelegateCommand(
                ReloadXml, ReloadXmlCanExecute);

            Message = "CommandCockpitViewModel says hello";

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void PopulateControlsFromXmlFile(string fileNameAndPath)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            VisioCommands = new ObservableCollection<XElement>();
            //SolutionFiles = new ObservableCollection<XElement>();
            //ProjectFiles = new ObservableCollection<XElement>();
            //SourceFiles = new ObservableCollection<XElement>();

            //FilesDemoCS = new ObservableCollection<string>();
            //FilesDemoVB = new ObservableCollection<string>();

            //SelectedSolutionFiles = new ObservableCollection<XElement>();
            //SelectedProjectFiles = new ObservableCollection<XElement>();
            //SelectedSourceFiles = new ObservableCollection<XElement>();

            try
            {
                XmlTextReader xtr = new XmlTextReader(fileNameAndPath);

                XDocument xDocument = XDocument.Load(xtr, LoadOptions.PreserveWhitespace);

                var commmandCockput = xDocument.Descendants("CommandCockpit");

                foreach (var command in commmandCockput.Elements("Command"))
                {
                    VisioCommands.Add(command);
                }
            }
            catch (Exception ex)
            {
                
            }

            //foreach (var file in xDocument.Descendants("DemoFiles")
            //    .Elements("File")
            //    .Where(df => df.Attribute("Language").Value == "CS"))
            //{
            //    FilesDemoCS.Add(file.Attribute("FullPath").Value);
            //}

            //foreach (var file in xDocument.Descendants("DemoFiles")
            //    .Elements("File")
            //    .Where(df => df.Attribute("Language").Value == "VB"))
            //{
            //    FilesDemoVB.Add(file.Attribute("FullPath").Value);
            //}

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);
        }
        #endregion

        #region Enums


        #endregion

        #region Structures


        #endregion

        #region Fields and Properties

        public ObservableCollection<XElement> VisioCommands { get; set; }

        private XElement _selectedCommand;

        public XElement SelectedCommand
        {
            get => _selectedCommand;
            set
            {
                if (_selectedCommand == value)
                    return;
                _selectedCommand = value;

                //SolutionFiles.Clear();
                ////ProjectFiles.Clear();
                ////SourceFiles.Clear();

                //SelectedSolutionFiles.Clear();
                //SelectedProjectFiles.Clear();
                //SelectedSourceFiles.Clear();

                //Repository = _selectedBranch.Attribute("Repository").Value;
                //Branch = _selectedBranch.Attribute("Name").Value;
                //RepositoryPath = $"{_selectedBranch.Attribute("RepositoryPath").Value}";

                //var solutions = _selectedBranch.Elements("Solution");

                //SolutionFiles.AddRange(solutions);

                OnPropertyChanged();
            }
        }

        public ICommand SayHelloCommand { get; private set; }

        private string _message;

        public string Message
        {
            get => _message;
            set
            {
                if (_message == value)
                    return;
                _message = value;
                OnPropertyChanged();
            }
        }

        #endregion

        #region Event Handlers



        #endregion

        #region Public Methods


        #endregion

        #region Protected Methods


        #endregion

        #region Private Methods

        #region Commands

        #region ReloadXml Command

        public DelegateCommand ReloadXmlCommand { get; set; }
        public string ReloadXmlContent { get; set; } = "...";
        public string ReloadXmlToolTip { get; set; } = "ReloadXml ToolTip";

        // Can get fancy and use Resources
        //public string ReloadXmlContent { get; set; } = "ViewName_ReloadXmlContent";
        //public string ReloadXmlToolTip { get; set; } = "ViewName_ReloadXmlContentToolTip";

        // Put these in Resource File
        //    <system:String x:Key="ViewName_ReloadXmlContent">ReloadXml</system:String>
        //    <system:String x:Key="ViewName_ReloadXmlContentToolTip">ReloadXml ToolTip</system:String>  

        public void ReloadXml()
        {
            // TODO(crhodes)
            // Do something amazing.
            Message = "Cool, you called ReloadXml";

            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*";
            openFileDialog1.FileName = "";

            if (true == openFileDialog1.ShowDialog())
            {
                string fileName = openFileDialog1.FileName;

                PopulateControlsFromXmlFile(fileName);
            }
            //Common.EventAggregator.GetEvent<ReloadXmlEvent>().Publish();

            // Start Cut Four - Put this in places that listen for event

            //Common.EventAggregator.GetEvent<ReloadXmlEvent>().Subscribe(ReloadXml);

            // End Cut Four

        }

        public bool ReloadXmlCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion

        #endregion

        #region Execute Command

        public DelegateCommand ExecuteCommand { get; set; }
        public string ExecuteContent { get; set; } = "Execute";
        public string ExecuteToolTip { get; set; } = "Execute ToolTip";

        // Can get fancy and use Resources
        //public string ExecuteContent { get; set; } = "ViewName_ExecuteContent";
        //public string ExecuteToolTip { get; set; } = "ViewName_ExecuteContentToolTip";

        // Put these in Resource File
        //    <system:String x:Key="ViewName_ExecuteContent">Execute</system:String>
        //    <system:String x:Key="ViewName_ExecuteContentToolTip">Execute ToolTip</system:String>  

        public void Execute()
        {
            // TODO(crhodes)
            // Do something amazing.
            Message = "Cool, you called Execute";
            //Common.EventAggregator.GetEvent<ExecuteEvent>().Publish();

            ParseCommand(SelectedCommand);

            // Start Cut Four

            // Put this in places that listen for event
            //Common.EventAggregator.GetEvent<ExecuteEvent>().Subscribe(Execute);

            // End Cut Four
        }

        public bool ExecuteCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion

        private void SayHello()
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Message = $"Hello from {this.GetType()}";

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private bool SayHelloCanExecute()
        {
            return true;
        }

        #endregion

        #region XML Commands

        private void ParseCommand(XElement commandElement)
        {
            // Wrap a big, OMG, what have I done ???, undo around the whole thing !!!

            int undoScope = Common.VisioApplication.BeginUndoScope("ParseCommand");

            MSVisio.Application app = Common.VisioApplication;

            MSVisio.Document doc = app.ActiveDocument;

            MSVisio.Page page = app.ActivePage;

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

            Common.VisioApplication.EndUndoScope(undoScope, true);
        }

        private void ProcessCommand_Documents(IEnumerable<XElement> documentsElement)
        {
            VNCVisioAddIn.Common.WriteToWatchWindow($"{System.Reflection.MethodInfo.GetCurrentMethod().Name}()");

            MSVisio.Application app = Common.VisioApplication;
            MSVisio.Document doc = app.ActiveDocument;

            foreach (XElement element in documentsElement)
            {
                VNCVisioAddIn.Common.WriteToWatchWindow($"{element}");
                VNCVisioAddIn.Common.WriteToWatchWindow($"{element.Name.LocalName}");

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
                        foreach (MSVisio.Page page in doc.Pages)
                        {
                            ProcessCommand_Layers(page, element.Elements());
                        }
                        break;

                    default:
                        VNCVisioAddIn.Common.WriteToWatchWindow($"Element >{element.Name.LocalName}< not supported");
                        break;
                }
            }
        }

        private void ProcessCommand_Document(XElement documentElement)
        {
            VNCVisioAddIn.Common.WriteToWatchWindow(string.Format("{0}()",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));

            MSVisio.Application app = Common.VisioApplication;

            // TODO(crhodes):
            // Add some error handling here.

            MSVisio.Document doc = app.Documents[documentElement.Attribute("Name").Value];

            if (documentElement.Elements("Layers").Any())
            {
                foreach (MSVisio.Page page in doc.Pages)
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

            MSVisio.Application app = Common.VisioApplication;

            MSVisio.Document doc = app.ActiveDocument;

            if (activeDocumentElement.Elements("Layers").Any())
            {
                foreach (MSVisio.Page page in doc.Pages)
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
                MSVisio.Application app = Common.VisioApplication;

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

        private void ProcessCommand_Pages(MSVisio.Document doc, IEnumerable<XElement> pagesElement)
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
                        foreach (MSVisio.Page page in doc.Pages)
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

        private void ProcessCommand_Page(MSVisio.Document doc, XElement pageElement)
        {
            VNCVisioAddIn.Common.WriteToWatchWindow(string.Format("{0}()",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));

            try
            {
                string pageName = pageElement.Attribute("Name") != null ? pageElement.Attribute("Name").Value : "";
                string backgroundPageName = pageElement.Attribute("BackgroundPageName") != null ? pageElement.Attribute("BackgroundPageName").Value : "";
                string isBackground = pageElement.Attribute("IsBackground") != null ? pageElement.Attribute("IsBackground").Value : "0";

                // TODO(crhodes)
                // Need more logic here to handle if page doesn't exist.

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

        private void ProcessCommand_Page_Add(MSVisio.Document document, XElement pageElement)
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

                MSVisio.Application app = Common.VisioApplication;

                MSVisio.Document doc = app.ActiveDocument;

                MSVisio.Page newPage = VNCVisioToolsApplication.Actions.Visio_Page.CreatePage(newPageName, backgroundPageName, short.Parse(isBackground));
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

        private void ProcessCommand_Page_Delete(MSVisio.Document doc, XElement pageElement)
        {
            VNCVisioAddIn.Common.WriteToWatchWindow(string.Format("{0}()",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));

            try
            {
                string pageName = pageElement.Attribute("Name") != null ? pageElement.Attribute("Name").Value : "";

                if (pageName != "")
                {
                    MSVisio.Page deletePage = doc.Pages[pageName];

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

        private void ProcessCommand_Layers(MSVisio.Page page, IEnumerable<XElement> layersElement)
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

        private void ProcessCommand_Layer(MSVisio.Page page, XElement layerElement)
        {
            VNCVisioAddIn.Common.WriteToWatchWindow(string.Format("{0}()",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));

            try
            {
                string layerName = layerElement.Attribute("Name").Value;

                MSVisio.Layer targetLayer = page.Layers[layerName];

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
                        targetLayer.CellsC[(short)MSVisio.VisCellIndices.visLayerVisible].FormulaU = layerVisible;
                    }

                    if (layerPrint != null && layerPrint != "")
                    {
                        targetLayer.CellsC[(short)MSVisio.VisCellIndices.visLayerPrint].FormulaU = layerPrint;
                    }

                    if (layerActive != null && layerActive != "")
                    {
                        targetLayer.CellsC[(short)MSVisio.VisCellIndices.visLayerActive].FormulaU = layerActive;
                    }

                    if (layerLock != null && layerLock != "")
                    {
                        targetLayer.CellsC[(short)MSVisio.VisCellIndices.visLayerLock].FormulaU = layerLock;
                    }

                    if (layerSnap != null && layerSnap != "")
                    {
                        targetLayer.CellsC[(short)MSVisio.VisCellIndices.visLayerSnap].FormulaU = layerSnap;
                    }

                    if (layerGlue != null && layerGlue != "")
                    {
                        targetLayer.CellsC[(short)MSVisio.VisCellIndices.visLayerGlue].FormulaU = layerGlue;
                    }
                }
            }
            catch (Exception ex)
            {
                VNCVisioAddIn.Common.WriteToWatchWindow(ex.ToString());
            }
        }

        private void ProcessCommand_Layer_Add(MSVisio.Page page, XElement addElement)
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

                VNCVisioToolsApplication.Actions.Visio_Page.AddLayer(page, layerName, layerVisible, layerPrint, layerActive, layerLock, layerSnap, layerGlue);

            }
            catch (Exception ex)
            {
                VNCVisioAddIn.Common.WriteToWatchWindow(ex.ToString());
            }
        }

        private void ProcessCommand_Layer_Delete(MSVisio.Page page, XElement deleteElement)
        {
            VNCVisioAddIn.Common.WriteToWatchWindow(string.Format("{0}()",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));

            try
            {
                string layerName = deleteElement.Attribute("Name").Value;
                string deleteShapes = deleteElement.Attribute("DeleteShapes").Value;

                VNCVisioToolsApplication.Actions.Visio_Page.DeleteLayer(page, layerName, short.Parse(deleteShapes));
            }
            catch (Exception ex)
            {
                VNCVisioAddIn.Common.WriteToWatchWindow(ex.ToString());
            }
        }

        private void ProcessCommand_Layer_DeleteAll(MSVisio.Page page, XElement deleteAllElement)
        {
            VNCVisioAddIn.Common.WriteToWatchWindow(string.Format("{0}()",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));

            try
            {
                foreach (MSVisio.Layer layer in page.Layers)
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

            MSVisio.Application app = Common.VisioApplication;

            MSVisio.Selection selection = app.ActiveWindow.Selection;

            foreach (XElement element in shapesElement)
            {
                VNCVisioAddIn.Common.WriteToWatchWindow(element.ToString());
                VNCVisioAddIn.Common.WriteToWatchWindow(element.Name.LocalName);

                switch (element.Name.LocalName)
                {
                    case "Shape":
                        foreach (MSVisio.Shape shape in selection)
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

        private void ProcessCommand_Shape(MSVisio.Shape shape, XElement shapeElement)
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

        private void ProcessCommand_ShapeSheet(MSVisio.Document document, XElement shapeSheetElement)
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
                        Common.WriteToWatchWindow(string.Format("Element >{0}< not supported", element.Name.LocalName));
                        break;
                }
            }
        }

        private void ProcessCommand_ShapeSheet(MSVisio.Page page, XElement shapeSheetElement)
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


        private void ProcessCommand_ShapeSheet(MSVisio.Shape shape, XElement shapeSheetElement)
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

        private void ProcessCommand_ShapeSheet_ActionsRow(MSVisio.Shape shape, XElement addActionsElement)
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

                VNCVisioAddIn.Helpers.Populate_Actions_Section(shape, rowName, action, menu, tagName, buttonFace, sortKey, isChecked, disabled, readOnly, invisible, beginGroup, flyoutChild);
            }
            catch (Exception ex)
            {
                VNCVisioAddIn.Common.WriteToWatchWindow(ex.ToString());
            }
        }

        private void ProcessCommand_ShapeSheet_AddHyperlinksRow(MSVisio.Shape shape, XElement addHyperlinksElement)
        {
            VNCVisioAddIn.Common.WriteToWatchWindow(string.Format("{0}()",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));

            try
            {
                // Required attributes
                string rowName = string.Empty;
                string address = string.Empty;
                string subAddress = string.Empty;

                if (addHyperlinksElement.Attribute("Name") != null) rowName = addHyperlinksElement.Attribute("Name").Value;
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


                VNCVisioAddIn.Helpers.Populate_Hyperlinks_Section(shape, rowName, description, address, subAddress, extraInfo, frame, sortKey, newWindow, default1, invisible);
            }
            catch (Exception ex)
            {
                VNCVisioAddIn.Common.WriteToWatchWindow(ex.ToString());
            }

        }

        private void ProcessCommand_ShapeSheet_AddControlsRow(MSVisio.Shape shape, XElement addControlsRowElement)
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

                VNCVisioAddIn.Helpers.Populate_Controls_Section(shape, rowName, X, Y, XDynamics, YDynamics, XBehavior, YBehavior, canGlue, tip);
            }
            catch (Exception ex)
            {
                VNCVisioAddIn.Common.WriteToWatchWindow(ex.ToString());
            }
        }

        private void ProcessCommand_ShapeSheet_SetTextBlockFormat(MSVisio.Shape shape, XElement setTextBlockFormatElement)
        {
            VNCVisioAddIn.Common.WriteToWatchWindow(string.Format("{0}()",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));

            try
            {
                VNCVisioAddIn.Domain.TextBlockFormatRow textBlockFormat = new VNCVisioAddIn.Domain.TextBlockFormatRow();

                textBlockFormat.LeftMargin = setTextBlockFormatElement.Attribute("LeftMargin").Value;
                textBlockFormat.TopMargin = setTextBlockFormatElement.Attribute("TopMargin").Value;
                textBlockFormat.RightMargin = setTextBlockFormatElement.Attribute("RightMargin").Value;
                textBlockFormat.BottomMargin = setTextBlockFormatElement.Attribute("BottomMargin").Value;
                textBlockFormat.TextDirection = setTextBlockFormatElement.Attribute("TextDirection").Value;
                textBlockFormat.VerticalAlign = setTextBlockFormatElement.Attribute("VerticalAlign").Value;
                textBlockFormat.TextBkgnd = setTextBlockFormatElement.Attribute("TextBkgnd").Value;
                textBlockFormat.TextBkgndTrans = setTextBlockFormatElement.Attribute("TextBkgndTrans").Value;
                textBlockFormat.DefaultTabStop = setTextBlockFormatElement.Attribute("DefaultTabStop").Value;

                VNCVisioAddIn.Domain.TextBlockFormatRow.SetRow(shape, textBlockFormat);
            }
            catch (Exception ex)
            {
                VNCVisioAddIn.Common.WriteToWatchWindow(ex.ToString());
            }
        }

        private void ProcessCommand_ShapeSheet_SetTextTransform(MSVisio.Shape shape, XElement setTextTransformElement)
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


                VNCVisioAddIn.Helpers.Set_TextXForm_Section(shape, width, height, pinX, pinY, locPinX, locPinY, angle);
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

        private void ProcessCommand_ShapeSheet_SetFillFormat(MSVisio.Shape shape, XElement setFillFormatElement)
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


                VNCVisioAddIn.Helpers.Set_FillFormat_SectionOld(shape,
                    fillForegnd, fillForegndTrans, fillBkgnd, fillBkgndTrans, fillPattern,
                    shdwForegnd, shdwForegndTrans, shdwPattern, shapeShdwOffsetX, shapeShdwOffsetY,
                    shapeShdwType, shapeShdwObliqueAngle, shapeShdwScaleFactor, shapeShdwBlur, shapeShdwShow);
            }
            catch (Exception ex)
            {
                VNCVisioAddIn.Common.WriteToWatchWindow(ex.ToString());
            }
        }

        private void ProcessCommand_ShapeSheet_ShapeTransform(MSVisio.Shape shape, XElement shapeTransformElement)
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

                VNCVisioAddIn.Helpers.Set_ShapeTransform_Section(shape, width, height, pinX, pinY, flipX, flipY, locPinX, locPinY, angle, resizeMode);
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
                case "3":
                    visPropType = (short)MSVisio.VisCellVals.visPropTypeBool;
                    break;

                case "VisCellVals.visPropTypeCurrency":
                case "7":
                    visPropType = (short)MSVisio.VisCellVals.visPropTypeCurrency;
                    break;

                case "VisCellVals.visPropTypeDate":
                case "5":
                    visPropType = (short)MSVisio.VisCellVals.visPropTypeDate;
                    break;

                case "VisCellVals.visPropTypeDuration":
                case "6":
                    visPropType = (short)MSVisio.VisCellVals.visPropTypeDuration;
                    break;

                case "VisCellVals.visPropTypeListFix":
                case "1":
                    visPropType = (short)MSVisio.VisCellVals.visPropTypeListFix;
                    break;

                case "VisCellVals.visPropTypeListVar":
                case "4":
                    visPropType = (short)MSVisio.VisCellVals.visPropTypeListVar;
                    break;

                case "VisCellVals.visPropTypeNumber":
                case "2":
                    visPropType = (short)MSVisio.VisCellVals.visPropTypeNumber;
                    break;

                case "VisCellVals.visPropTypeString":
                case "0":
                    visPropType = (short)MSVisio.VisCellVals.visPropTypeString;
                    break;

                default:
                    VNCVisioAddIn.Common.WriteToWatchWindow(string.Format("Unrecognized VisPropType >{0}<", value));

                    break;
            }

            return visPropType;
        }

        private void ProcessCommand_ShapeSheet_AddPropRow(MSVisio.Shape shape, XElement addPropRowElement)
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
                    VNCVisioAddIn.Helpers.Add_Prop_Row(shape, row, label, type, format, value, prompt, sortKey);
                }
                else
                {
                    VNCVisioAddIn.Helpers.Add_Prop_Row(shape, row, label, type, format, value);
                }
            }
            catch (Exception ex)
            {
                VNCVisioAddIn.Common.WriteToWatchWindow(ex.ToString());
            }
        }

        private void ProcessCommand_ShapeSheet_AddUserRow(MSVisio.Shape shape, XElement addUserRowElement)
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
                    VNCVisioAddIn.Helpers.Add_User_Row(shape, row, value, promptAttribute.Value);
                }
                else
                {
                    VNCVisioAddIn.Helpers.Add_User_Row(shape, row, value);
                }
            }
            catch (Exception ex)
            {
                VNCVisioAddIn.Common.WriteToWatchWindow(ex.ToString());
            }
        }

        #endregion

        #endregion

        #region IInstanceCount

        private static int _instanceCountVM;

        public int InstanceCountVM
        {
            get => _instanceCountVM;
            set => _instanceCountVM = value;
        }

        #endregion
    }
}
