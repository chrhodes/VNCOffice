using System;
using System.Windows;
using System.Xml.Linq;

using VNC;
using VNC.Core.Mvvm;

using MSVisio = Microsoft.Office.Interop.Visio;
using VNCVisioAddIn = VNC.Visio.VSTOAddIn;

namespace VNCShapeSheetApplication.Presentation.ViewModels
{
    public class RowsViewModel<TRow, TRowWrapper> : ShapeSheetSectionBase
        where TRowWrapper : ModelWrapper<TRow>, new()
    {

        #region Constructors, Initialization, and Load

        public RowsViewModel(string updateButtonMessage, GetRows getRowsMethod, VNCVisioAddIn.Domain.ShapeType shapeType)
            : base()
        {
            Int64 startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            UpdateButtonContent = updateButtonMessage;
            _getRowsCommand = getRowsMethod;
            _shapeType = shapeType;

            InitializeViewModel();

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void InitializeViewModel()
        {
            Int64 startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            OnLoadCurrentSettingsExecute();

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region Enums (one)


        #endregion

        #region Structures (none)


        #endregion

        #region Fields and Properties

        public System.Collections.ObjectModel.ObservableCollection<TRowWrapper> Rows { get; set; }

        public delegate System.Collections.ObjectModel.ObservableCollection<TRow> GetRows(MSVisio.Shape shape);


        GetRows _getRowsCommand;

        VNCVisioAddIn.Domain.ShapeType _shapeType;

        TRowWrapper _selectedItem;
        public TRowWrapper SelectedItem
        {
            get
            {
                return _selectedItem;
            }
            set
            {
                _selectedItem = value;
                OnPropertyChanged();
            }
        }
        private XElement _exportedElement;
        public XElement ExportedElement
        {
            get => _exportedElement;
            set
            {
                _exportedElement = value;
                OnPropertyChanged();
            }
        }

        #endregion

        #region Event Handlers (none)


        #endregion

        #region Commands

        public override void OnLoadCurrentSettingsExecute()
        {
            Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            MSVisio.Application app = Common.VisioApplication;
            MSVisio.Window wnd = app.ActiveWindow;

            if (wnd.Type != (short)MSVisio.VisWinTypes.visDrawing)
            {
                MessageBox.Show("Must use from drawing window");
                return;
            }

            Rows = new System.Collections.ObjectModel.ObservableCollection<TRowWrapper>();

            switch (_shapeType)
            {
                case VNCVisioAddIn.Domain.ShapeType.Document:
                    GetRowsFromShape(((MSVisio.Document)app.ActiveDocument).DocumentSheet);
                    break;

                case VNCVisioAddIn.Domain.ShapeType.Page:
                    GetRowsFromShape(((MSVisio.Page)app.ActivePage).PageSheet);
                    break;

                case VNCVisioAddIn.Domain.ShapeType.Shape:
                    MSVisio.Selection selection = app.ActiveWindow.Selection;

                    foreach (MSVisio.Shape shape in selection)
                    {
                        GetRowsFromShape(shape);
                    }
                    break;

                default:
                    MessageBox.Show($"Unexpected _shapeType({_shapeType.GetType()}");
                    break;
            }

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY);
        }

        public override bool OnLoadCurrentSettingsCanExecute()
        {
            MSVisio.Application app = Common.VisioApplication;

            switch (_shapeType)
            {
                case VNCVisioAddIn.Domain.ShapeType.Document:
                    return app.ActiveDocument != null ? true : false;

                case VNCVisioAddIn.Domain.ShapeType.Page:
                    return app.ActivePage != null ? true : false;

                case VNCVisioAddIn.Domain.ShapeType.Shape:
                    return base.OnLoadCurrentSettingsCanExecute();

                default:
                    MessageBox.Show($"Unexpected _shapeType({_shapeType.GetType()}");
                    break;
            }

            return false;
        }

        public override void OnUpdateSettingsExecute()
        {
            Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            // Wrap a big, OMG, what have I done ???, undo around the whole thing !!!
            int undoScope = Common.VisioApplication.BeginUndoScope("UpdateRows");

            // Just need to pass in the model.

            //Visio.Application app = Common.VisioApplication;

            //Visio.Selection selection = app.ActiveWindow.Selection;

            //// Verify only one shape, for now just grab first.

            //foreach (Visio.Shape shape in selection)
            //{
            //    //Visio_Shape.Set_XXX_Section(shape, XXX.Model);
            //}

            Common.VisioApplication.EndUndoScope(undoScope, true);

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY);
        }

        public override bool OnUpdateSettingsCanExecute()
        {
            MSVisio.Application app = Common.VisioApplication;

            switch (_shapeType)
            {
                case VNCVisioAddIn.Domain.ShapeType.Document:
                    return app.ActiveDocument != null ? true : false;

                case VNCVisioAddIn.Domain.ShapeType.Page:
                    return app.ActivePage != null ? true : false;

                case VNCVisioAddIn.Domain.ShapeType.Shape:
                    return base.OnLoadCurrentSettingsCanExecute();

                default:
                    MessageBox.Show($"Unexpected _shapeType({_shapeType.GetType()}");
                    break;
            }

            return false;
        }

        public override void ExportSettingsExecute()
        {
            Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            MSVisio.Application app = Common.VisioApplication;

            Rows = new System.Collections.ObjectModel.ObservableCollection<TRowWrapper>();

            switch (_shapeType)
            {
                case VNCVisioAddIn.Domain.ShapeType.Document:
                    ExportRowsFromShape(((MSVisio.Document)app.ActiveDocument).DocumentSheet);
                    break;

                case VNCVisioAddIn.Domain.ShapeType.Page:
                    ExportRowsFromShape(((MSVisio.Page)app.ActivePage).PageSheet);
                    break;

                case VNCVisioAddIn.Domain.ShapeType.Shape:
                    MSVisio.Selection selection = app.ActiveWindow.Selection;

                    foreach (MSVisio.Shape shape in selection)
                    {
                        ExportRowsFromShape(shape);
                    }
                    break;

                default:
                    MessageBox.Show($"Unexpected _shapeType({_shapeType.GetType()}");
                    break;
            }

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY);
        }

        public override bool ExportSettingsCanExecute()
        {
            MSVisio.Application app = Common.VisioApplication;

            switch (_shapeType)
            {
                case VNCVisioAddIn.Domain.ShapeType.Document:
                    return app.ActiveDocument != null ? true : false;

                case VNCVisioAddIn.Domain.ShapeType.Page:
                    return app.ActivePage != null ? true : false;

                case VNCVisioAddIn.Domain.ShapeType.Shape:
                    return base.OnLoadCurrentSettingsCanExecute();

                default:
                    MessageBox.Show($"Unexpected _shapeType({_shapeType.GetType()}");
                    break;
            }

            return false;
        }

        #endregion

        #region Public Methods (none)


        #endregion

        #region Protected Methods (none)


        #endregion

        #region Private Methods


        #endregion

        private void GetRowsFromShape(MSVisio.Shape shape)
        {
            foreach (TRow row in _getRowsCommand(shape))
            {
                //Rows.Add(new TRowWrapper(row));
                Rows.Add((TRowWrapper)Activator.CreateInstance(typeof(TRowWrapper), row));
            }

            OnPropertyChanged("Rows");
        }

        private void ExportRowsFromShape(MSVisio.Shape shape)
        {

            var shapeSheet = new XElement("ShapeSheet");

            foreach (TRow row in _getRowsCommand(shape))
            {
                // TODO(crhodes)
                // This is where we have to figure out what is in Row and and where to export
                // Maybe we skip the write to file and just dump on UI for now.

                // HACK(crhodes)
                // There must be a good way to handle the type knowledge here
                // without a case statement for the different row types

                //var foo = row;
                //var bar = row as ShapeDataRow;

                //var barX = bar.ToXElement();
                //shapeSheet.Add(barX);

                var foo = row.GetType();

                switch (foo.Name)
                {
                    case "ShapeDataRow":
                        shapeSheet.Add((row as VNCVisioAddIn.Domain.ShapeDataRow).ToXElement());
                        break;

                    default:

                        break;
                }

                // NOTE(crhodes)
                // We know the type, e.g. Domain.ShapeDataRow
                // If we reflect over the properties or maybe have a method in the Domain Object that
                // Know how to build Xml we want.

                //Rows.Add(new TRowWrapper(row));
                //Rows.Add((TRowWrapper)Activator.CreateInstance(typeof(TRowWrapper), row));
            }

            //var context = new XElement("Command");
            //context.Add(new XAttribute("Name", "NewCommand"));
            //context.Add(new XAttribute("Description", "NewCommand from Export"));
            //context.Add(new XElement("Shapes"));
            //context.Element("Shapes").Add(new XElement("Shape"));

            var context = new XElement("Command"
                   , new XAttribute("Name", "NewCommand")
                   , new XAttribute("Description", "NewCommand from Export")
                   , new XElement("Shapes"
                        , new XElement("Shape", shapeSheet))
                   );

            //context.Element("Shape").Add(shapeSheet);

            //var context = new XElement("Shape");
            //context.Add(shapeSheet);
            ExportedElement = context;
        }
    }
}
