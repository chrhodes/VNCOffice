using System;
using System.Windows;

using Microsoft.Office.Interop.Visio;

using SupportTools_Visio.Domain;

using VNC;
using VNC.Core.Mvvm;

using Visio = Microsoft.Office.Interop.Visio;

namespace SupportTools_Visio.Presentation.ViewModels
{
    public class ObjectViewModel<TInfo, TInfoWrapper> : ShapeSheetSectionBase
        where TInfoWrapper : ModelWrapper<TInfo>, new()
    {
        public ObjectViewModel(string updateButtonMessage, GetRow getRowMethod, SetRow setRowMethod, ShapeType shapeType) 
            : base()
        {
            Int64 startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            UpdateButtonContent = updateButtonMessage;
            _getCommand = getRowMethod;
            _setCommand = setRowMethod;
            _shapeType = shapeType;

            OnLoadCurrentSettingsExecute();
            // TODO(crhodes)
            // Decide if we want defaults
            //XXX = new XXXWrapper(new Domain.XXX());

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public TInfoWrapper InfoWrapper { get; set; }
        public TInfo Info { get; set; }

        public delegate TInfo GetRow(Visio.Shape shape);
        public delegate void SetRow(Visio.Shape shape, TInfo values);

        GetRow _getCommand;
        SetRow _setCommand;

        ShapeType _shapeType;
        
        //TInfoWrapper _selectedItem;
        //public TInfoWrapper SelectedItem
        //{
        //    get
        //    {
        //        return _selectedItem;
        //    }
        //    set
        //    {
        //        _selectedItem = value;
        //        OnPropertyChanged();
        //    }
        //}

        public override void OnUpdateSettingsExecute()
        {
            Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            // Wrap a big, OMG, what have I done ???, undo around the whole thing !!!
            int undoScope = Globals.ThisAddIn.Application.BeginUndoScope("UpdateRow");

            Visio.Application app = Globals.ThisAddIn.Application;

            switch (_shapeType)
            {
                case ShapeType.Document:
                    SetShapeRow(((Visio.Document)app.ActiveDocument).DocumentSheet, Info);
                    break;

                case ShapeType.Page:
                    SetShapeRow(((Visio.Page)app.ActivePage).PageSheet, Info);
                    break;

                case ShapeType.Shape:
                    Visio.Selection selection = app.ActiveWindow.Selection;

                    foreach (Visio.Shape shape in selection)
                    {
                        SetShapeRow(shape, Info);
                    }
                    break;

                default:
                    MessageBox.Show($"Unexpected _shapeType({_shapeType.GetType()}");
                    break;
            }

            Globals.ThisAddIn.Application.EndUndoScope(undoScope, true);

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY);
        }

        public override bool OnUpdateSettingsCanExecute()
        {
            Visio.Application app = Globals.ThisAddIn.Application;

            switch (_shapeType)
            {
                case ShapeType.Document:
                    return app.ActiveDocument != null ? true : false;

                case ShapeType.Page:
                    return app.ActivePage != null ? true : false;

                case ShapeType.Shape:
                    return base.OnLoadCurrentSettingsCanExecute();

                default:
                    MessageBox.Show($"Unexpected _shapeType({_shapeType.GetType()}");
                    break;
            }

            return false;
        }

        public override void OnLoadCurrentSettingsExecute()
        {
            Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Visio.Application app = Globals.ThisAddIn.Application;

            switch (_shapeType)
            {
                case ShapeType.Document:
                    GetRowFromShape(((Visio.Document)app.ActiveDocument).DocumentSheet);
                    break;

                case ShapeType.Page:
                    GetRowFromShape(((Visio.Page)app.ActivePage).PageSheet);
                    break;

                case ShapeType.Shape:
                    Visio.Selection selection = app.ActiveWindow.Selection;

                    foreach (Visio.Shape shape in selection)
                    {
                        GetRowFromShape(shape);
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
            Visio.Application app = Globals.ThisAddIn.Application;

            switch (_shapeType)
            {
                case ShapeType.Document:
                    return app.ActiveDocument != null ? true : false;

                case ShapeType.Page:
                    return app.ActivePage != null ? true : false;

                case ShapeType.Shape:
                    return base.OnLoadCurrentSettingsCanExecute();

                default:
                    MessageBox.Show($"Unexpected _shapeType({_shapeType.GetType()}");
                    break;
            }

            return false;
        }

        private void GetRowFromShape(Shape shape)
        {
            Info = _getCommand(shape);
            InfoWrapper = (TInfoWrapper)Activator.CreateInstance(typeof(TInfoWrapper), Info);

            OnPropertyChanged("InfoWrapper");
        }

        private void SetShapeRow(Shape shape, TInfo values)
        {
            _setCommand(shape, values);
        }
    }
}
