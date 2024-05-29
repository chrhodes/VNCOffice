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
        public ObjectViewModel(string updateButtonMessage, GetInfo getInfoMethod, ShapeType shapeType) 
            : base()
        {
            Int64 startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            UpdateButtonContent = updateButtonMessage;
            _getCommand = getInfoMethod;
            _shapeType = shapeType;

            OnLoadCurrentSettingsExecute();
            // TODO(crhodes)
            // Decide if we want defaults
            //XXX = new XXXWrapper(new Domain.XXX());

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public TInfoWrapper Info { get; set; }

        public delegate TInfo GetInfo(Visio.Shape shape);

        GetInfo _getCommand;

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
            int undoScope = Globals.ThisAddIn.Application.BeginUndoScope("UpdateRows");

            // Just need to pass in the model.

            //Visio.Application app = Globals.ThisAddIn.Application;

            //Visio.Selection selection = app.ActiveWindow.Selection;

            //// Verify only one shape, for now just grab first.

            //foreach (Visio.Shape shape in selection)
            //{
            //    //Visio_Shape.Set_XXX_Section(shape, XXX.Model);
            //}

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
                    GetInfoFromShape(((Visio.Document)app.ActiveDocument).DocumentSheet);
                    break;

                case ShapeType.Page:
                    GetInfoFromShape(((Visio.Page)app.ActivePage).PageSheet);
                    break;

                case ShapeType.Shape:
                    Visio.Selection selection = app.ActiveWindow.Selection;

                    foreach (Visio.Shape shape in selection)
                    {
                        GetInfoFromShape(shape);
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

        private void GetInfoFromShape(Shape shape)
        {
            var info = _getCommand(shape);
            Info = (TInfoWrapper)Activator.CreateInstance(typeof(TInfoWrapper), info);

            OnPropertyChanged("Info");
        }
    }
}
