using System;
using System.Windows;

using VNC;
using VNC.Core.Mvvm;

using MSVisio = Microsoft.Office.Interop.Visio;
using VNCVisioAddIn = VNC.Visio.VSTOAddIn;

namespace VNCShapeSheetApplication.Presentation.ViewModels
{
    public class ObjectViewModel<TInfo, TInfoWrapper> : ShapeSheetSectionBase
        where TInfoWrapper : ModelWrapper<TInfo>, new()
    {
        public ObjectViewModel(string updateButtonMessage, GetRow getRowMethod, SetRow setRowMethod, VNCVisioAddIn.Domain.ShapeType shapeType) 
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

        public delegate TInfo GetRow(MSVisio.Shape shape);
        public delegate void SetRow(MSVisio.Shape shape, TInfo values);

        GetRow _getCommand;
        SetRow _setCommand;

        VNCVisioAddIn.Domain.ShapeType _shapeType;
        
        public override void OnUpdateSettingsExecute()
        {
            Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            // Wrap a big, OMG, what have I done ???, undo around the whole thing !!!
            int undoScope = Common.VisioApplication.BeginUndoScope("UpdateRow");

            MSVisio.Application app = Common.VisioApplication;

            switch (_shapeType)
            {
                case VNCVisioAddIn.Domain.ShapeType.Document:
                    SetShapeRow(((MSVisio.Document)app.ActiveDocument).DocumentSheet, Info);
                    break;

                case VNCVisioAddIn.Domain.ShapeType.Page:
                    SetShapeRow(((MSVisio.Page)app.ActivePage).PageSheet, Info);
                    break;

                case VNCVisioAddIn.Domain.ShapeType.Shape:
                    MSVisio.Selection selection = app.ActiveWindow.Selection;

                    foreach (MSVisio.Shape shape in selection)
                    {
                        SetShapeRow(shape, Info);
                    }
                    break;

                default:
                    MessageBox.Show($"Unexpected _shapeType({_shapeType.GetType()}");
                    break;
            }

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

        public override void OnLoadCurrentSettingsExecute()
        {
            Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            MSVisio.Application app = Common.VisioApplication;

            switch (_shapeType)
            {
                case VNCVisioAddIn.Domain.ShapeType.Document:
                    GetRowFromShape(((MSVisio.Document)app.ActiveDocument).DocumentSheet);
                    break;

                case VNCVisioAddIn.Domain.ShapeType.Page:
                    GetRowFromShape(((MSVisio.Page)app.ActivePage).PageSheet);
                    break;

                case VNCVisioAddIn.Domain.ShapeType.Shape:
                    MSVisio.Selection selection = app.ActiveWindow.Selection;

                    foreach (MSVisio.Shape shape in selection)
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

        private void GetRowFromShape(MSVisio.Shape shape)
        {
            Info = _getCommand(shape);
            InfoWrapper = (TInfoWrapper)Activator.CreateInstance(typeof(TInfoWrapper), Info);

            OnPropertyChanged("InfoWrapper");
        }

        private void SetShapeRow(MSVisio.Shape shape, TInfo values)
        {
            _setCommand(shape, values);
        }
    }
}
