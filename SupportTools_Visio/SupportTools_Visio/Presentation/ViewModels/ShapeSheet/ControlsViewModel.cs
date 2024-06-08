using System;

using SupportTools_Visio.Actions;
using SupportTools_Visio.Domain;
using SupportTools_Visio.Presentation.ModelWrappers;

using VNC;

using Visio = Microsoft.Office.Interop.Visio;
using VNCVisioAddIn = VNC.Visio.VSTOAddIn;

namespace SupportTools_Visio.Presentation.ViewModels
{
    public class ControlsViewModel : ShapeSheetSectionBase
    {
        public ControlsViewModel()
        {
            Int64 startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            UpdateButtonContent = "Update Controls for selected shapes";

            // TODO(crhodes)
            // Decide if we want defaults
            //ControlsRowViewModel = new ControlsRowWrapper(new Domain.ControlsRowViewModel());

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public System.Collections.ObjectModel.ObservableCollection<ControlsRowWrapper> Controls { get; set; }

        ControlsRowWrapper _selectedItem;
        public ControlsRowWrapper SelectedItem
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

        public override void OnUpdateSettingsExecute()
        {
            Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            // Wrap a big, OMG, what have I done ???, undo around the whole thing !!!
            int undoScope = Globals.ThisAddIn.Application.BeginUndoScope("UpdateControlRow");

            // Just need to pass in the model.

            //Visio.Application app = Globals.ThisAddIn.Application;

            //Visio.Selection selection = app.ActiveWindow.Selection;

            //// Verify only one shape, for now just grab first.

            //foreach (Visio.Shape shape in selection)
            //{
            //    //Visio_Shape.Set_ControlsRowViewModel_Section(shape, ControlsRowViewModel.Model);
            //}

            Globals.ThisAddIn.Application.EndUndoScope(undoScope, true);
            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY);
        }

        public override void OnLoadCurrentSettingsExecute()
        {
            Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Visio.Application app = Globals.ThisAddIn.Application;

            Visio.Selection selection = app.ActiveWindow.Selection;

            Controls = new System.Collections.ObjectModel.ObservableCollection<ControlsRowWrapper>();

            foreach (Visio.Shape shape in selection)
            {
                foreach (VNCVisioAddIn.Domain.ControlsRow row in VNC.Visio.VSTOAddIn.Domain.ControlsRow.Get_ControlsRows(shape))
                {
                    Controls.Add(new ControlsRowWrapper(row));
                }
            }

            OnPropertyChanged("Controls");

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY);
        }
    }
}
