using System;

using SupportTools_Visio.Actions;
using SupportTools_Visio.Domain;
using SupportTools_Visio.Presentation.ModelWrappers;

using VNC;

using Visio = Microsoft.Office.Interop.Visio;
using VNCVisioAddIn = VNC.Visio.VSTOAddIn;
namespace SupportTools_Visio.Presentation.ViewModels
{
    public class ConnectionPointsViewModel : ShapeSheetSectionBase //, IConnectionPointRowViewModelViewModel
    {
        public ConnectionPointsViewModel()
        {
            Int64 startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);
            UpdateButtonContent = "Update ConnectionPoints for selected shapes";
            // TODO(crhodes)
            // Decide if we want defaults
            //ConnectionPointRowViewModel = new ConnectionPointRowWrapper(new Domain.ConnectionPointRowViewModel());

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public System.Collections.ObjectModel.ObservableCollection<ConnectionPointRowWrapper> ConnectionPoints { get; set; }


        ConnectionPointRowWrapper _selectedItem;
        public ConnectionPointRowWrapper SelectedItem
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
            //    //Visio_Shape.Set_ConnectionPointRowViewModel_Section(shape, ConnectionPointRowViewModel.Model);
            //}

            Globals.ThisAddIn.Application.EndUndoScope(undoScope, true);

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY);
        }

        public override void OnLoadCurrentSettingsExecute()
        {
            Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Visio.Application app = Globals.ThisAddIn.Application;

            Visio.Selection selection = app.ActiveWindow.Selection;

            ConnectionPoints = new System.Collections.ObjectModel.ObservableCollection<ConnectionPointRowWrapper>();

            foreach (Visio.Shape shape in selection)
            {
                foreach (VNCVisioAddIn.Domain.ConnectionPointRow row in Visio_Shape.Get_ConnectionPointRows(shape))
                {
                    ConnectionPoints.Add(new ConnectionPointRowWrapper(row));
                }
            }

            OnPropertyChanged("ConnectionPoints");

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY);
        }
    }
}
