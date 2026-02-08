using System;

using Prism.Commands;

using VNCVisioToolsApplication.Presentation.ModelWrappers;

using MSVisio = Microsoft.Office.Interop.Visio;

using VNC;
using VNC.Core.Mvvm;
using Microsoft.Office.Interop.Visio;
using VNCVisioToolsApplication.Actions;
using VNCVisioToolsApplication.Domain;
using VNCVisioAddIn = VNC.Visio.VSTOAddIn;

namespace VNCVisioToolsApplication.Presentation.ViewModels
{
    public class EditControlRowsViewModel : ViewModelBase, IEditTextViewModel, IInstanceCountVM
    {

        #region Constructors, Initialization, and Load

        public EditControlRowsViewModel()
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.Constructor) startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            InstanceCountVM++;

            // TODO(crhodes)
            // Save constructor parameters here

            InitializeViewModel();

            if (Common.VNCLogging.Constructor) Log.CONSTRUCTOR($"Exit VM:{InstanceCountVM}", Common.LOG_CATEGORY, startTicks);
        }

        private void InitializeViewModel()
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.ViewModelLow) startTicks = Log.VIEWMODEL_LOW("Enter", Common.LOG_CATEGORY);

            // NOTE(crhodes)
            // Put things here that initialize the ViewModel
            // Initialize EventHandlers, Commands, etc.

            UpdateSettings = new DelegateCommand(OnUpdateSettingsExecute, OnUpdateSettingsCanExecute);
            LoadCurrentSettings = new DelegateCommand(OnLoadCurrentSettingsExecute, OnLoadCurrentSettingsCanExecute);

            if (Common.VNCLogging.ViewModelLow) Log.VIEWMODEL_LOW("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region Enums (None)


        #endregion

        #region Structures (None)


        #endregion

        #region Fields and Properties (None)


        #endregion

        #region Event Handlers (None)


        #endregion

        #region Commands (None)

        #endregion

        #region Public Methods (None)


        #endregion

        #region Protected Methods (None)


        #endregion

        #region Private Methods (None)


        #endregion
        public System.Collections.ObjectModel.ObservableCollection<VNCVisioAddIn.Domain.ControlsRow> ControlRows { get; set; }

        public DelegateCommand UpdateSettings { get; private set; }
        public DelegateCommand LoadCurrentSettings { get; private set; }

        public VNCVisioAddIn.Presentation.ModelWrappers.ControlsRowWrapper ControlRow { get; set; }

        public void OnUpdateSettingsExecute()
        {
            Log.TRACE("Enter", Common.LOG_CATEGORY);
            // Wrap a big, OMG, what have I done ???, undo around the whole thing !!!

            int undoScope = Common.VisioApplication.BeginUndoScope("UpdateControlRow");

            // Just need to pass in the model.

            MSVisio.Application app = Common.VisioApplication;

            MSVisio.Selection selection = app.ActiveWindow.Selection;

            // Verify only one shape, for now just grab first.

            foreach (MSVisio.Shape shape in selection)
            {
                //Visio_Shape.Set_Paragraph_Section(shape, Paragraph.Model);
            }

            Common.VisioApplication.EndUndoScope(undoScope, true);
            Log.TRACE("Exit", Common.LOG_CATEGORY);
        }

        public Boolean OnUpdateSettingsCanExecute()
        {
            // TODO(crhodes)
            // Validate we have new settings

            return true;
        }

        void OnLoadCurrentSettingsExecute()
        {
            MSVisio.Application app = Common.VisioApplication;
            
            MSVisio.Selection selection = app.ActiveWindow.Selection;

            // Verify only one shape, for now just grab first.

            foreach (MSVisio.Shape shape in selection)
            {
                ControlRow = new VNCVisioAddIn.Presentation.ModelWrappers.ControlsRowWrapper(VNCVisioAddIn.Domain.ControlsRow.GetRow(shape));
                OnPropertyChanged("ControlRow");
            }
        }

        bool OnLoadCurrentSettingsCanExecute()
        {
            // TODO(crhodes)
            // Check if shape selected

            return true;
        }

        public void SomeMethod()
        {
            
        }

        #region IInstanceCountVM

        private static int _instanceCountVM;

        public int InstanceCountVM
        {
            get => _instanceCountVM;
            set => _instanceCountVM = value;
        }

        #endregion
    }
}
