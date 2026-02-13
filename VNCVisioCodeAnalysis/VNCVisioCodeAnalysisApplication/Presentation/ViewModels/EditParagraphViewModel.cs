using System;

using Prism.Commands;

using VNC;
using VNC.Core.Mvvm;

using MSVisio = Microsoft.Office.Interop.Visio;
using VNCVisioAddIn = VNC.Visio.VSTOAddIn;

namespace VNCVisioToolsApplication.Presentation.ViewModels
{
    public class EditParagraphViewModel : ViewModelBase, IEditTextViewModel, IInstanceCountVM
    {
        #region Constructors, Initialization, and Load

        public EditParagraphViewModel()
        {
            UpdateSettings = new DelegateCommand(OnUpdateSettingsExecute, OnUpdateSettingsCanExecute);
            LoadCurrentSettings = new DelegateCommand(OnLoadCurrentSettingsExecute, OnLoadCurrentSettingsCanExecute);

            HorizontalAlignmentChoices = new System.Collections.ObjectModel.ObservableCollection<ItemInfo>()
            {
                new ItemInfo(){ Content="Left", Value="0"}
                , new ItemInfo(){ Content="Center", Value="1"}
                , new ItemInfo(){ Content="Right", Value="2"}
                , new ItemInfo(){ Content="TextControlBound", Value="=IF(Controls.Row_1 > Width, 0, IF(Controls.Row_1 < 0, 2, 1))"}
                , new ItemInfo(){ Content="Foo", Value="Bar"}
            };


            // TODO(crhodes)
            // Decide if we want defaults
            Paragraph = new VNCVisioAddIn.Presentation.ModelWrappers.ParagraphWrapper(new VNCVisioAddIn.Domain.ParagraphRow());
        }

        #endregion

        #region Enums (None)


        #endregion

        #region Structures (None)


        #endregion

        #region Fields and Properties

        public System.Collections.ObjectModel.ObservableCollection<ItemInfo> HorizontalAlignmentChoices { get; set; }


        ItemInfo _selectedHorizontalAlignmentItem;
        public ItemInfo SelectedHorizontalAlignmentItem
        {
            get
            {
                return _selectedHorizontalAlignmentItem;
            }
            set
            {
                _selectedHorizontalAlignmentItem = value;
                Paragraph.HAlign = value.Value;
            }
        }

        public VNCVisioAddIn.Presentation.ModelWrappers.ParagraphWrapper Paragraph { get; set; }

        #endregion

        #region Event Handlers (None)


        #endregion

        #region Commands

        #region UpdateSettings

        public DelegateCommand UpdateSettings { get; private set; }

        public void OnUpdateSettingsExecute()
        {
            Log.TRACE("Enter", Common.LOG_CATEGORY);
            // Wrap a big, OMG, what have I done ???, undo around the whole thing !!!

            int undoScope = Common.VisioApplication.BeginUndoScope("UpdateParagraphSection");

            // Just need to pass in the model.

            MSVisio.Application app = Common.VisioApplication;

            MSVisio.Selection selection = app.ActiveWindow.Selection;

            // Verify only one shape, for now just grab first.

            foreach (MSVisio.Shape shape in selection)
            {
                VNC.Visio.VSTOAddIn.Domain.ParagraphRow.SetRow(shape, Paragraph.Model);
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

        #endregion

        #region LoadCurrentSettings

        public DelegateCommand LoadCurrentSettings { get; private set; }

        void OnLoadCurrentSettingsExecute()
        {
            MSVisio.Application app = Common.VisioApplication;

            MSVisio.Selection selection = app.ActiveWindow.Selection;

            // Verify only one shape, for now just grab first.

            foreach (MSVisio.Shape shape in selection)
            {
                Paragraph = new VNCVisioAddIn.Presentation.ModelWrappers.ParagraphWrapper(VNCVisioAddIn.Domain.ParagraphRow.GetRow(shape));
                OnPropertyChanged("Paragraph");
            }
        }

        bool OnLoadCurrentSettingsCanExecute()
        {
            // TODO(crhodes)
            // Check if shape selected

            return true;
        }

        #endregion

        #endregion

        #region Public Methods (None)

        public void SomeMethod()
        {

        }

        #endregion

        #region Protected Methods (None)


        #endregion

        #region Private Methods (None)


        #endregion





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
