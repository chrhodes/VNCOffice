using System;

using Prism.Commands;

using VNC;
using VNC.Core.Mvvm;

using MSVisio = Microsoft.Office.Interop.Visio;
using VNCVisioAddIn = VNC.Visio.VSTOAddIn;

namespace VNCVisioToolsApplication.Presentation.ViewModels
{

    public class ItemInfo
    {
        public string Content { get; set; }
        public string Value { get; set; }
    }

    public class EditParagraphViewModel : ViewModelBase, IEditTextViewModel
    {
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

        public DelegateCommand UpdateSettings { get; private set; }
        public DelegateCommand LoadCurrentSettings { get; private set; }

        public VNCVisioAddIn.Presentation.ModelWrappers.ParagraphWrapper Paragraph { get; set; }

        //{ 
        //    get => textBlockFormat; 
        //    set => textBlockFormat = value; 
        //}
        //private string message = "Fox Lady";
        //public string Message
        //{ 
        //    get => message; 
        //    set => message = value; 
        //}


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

        public void SomeMethod()
        {
            
        }
    }
}
