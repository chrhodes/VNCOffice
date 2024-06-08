using System;

using Prism.Commands;

using SupportTools_Visio.Presentation.ModelWrappers;

using Visio = Microsoft.Office.Interop.Visio;

using VNC;
using VNC.Core.Mvvm;
using Microsoft.Office.Interop.Visio;
using SupportTools_Visio.Actions;
using VNCVisioAddIn = VNC.Visio.VSTOAddIn;

namespace SupportTools_Visio.Presentation.ViewModels
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

        public ParagraphWrapper Paragraph { get; set; }

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
            Paragraph = new ParagraphWrapper(new VNCVisioAddIn.Domain.ParagraphRow());
        }

        public void OnUpdateSettingsExecute()
        {
            Log.Trace("Enter", Common.LOG_CATEGORY);
            // Wrap a big, OMG, what have I done ???, undo around the whole thing !!!

            int undoScope = Globals.ThisAddIn.Application.BeginUndoScope("UpdateParagraphSection");

            // Just need to pass in the model.

            Visio.Application app = Globals.ThisAddIn.Application;

            Visio.Selection selection = app.ActiveWindow.Selection;

            // Verify only one shape, for now just grab first.

            foreach (Visio.Shape shape in selection)
            {
                VNC.Visio.VSTOAddIn.Domain.ParagraphRow.Set_Paragraph_Section(shape, Paragraph.Model);
            }

            Globals.ThisAddIn.Application.EndUndoScope(undoScope, true);
            Log.Trace("Exit", Common.LOG_CATEGORY);
        }

        public Boolean OnUpdateSettingsCanExecute()
        {
            // TODO(crhodes)
            // Validate we have new settings

            return true;
        }

        void OnLoadCurrentSettingsExecute()
        {
            Visio.Application app = Globals.ThisAddIn.Application;
            
            Visio.Selection selection = app.ActiveWindow.Selection;

            // Verify only one shape, for now just grab first.

            foreach (Visio.Shape shape in selection)
            {
                Paragraph = new ParagraphWrapper(VNC.Visio.VSTOAddIn.Domain.ParagraphRow.Get_ParagraphSection(shape));
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
