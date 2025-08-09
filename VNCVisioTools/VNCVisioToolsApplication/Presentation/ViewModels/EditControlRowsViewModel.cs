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

    //public class ItemInfo
    //{
    //    public string Content { get; set; }
    //    public string Value { get; set; }
    //}

    public class EditControlRowsViewModel : ViewModelBase, IEditTextViewModel
    {
        public System.Collections.ObjectModel.ObservableCollection<VNCVisioAddIn.Domain.ControlsRow> ControlRows { get; set; }


        //ItemInfo _selectedHorizontalAlignmentItem;
        //public ItemInfo SelectedHorizontalAlignmentItem
        //{
        //    get
        //    {
        //        return _selectedHorizontalAlignmentItem;
        //    }
        //    set
        //    {
        //        _selectedHorizontalAlignmentItem = value;
        //        Paragraph.HAlign = value.Value;
        //    }
        //}

        public DelegateCommand UpdateSettings { get; private set; }
        public DelegateCommand LoadCurrentSettings { get; private set; }

        public VNCVisioAddIn.Presentation.ModelWrappers.ControlsRowWrapper ControlRow { get; set; }

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


        public EditControlRowsViewModel()
        {
            UpdateSettings = new DelegateCommand(OnUpdateSettingsExecute, OnUpdateSettingsCanExecute);
            LoadCurrentSettings = new DelegateCommand(OnLoadCurrentSettingsExecute, OnLoadCurrentSettingsCanExecute);

            //HorizontalAlignmentChoices = new System.Collections.ObjectModel.ObservableCollection<ItemInfo>()
            //{
            //    new ItemInfo(){ Content="Left", Value="0"}
            //    , new ItemInfo(){ Content="Center", Value="1"}
            //    , new ItemInfo(){ Content="Right", Value="2"}
            //    , new ItemInfo(){ Content="TextControlBound", Value="=IF(Controls.Row_1 > Width, 0, IF(Controls.Row_1 < 0, 2, 1))"}
            //    , new ItemInfo(){ Content="Foo", Value="Bar"}
            //};


            // TODO(crhodes)
            // Decide if we want defaults
            //Paragraph = new ParagraphWrapper(new Domain.Paragraph());
        }

        public void OnUpdateSettingsExecute()
        {
            Log.Trace("Enter", Common.LOG_CATEGORY);
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
    }
}
