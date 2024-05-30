using Prism.Commands;

using VNC;
using VNC.Core.Mvvm;

using SupportTools_Excel.ActiveDirectoryExplorer.Presentation.ModelWrappers;
using SupportTools_Excel.ActiveDirectoryExplorer.Presentation.Views;

namespace SupportTools_Excel.ActiveDirectoryExplorer.Presentation.ViewModels
{
    public class ADPickerViewModel : ViewModelBase
    {
        #region Constructors and Load

        // View First
        // View creates new ViewModel in code or Xaml
        // or ViewModel passed into View constructor

        public ADPickerViewModel()
        {
            long startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            // TODO(crhodes)
            // Decide if we want defaults
            //ADPicker = new ADPickerWrapper(new Domain.ADPicker());

            InitializeViewModel();

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // ViewModel First
        // Calling base(view) wires this ViewModel into the View

        public ADPickerViewModel(ADPicker view) : base(view)
        {
            long startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            InitializeViewModel();

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void InitializeViewModel()
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            // TODO(crhodes)
            // Initialize any controls and/or properties that need to be

            DoSomethingCommand = new DelegateCommand(OnDoSomethingExecute, OnDoSomethingCanExecute);
            DoSomethingContent = "Update Actions for selected shapes";
            DoSomethingToolTip = "ToolTip for DoSomething Button";

            Message_DoubleClick_Command = new DelegateCommand(Message_DoubleClick);

            //InitializeRows();

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region Fields

        #endregion

        #region Properties

        string _message = "Click Button to do something";
        public string Message
        {
            get
            {
                return _message;
            }
            set
            {
                _message = value;
                OnPropertyChanged();
            }
        }

        // TODO(crhodes)
        // This is for a Grid or List

        // public System.Collections.ObjectModel.ObservableCollection<string> SelectedFruits { get; set; }

        // public System.Collections.ObjectModel.ObservableCollection<ADPickerWrapper> Rows { get; set; }

        // // and the SelectedItem in the Grid or List

        // ADPickerWrapper _selectedItem;
        // public ADPickerWrapper SelectedItem
        // {
        // get
        // {
        // return _selectedItem;
        // }
        // set
        // {
        // _selectedItem = value;
        // OnPropertyChanged();
        // }
        // }

        // Don't forget to uncomment InitializeRows in InitializeViewModel()

        // void InitializeRows()
        // {
        // Rows = new System.Collections.ObjectModel.ObservableCollection<ADPickerWrapper>();
        // Rows.Add(new ADPickerWrapper(new Domain.ADPicker(){ StringProperty ="Red", IntProperty = 1}));
        // Rows.Add(new ADPickerWrapper(new Domain.ADPicker(){ StringProperty = "Green", IntProperty = 2 }));
        // Rows.Add(new ADPickerWrapper(new Domain.ADPicker(){ StringProperty = "Blue", IntProperty = 3 }));

        // OnPropertyChanged("Rows");
        // }		

        #endregion

        #region Commands

        #region DoSomething Command

        public DelegateCommand DoSomethingCommand { get; set; }
        public string DoSomethingContent { get; set; }
        public string DoSomethingToolTip { get; set; }

        public void OnDoSomethingExecute()
        {
            // TODO(crhodes)
            // Do something amazing.

            Message = "Cool, you did something!";
        }

        public bool OnDoSomethingCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.

            return true;
        }

        #endregion

        #region Control Commands (Not Buttons)

        public DelegateCommand Message_DoubleClick_Command { get; set; }

        public void Message_DoubleClick()
        {
            Message = "Message DoubleClicked!";
        }

        #endregion

        #endregion Commands

    }
}
