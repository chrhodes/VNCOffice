using System;
using System.Collections.ObjectModel;
using System.Linq;

using Prism.Commands;

using SupportTools_Excel.AzureDevOpsExplorer.Domain;
using SupportTools_Excel.AzureDevOpsExplorer.Presentation.ModelWrappers;
using SupportTools_Excel.AzureDevOpsExplorer.Presentation.Views;
using SupportTools_Excel.Core.Presentation.ViewModels;
using SupportTools_Excel.Presentation.ModelWrappers;

using VNC;
using VNC.Core.Mvvm;

namespace SupportTools_Excel.AzureDevOpsExplorer.Presentation.ViewModels
{
    public class OptionsViewModel : ViewModelBase, IAZDOOptionsViewModel
    {
        #region Constructors and Load

        // View First
        // View creates new ViewModel in code or Xaml
        // or ViewModel passed into View constructor

        public OptionsViewModel()
        {
            long startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            DoSomethingCommand = new DelegateCommand(OnDoSomethingExecute, OnDoSomethingCanExecute);
            DoSomethingContent = "Update Actions for selected shapes";
            DoSomethingToolTip = "ToolTip for DoSomething Button";

            InitializeViewModel();
            // TODO(crhodes)
            // Decide if we want defaults
            //XXX = new XXXWrapper(new Domain.XXX());

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // ViewModel First
        // Calling base(view) wires this ViewModel into the View

        public OptionsViewModel(Options view) : base(view)
        {
            long startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            DoSomethingCommand = new DelegateCommand(OnDoSomethingExecute, OnDoSomethingCanExecute);
            DoSomethingContent = "Update Actions for selected shapes";
            DoSomethingToolTip = "ToolTip for DoSomething Button";

            InitializeViewModel();

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void InitializeViewModel()
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            InitializeOptions();
            InitializeTeamProjects();

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void InitializeOptions()
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            Options = new Options_AZDO_TFSWrapper(new Options_AZDO_TFS());

            Options.GoBackDays = 30;
            Options.StartDate = DateTime.Now - TimeSpan.FromDays(Options.GoBackDays);
            Options.EndDate = DateTime.Now;

            Options.ShowAllNodeLevels = true;

            //Options.ShowWorkItemFieldData = true;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void InitializeTeamProjects()
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            TeamProjects = new ObservableCollection<string>();
            //{
            //    "Team One",
            //    "Team Two",
            //    "Team Three",
            //    "Team Four"
            //};

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

        private Options_AZDO_TFSWrapper _options;
        public Options_AZDO_TFSWrapper Options
        {
            get { return _options; }
            set
            {
                if (_options == value)
                    return;
                _options = value;
                OnPropertyChanged();
            }
        }

        private Options_ExcelWrapper _optionsExcel;
        public Options_ExcelWrapper OptionsExcel
        {
            get { return _optionsExcel; }
            set
            {
                if (_optionsExcel == value)
                    return;
                _optionsExcel = value;
                OnPropertyChanged();
            }
        }

        // TODO(crhodes)
        // This is for a Grid or List

        public ObservableCollection<string> TeamProjects { get; set; }

        public ObservableCollection<string> WorkItemTypes { get; set; }

        //public WorkItemQuery workItemQuery

        // and the SelectedItem in the Grid or List or ComboBox, etc.

        ObservableCollection<string> _selectedTeamProjectSelectedItem2;
        public ObservableCollection<string> SelectedTeamProjectSelectedItem2
        {
            get
            {
                return _selectedTeamProjectSelectedItem2;
            }
            set
            {
                _selectedTeamProjectSelectedItem2 = value;
                OnPropertyChanged();
            }
        }

        string _selectedTeamProjectSelectedItem;
        public string SelectedTeamProjectSelectedItem
        {
            get
            {
                return _selectedTeamProjectSelectedItem;
            }
            set
            {
                _selectedTeamProjectSelectedItem = value;
                OnPropertyChanged();
            }
        }

        string _selectedTeamProjectText;
        public string SelectedTeamProjectText
        {
            get
            {
                return _selectedTeamProjectText;
            }
            set
            {
                _selectedTeamProjectText = value;
                OnPropertyChanged();
            }
        }

        string _selectedTeamProjectEditValue;
        public string SelectedTeamProjectEditValue
        {
            get
            {
                return _selectedTeamProjectEditValue;
            }
            set
            {
                _selectedTeamProjectEditValue = value;
                OnPropertyChanged();
            }
        }

        ObservableCollection<string> _selectedTeamProjects;
        public ObservableCollection<string> SelectedTeamProjects
        {
            get
            {
                return _selectedTeamProjects;
            }
            set
            {
                _selectedTeamProjects = value;
                
                OnPropertyChanged();
            }
        }

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

        #endregion Commands



        public Options_AZDO_TFS GetOptions()
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            Options_AZDO_TFS options = Options.Model;

            if (((QueriesViewModel)((Options)View).ucQueries.ViewModel).SelectedQuery != null)
            {
                Options.Model.WorkItemQuerySpec = ((QueriesViewModel)((Options)View).ucQueries.ViewModel).SelectedQuery.Model;
            }

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return options;
        }
    }
}
