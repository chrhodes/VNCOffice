using System.Windows.Input;

using Microsoft.TeamFoundation.Build.Client;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.Framework.Client;
using Microsoft.TeamFoundation.Git.Client;
using Microsoft.TeamFoundation.Server;
using Microsoft.TeamFoundation.TestManagement.Client;
using Microsoft.TeamFoundation.VersionControl.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;

using Prism.Commands;

using SupportTools_Excel.Core.Presentation.ViewModels;

using VNC;
using VNC.Core.Mvvm;

namespace SupportTools_Excel.AzureDevOpsExplorer.Presentation.ViewModels
{
    public class ServerViewModel : ViewModelBase, IAZDOServerViewModel
    {
        #region Constructors and Load

        // View First
        // View creates new ViewModel in code or Xaml
        // or ViewModel passed into View constructor

        public ServerViewModel()
        {
            long startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            // TODO(crhodes)
            // Decide if we want defaults
            //Server = new ServerWrapper(new Domain.Server());

            InitializeViewModel();

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // ViewModel First
        // Calling base(view) wires this ViewModel into the View

        public ServerViewModel(AzureDevOpsExplorer.Presentation.Views.Server view) : base(view)
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

            TeamProjectCollections = new System.Collections.ObjectModel.ObservableCollection<string>();

            DoSomethingCommand = new DelegateCommand(OnDoSomethingExecute, OnDoSomethingCanExecute);
            DoSomethingContent = "Update Actions for selected shapes";
            DoSomethingToolTip = "ToolTip for DoSomething Button";

            //InitializeRows();
            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region Fields

        #endregion

        #region Properties

        public TfsConfigurationServer ConfigurationServer { get; set; }

        // These are updated when the Team Project Collection Changes

        public IBuildServer BuildServer { get; set; }
        public ICommonStructureService CommonStructureService { get; set; }
        public GitRepositoryService GitRepositoryService { get; set; }
        public IIdentityManagementService IdentityManagementService { get; set; }
        public TestManagementService TestManagementService { get; set; }
        public TfsTeamProjectCollection TfsTeamProjectCollection { get; set; }
        public VersionControlServer VersionControlServer { get; set; }
        public WorkItemStore WorkItemStore { get; set; }


        private string _ball;
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

        public System.Collections.ObjectModel.ObservableCollection<string> TeamProjectCollections { get; set; }

        // and the SelectedItem in the Grid or List

        string _selectedItem;
        public string SelectedItem
        {
            get
            {
                return _selectedItem;
            }
            set
            {
                // TODO(crhodes)
                // Need to move code out of Server.cbeTeamProjectCollections_SelectedIndexChanged

                _selectedItem = value;
                OnPropertyChanged();
            }
        }

        //Don't forget to uncomment InitializeRows in Constructors

        // void InitializeRows()
        //{
        //    Rows = new System.Collections.ObjectModel.ObservableCollection<ServerWrapper>();
        //    Rows.Add(new ServerWrapper(new Domain.Server() { StringProperty = "Red", IntProperty = 1 }));
        //    Rows.Add(new ServerWrapper(new Domain.Server() { StringProperty = "Green", IntProperty = 2 }));
        //    Rows.Add(new ServerWrapper(new Domain.Server() { StringProperty = "Blue", IntProperty = 3 }));

        //    OnPropertyChanged("Rows");
        //}

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

        public void OnServerProviderChanged()
        {

        }

        ICommand _serverProviderChanged;

        public ICommand ServerProviderChanged
        {
            get { return _serverProviderChanged; }
            set
            {
                OnServerProviderChanged();
                //if (_serverProviderChanged == value)
                //    return;
                //_serverProviderChanged = value;
                //OnPropertyChanged();
            }
        }

        string _uri;
        public string Uri
        {
            get { return _uri; }
            set
            {
                if (_uri == value)
                    return;
                _uri = value;
                OnPropertyChanged();
            }
        }
    }
}
