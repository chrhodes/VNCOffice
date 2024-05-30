using Prism.Commands;

using SupportTools_Excel.AzureDevOpsExplorer.Domain;
using SupportTools_Excel.AzureDevOpsExplorer.Presentation.ModelWrappers;
using SupportTools_Excel.AzureDevOpsExplorer.Presentation.Views;
using SupportTools_Excel.Core.Presentation.ViewModels;

using VNC;
using VNC.Core.Mvvm;

namespace SupportTools_Excel.AzureDevOpsExplorer.Presentation.ViewModels
{
    public class TeamProjectActionsViewModel : ViewModelBase, IAZDOTeamProjectActionsViewModel
    {
        #region Constructors and Load

        // View First
        // View creates new ViewModel in code or Xaml
        // or ViewModel passed into View constructor

        public TeamProjectActionsViewModel()
        {
            long startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            // TODO(crhodes)
            // Decide if we want defaults
            //AZDOTeamProjectActions = new AZDOTeamProjectActionsWrapper(new Domain.AZDOTeamProjectActions());

            InitializeViewModel();

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // ViewModel First
        // Calling base(view) wires this ViewModel into the View

        public TeamProjectActionsViewModel(TeamProjectActions view) : base(view)
        {
            long startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            InitializeViewModel();

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void InitializeViewModel()
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            GetTPInfoCommand = new DelegateCommand(OnGetTPInfoExecute, OnGetTPInfoCanExecute);
            GetTPXMLCommand = new DelegateCommand(OnGetTPXMLExecute, OnGetTPXMLCanExecute);

            TeamProjectActionRequest = new TeamProjectActionRequestWrapper(
                new TeamProjectActionRequest());

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region Fields

        #endregion

        #region Properties

        private TeamProjectActionRequestWrapper _teamProjectActionRequest;
        public TeamProjectActionRequestWrapper TeamProjectActionRequest
        {
            get { return _teamProjectActionRequest; }
            set
            {
                if (_teamProjectActionRequest == value)
                    return;
                _teamProjectActionRequest = value;
                OnPropertyChanged();
            }
        }

        //public System.Collections.ObjectModel.ObservableCollection<string> BSSectionsSelected { get; set; }
        //public System.Collections.ObjectModel.ObservableCollection<string> TMSectionsSelected { get; set; }
        //public System.Collections.ObjectModel.ObservableCollection<string> TPSectionsSelected { get; set; }
        //public System.Collections.ObjectModel.ObservableCollection<string> VCSSectionsSelected { get; set; }
        //public System.Collections.ObjectModel.ObservableCollection<string> WISSectionsSelected { get; set; }

        #endregion

        #region Commands

        #region GetTPInfo Command

        public DelegateCommand GetTPInfoCommand { get; set; }
        public string GetTPInfoContent { get; set; } = "Get Team Project Information";
        public string GetTPInfoToolTip { get; set; } = "Creates One WorkSheet containing selected Sections for Each Selected Team Project";

        public void OnGetTPInfoExecute()
        {
            long startTicks = Log.EVENT("Enter", Common.LOG_CATEGORY);

            Common.EventAggregator.GetEvent<GetTeamProjectInfoEvent>().Publish(TeamProjectActionRequest.Model);

            Log.EVENT("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public bool OnGetTPInfoCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion

        #region GetTPXML Command

        public DelegateCommand GetTPXMLCommand { get; set; }
        public string GetTPXMLContent { get; set; } = "Get Team Project XML";
        public string GetTPXMLToolTip { get; set; } = "Gets XML Definition for Team Project using Hosted XML Process Model";

        public void OnGetTPXMLExecute()
        {
            long startTicks = Log.EVENT("Enter", Common.LOG_CATEGORY);

            // TODO(crhodes)
            // May not need request

            Common.EventAggregator.GetEvent<GetTeamProjectXMLEvent>().Publish(TeamProjectActionRequest.Model);

            Log.EVENT("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public bool OnGetTPXMLCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion

        #endregion Commands

    }
}
