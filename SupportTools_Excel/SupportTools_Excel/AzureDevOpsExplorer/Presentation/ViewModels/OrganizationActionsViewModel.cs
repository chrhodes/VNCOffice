using Prism.Commands;

using VNC;
using VNC.Core.Mvvm;

using SupportTools_Excel.Presentation.ModelWrappers;
using SupportTools_Excel.Presentation.Views;
using SupportTools_Excel.Core.Presentation.ViewModels;
using SupportTools_Excel.AzureDevOpsExplorer.Presentation.Views;
using System;

namespace SupportTools_Excel.AzureDevOpsExplorer.Presentation.ViewModels
{
    public class OrganizationActionsViewModel : ViewModelBase, IAZDOOrganizationActionsViewModel
    {
        #region Constructors and Load

        // View First
        // View creates new ViewModel in code or Xaml
        // or ViewModel passed into View constructor

        public OrganizationActionsViewModel()
        {
            Int64 startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            // TODO(crhodes)
            // Decide if we want defaults
            //AZDOOrganizationsActions = new AZDOOrganizationsActionsWrapper(new Domain.AZDOOrganizationsActions());

            InitializeViewModel();

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // ViewModel First
        // Calling base(view) wires this ViewModel into the View

        public OrganizationActionsViewModel(OrganizationActions view) : base(view)
        {
            Int64 startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            InitializeViewModel();

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void InitializeViewModel()
        {
            Int64 startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            GetTPCInfoCommand = new DelegateCommand(OnGetTPCInfoExecute, OnGetTPCInfoCanExecute);
            GetTPCAreasCommand = new DelegateCommand(OnGetTPCAreasExecute, OnGetTPCAreasCanExecute);
            GetBranchesCommand = new DelegateCommand(OnGetBranchesExecute, OnGetBranchesCanExecute);
            GetAllTPDevelopersCommand = new DelegateCommand(OnGetAllTPDevelopersExecute, OnGetAllTPDevelopersCanExecute);
            GetTPCMembersCommand = new DelegateCommand(OnGetTPCMembersExecute, OnGetTPCMembersCanExecute);
            GetTPCShelfsetsCommand = new DelegateCommand(OnGetTPCShelfsetsExecute, OnGetTPCShelfsetsCanExecute);
            GetTPCBuildDefinitionsCommand = new DelegateCommand(OnGetTPCBuildDefinitionsExecute, OnGetTPCBuildDefinitionsCanExecute);
            GetTPCTeamsCommand = new DelegateCommand(OnGetTPCTeamsExecute, OnGetTPCTeamsCanExecute);
            GetTPCWorkItemTypesCommand = new DelegateCommand(OnGetTPCWorkItemTypesExecute, OnGetTPCWorkItemTypesCanExecute);
            GetTPCWorkItemFieldsCommand = new DelegateCommand(OnGetTPCWorkItemFieldsExecute, OnGetTPCWorkItemFieldsCanExecute);
            GetTPCWorkItemDetailsCommand = new DelegateCommand(OnGetTPCWorkItemDetailsExecute, OnGetTPCWorkItemDetailsCanExecute);
            GetTPCWorkspacesCommand = new DelegateCommand(OnGetTPCWorkspacesExecute, OnGetTPCWorkspacesCanExecute);
            GetTPCLastChangesetCommand = new DelegateCommand(OnGetTPCLastChangesetExecute, OnGetTPCLastChangesetCanExecute);
            GetTPCWorkItemActivityCommand = new DelegateCommand(OnGetTPCWorkItemActivityExecute, OnGetTPCWorkItemActivityCanExecute);
            GetTPCTestPlansCommand = new DelegateCommand(OnGetTPCTestPlansExecute, OnGetTPCTestPlansCanExecute);
            GetTPCTestSuitesCommand = new DelegateCommand(OnGetTPCTestSuitesExecute, OnGetTPCTestSuitesCanExecute);
            GetTPCTestCasesCommand = new DelegateCommand(OnGetTPCTestCasesExecute, OnGetTPCTestCasesCanExecute);

            //GetTPCReleasesCommand = new DelegateCommand(OnGetTPCReleasesExecute, OnGetTPCReleasesCanExecute);


            //TeamProjectActionRequest = new AZDOTeamProjectActionRequestWrapper(
            //    new Domain.AZDOTeamProjectActionRequest());
            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region Fields


        #endregion

        #region Properties


        #endregion

        #region Commands

        #region GetTPCInfo Command

        public DelegateCommand GetTPCInfoCommand { get; set; }
        public string GetTPCInfoContent { get; set; } = "Get TPC Information";
        public string GetTPCInfoToolTip { get; set; } = "Gets general information about the TeamProject Collection";

        public void OnGetTPCInfoExecute()
        {
            Common.EventAggregator.GetEvent<GetTPCInfoEvent>().Publish();
        }

        public bool OnGetTPCInfoCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion

        #region GetTPCAreas Command

        public DelegateCommand GetTPCAreasCommand { get; set; }
        public string GetTPCAreasContent { get; set; } = (string)Common.XamlApplication.FindResource("AZDOOrganizationActions_GetTPCAreasContent");
        public string GetTPCAreasToolTip { get; set; } = (string)Common.XamlApplication.FindResource("AZDOOrganizationActions_GetTPCAreasToolTip");

        public void OnGetTPCAreasExecute()
        {
            Common.EventAggregator.GetEvent<GetTPCAreasEvent>().Publish();
        }

        public bool OnGetTPCAreasCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion

        #region GetBranches Command

        public DelegateCommand GetBranchesCommand { get; set; }
        public string GetBranchesContent { get; set; } = (string)Common.XamlApplication.FindResource("AZDOOrganizationActions_GetBranchesContent");
        public string GetBranchesToolTip { get; set; } = (string)Common.XamlApplication.FindResource("AZDOOrganizationActions_GetBranchesToolTip");

        public void OnGetBranchesExecute()
        {
            Common.EventAggregator.GetEvent<GetBranchesEvent>().Publish();
        }

        public bool OnGetBranchesCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion

        #region GetAllTPDevelopers Command

        public DelegateCommand GetAllTPDevelopersCommand { get; set; }
        public string GetAllTPDevelopersContent { get; set; } = (string)Common.XamlApplication.FindResource("AZDOOrganizationActions_GetAllTPDevelopersContent");
        public string GetAllTPDevelopersToolTip { get; set; } = (string)Common.XamlApplication.FindResource("AZDOOrganizationActions_GetAllTPDevelopersToolTip");

        public void OnGetAllTPDevelopersExecute()
        {
            Common.EventAggregator.GetEvent<GetAllTPDevelopersEvent>().Publish();
        }

        public bool OnGetAllTPDevelopersCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion

        #region GetTPCMembers Command

        public DelegateCommand GetTPCMembersCommand { get; set; }
        public string GetTPCMembersContent { get; set; } = (string)Common.XamlApplication.FindResource("AZDOOrganizationActions_GetTPCMembersContent");
        public string GetTPCMembersToolTip { get; set; } = (string)Common.XamlApplication.FindResource("AZDOOrganizationActions_GetTPCMembersToolTip");

        public void OnGetTPCMembersExecute()
        {
            Common.EventAggregator.GetEvent<GetTPCMembersEvent>().Publish();
        }

        public bool OnGetTPCMembersCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion

        #region GetTPCShelfsets Command

        public DelegateCommand GetTPCShelfsetsCommand { get; set; }
        public string GetTPCShelfsetsContent { get; set; } = (string)Common.XamlApplication.FindResource("AZDOOrganizationActions_GetTPCShelfsetsContent");
        public string GetTPCShelfsetsToolTip { get; set; } = (string)Common.XamlApplication.FindResource("AZDOOrganizationActions_GetTPCShelfsetsToolTip");

        public void OnGetTPCShelfsetsExecute()
        {
            Common.EventAggregator.GetEvent<GetTPCShelfsetsEvent>().Publish();
        }

        public bool OnGetTPCShelfsetsCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion

        #region GetTPCBuildDefinitions Command

        public DelegateCommand GetTPCBuildDefinitionsCommand { get; set; }
        public string GetTPCBuildDefinitionsContent { get; set; } = (string)Common.XamlApplication.FindResource("AZDOOrganizationActions_GetTPCBuildDefinitionsContent");
        public string GetTPCBuildDefinitionsToolTip { get; set; } = (string)Common.XamlApplication.FindResource("AZDOOrganizationActions_GetTPCBuildDefinitionsToolTip");

        public void OnGetTPCBuildDefinitionsExecute()
        {
            Common.EventAggregator.GetEvent<GetTPCBuildDefinitionsEvent>().Publish();
        }

        public bool OnGetTPCBuildDefinitionsCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion

        #region GetTPCTeams Command

        public DelegateCommand GetTPCTeamsCommand { get; set; }
        public string GetTPCTeamsContent { get; set; } = (string)Common.XamlApplication.FindResource("AZDOOrganizationActions_GetTPCTeamsContent");
        public string GetTPCTeamsToolTip { get; set; } = (string)Common.XamlApplication.FindResource("AZDOOrganizationActions_GetTPCTeamsToolTip");

        public void OnGetTPCTeamsExecute()
        {
            Common.EventAggregator.GetEvent<GetTPCTeamsEvent>().Publish();
        }

        public bool OnGetTPCTeamsCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion

        #region GetTPCWorkItemTypes Command

        public DelegateCommand GetTPCWorkItemTypesCommand { get; set; }
        public string GetTPCWorkItemTypesContent { get; set; } = (string)Common.XamlApplication.FindResource("AZDOOrganizationActions_GetTPCWorkItemTypesContent");
        public string GetTPCWorkItemTypesToolTip { get; set; } = (string)Common.XamlApplication.FindResource("AZDOOrganizationActions_GetTPCWorkItemTypesToolTip");

        public void OnGetTPCWorkItemTypesExecute()
        {
            Common.EventAggregator.GetEvent<GetTPCWorkItemTypesEvent>().Publish();
        }

        public bool OnGetTPCWorkItemTypesCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion

        #region GetTPCWorkItemFields Command

        public DelegateCommand GetTPCWorkItemFieldsCommand { get; set; }
        public string GetTPCWorkItemFieldsContent { get; set; } = (string)Common.XamlApplication.FindResource("AZDOOrganizationActions_GetTPCWorkItemFieldsContent");
        public string GetTPCWorkItemFieldsToolTip { get; set; } = (string)Common.XamlApplication.FindResource("AZDOOrganizationActions_GetTPCWorkItemFieldsToolTip");

        public void OnGetTPCWorkItemFieldsExecute()
        {
            Common.EventAggregator.GetEvent<GetTPCWorkItemFieldsEvent>().Publish();
        }

        public bool OnGetTPCWorkItemFieldsCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion

        #region GetTPCWorkItemDetails Command

        public DelegateCommand GetTPCWorkItemDetailsCommand { get; set; }
        public string GetTPCWorkItemDetailsContent { get; set; } = (string)Common.XamlApplication.FindResource("AZDOOrganizationActions_GetTPCWorkItemDetailsContent");
        public string GetTPCWorkItemDetailsToolTip { get; set; } = (string)Common.XamlApplication.FindResource("AZDOOrganizationActions_GetTPCWorkItemDetailsToolTip");

        public void OnGetTPCWorkItemDetailsExecute()
        {
            Common.EventAggregator.GetEvent<GetTPCWorkItemDetailsEvent>().Publish();
        }

        public bool OnGetTPCWorkItemDetailsCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion

        #region GetTPCWorkspaces Command

        public DelegateCommand GetTPCWorkspacesCommand { get; set; } 
        public string GetTPCWorkspacesContent { get; set; } = (string)Common.XamlApplication.FindResource("AZDOOrganizationActions_GetTPCWorkspacesContent");
        public string GetTPCWorkspacesToolTip { get; set; } = (string)Common.XamlApplication.FindResource("AZDOOrganizationActions_GetTPCWorkspacesToolTip");

        public void OnGetTPCWorkspacesExecute()
        {
            Common.EventAggregator.GetEvent<GetTPCWorkspacesEvent>().Publish();
        }

        public bool OnGetTPCWorkspacesCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion

        #region GetTPCLastChangeset Command

        public DelegateCommand GetTPCLastChangesetCommand { get; set; }
        public string GetTPCLastChangesetContent { get; set; } = (string)Common.XamlApplication.FindResource("AZDOOrganizationActions_GetTPCLastChangesetContent");
        public string GetTPCLastChangesetToolTip { get; set; } = (string)Common.XamlApplication.FindResource("AZDOOrganizationActions_GetTPCLastChangesetToolTip");

        public void OnGetTPCLastChangesetExecute()
        {
            Common.EventAggregator.GetEvent<GetTPCLastChangesetEvent>().Publish();
        }

        public bool OnGetTPCLastChangesetCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion

        #region GetTPCLastWorkItem Command

        public DelegateCommand GetTPCWorkItemActivityCommand { get; set; }
        public string GetTPCWorkItemActivityContent { get; set; } = "Get TPC WorkItemActivity";
        public string GetTPCWorkItemActivityToolTip { get; set; } = "Get Count of WorkItems with Activity in Date Range";

        public void OnGetTPCWorkItemActivityExecute()
        {
            Common.EventAggregator.GetEvent<GetTPCWorkItemActivityEvent>().Publish();
        }

        public bool OnGetTPCWorkItemActivityCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion

        #region GetTPCTestPlans Command

        public DelegateCommand GetTPCTestPlansCommand { get; set; }
        public string GetTPCTestPlansContent { get; set; } = (string)Common.XamlApplication.FindResource("AZDOOrganizationActions_GetTPCTestPlansContent");
        public string GetTPCTestPlansToolTip { get; set; } = (string)Common.XamlApplication.FindResource("AZDOOrganizationActions_GetTPCTestPlansToolTip");

        public void OnGetTPCTestPlansExecute()
        {
            Common.EventAggregator.GetEvent<GetTPCTestPlansEvent>().Publish();
        }

        public bool OnGetTPCTestPlansCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion

        #region GetTPCTestSuites Command

        public DelegateCommand GetTPCTestSuitesCommand { get; set; }
        public string GetTPCTestSuitesContent { get; set; } = (string)Common.XamlApplication.FindResource("AZDOOrganizationActions_GetTPCTestSuitesContent");
        public string GetTPCTestSuitesToolTip { get; set; } = (string)Common.XamlApplication.FindResource("AZDOOrganizationActions_GetTPCTestSuitesToolTip");

        public void OnGetTPCTestSuitesExecute()
        {
            Common.EventAggregator.GetEvent<GetTPCTestSuitesEvent>().Publish();
        }

        public bool OnGetTPCTestSuitesCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion

        #region GetTPCTestCases Command

        public DelegateCommand GetTPCTestCasesCommand { get; set; }
        public string GetTPCTestCasesContent { get; set; } = (string)Common.XamlApplication.FindResource("AZDOOrganizationActions_GetTPCTestCasesContent");
        public string GetTPCTestCasesToolTip { get; set; } = (string)Common.XamlApplication.FindResource("AZDOOrganizationActions_GetTPCTestCasesToolTip");

        public void OnGetTPCTestCasesExecute()
        {
            Common.EventAggregator.GetEvent<GetTPCTestCasesEvent>().Publish();
        }

        public bool OnGetTPCTestCasesCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion

        //#region GetTPCReleases Command

        //public DelegateCommand GetTPCReleasesCommand { get; set; }
        //public string GetTPCReleasesContent { get; set; } = "GetTPCReleases";
        //public string GetTPCReleasesToolTip { get; set; } = "GetTPCReleases ToolTip";
        //// Can get fancy and use Resources
        ////public string GetTPCReleasesContent { get; set; } = "ViewName_GetTPCReleasesContent";
        ////public string GetTPCReleasesToolTip { get; set; } = "ViewName_GetTPCReleasesContentToolTip";

        //// Put these in Resource File

        ////    <system:String x:Key="ViewName_GetTPCReleasesContent">GetTPCReleases</system:String>
        ////    <system:String x:Key="ViewName_GetTPCReleasesContentToolTip">GetTPCReleases ToolTip</system:String>  

        //public void OnGetTPCReleasesExecute()
        //{
        //    Common.EventAggregator.GetEvent<GetTPCReleasesEvent>().Publish();
        //}

        //public bool OnGetTPCReleasesCanExecute()
        //{
        //    // TODO(crhodes)
        //    // Add any before button is enabled logic.
        //    return true;
        //}

        //#endregion

        #endregion Commands

    }
}
