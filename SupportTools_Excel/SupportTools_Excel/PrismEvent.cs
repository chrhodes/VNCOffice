using Prism.Events;
using VNC.TFS.User_Interface.User_Controls;
using SupportTools_Excel.Presentation.Views;
using System.Windows;
using SupportTools_Excel.Domain;
using SupportTools_Excel.Presentation.ModelWrappers;
using SupportTools_Excel.AzureDevOpsExplorer.Domain;
using SupportTools_Excel.AzureDevOpsExplorer.Presentation.ModelWrappers;

namespace SupportTools_Excel
{
    //public class RunQueryEvent : PubSubEvent { }
    //public class RunTeamProjectQueryEvent : PubSubEvent { }
    //public class RunTeamProjectQueriesEvent : PubSubEvent { }

    #region Active Directory (AD) Events

    public class AddUserEvent : PubSubEvent { }

    public class FindUserEvent : PubSubEvent { }

    #endregion

    #region Server

    public class GetConfigurationServerInfoEvent : PubSubEvent<wucTFSProvider_Picker> { }
    //public class SelectTeamProjectCollectionEvent : PubSubEvent<string> { }

    public class PopulateTeamProjectsEvent : PubSubEvent { }
    public class EnableMainUIEvent: PubSubEvent<Visibility> { };

    #endregion

    #region AZDOQueries

    //public class RunQueryEvent : PubSubEvent<wucTFSQuery_Picker> { }
    //public class RunTeamProjectQueryEvent : PubSubEvent<wucTFSQuery_Picker> { }
    //public class RunTeamProjectQueriesEvent : PubSubEvent<wucTFSQuery_Picker> { }

    //public class RunQueryEvent : PubSubEvent<WorkItemQuery> { }
    //public class RunTeamProjectQueryEvent : PubSubEvent<WorkItemQuery> { }
    //public class RunTeamProjectQueriesEvent : PubSubEvent<WorkItemQuery> { }

    #endregion

    #region AZDOWorkItems

    // Need to pass wrapper so PropertyChanged gets handled on DoubleClick events

    public class WorkItemIDDoubleClickEvent : PubSubEvent<WorkItemActionRequestWrapper> { }
    public class GetWorkItemInfoEvent : PubSubEvent<WorkItemActionRequest> { }
    public class AddPivotSummaryEvent : PubSubEvent<WorkItemActionRequest> { }

    #endregion

    #region AZDOTeamProjectActions

    public class GetTeamProjectInfoEvent : PubSubEvent<TeamProjectActionRequest> { };
    public class GetTeamProjectXMLEvent : PubSubEvent<TeamProjectActionRequest> { };

    #endregion

    #region AZDO Organization Actions

    public class GetTPCInfoEvent : PubSubEvent { }

    public class GetTPCAreasEvent : PubSubEvent { }
    public class GetBranchesEvent : PubSubEvent { }
    public class GetAllTPDevelopersEvent : PubSubEvent { }
    public class GetTPCMembersEvent : PubSubEvent { }
    public class GetTPCShelfsetsEvent : PubSubEvent { }
    public class GetTPCBuildDefinitionsEvent : PubSubEvent { }
    public class GetTPCTeamsEvent : PubSubEvent { }
    public class GetTPCWorkItemTypesEvent : PubSubEvent { }
    public class GetTPCWorkItemFieldsEvent : PubSubEvent { }
    public class GetTPCWorkItemDetailsEvent : PubSubEvent { }
    public class GetTPCWorkspacesEvent : PubSubEvent { }
    public class GetTPCLastChangesetEvent : PubSubEvent { }
    public class GetTPCWorkItemActivityEvent : PubSubEvent { }
    public class GetTPCTestPlansEvent : PubSubEvent { }
    public class GetTPCTestSuitesEvent : PubSubEvent { }
    public class GetTPCTestCasesEvent : PubSubEvent { }

    public class GetTPCReleasesEvent : PubSubEvent { }

    #endregion

    #region AZDOTestManagement

    // Need to pass wrapper so PropertyChanged gets handled on DoubleClick events

    public class TestPlanIDDoubleClickEvent : PubSubEvent<TestPlanRequestWrapper> { }
    public class GetTestPlanInfoEvent : PubSubEvent<TestPlanRequest> { }

    public class TestSuiteIDDoubleClickEvent : PubSubEvent<TestSuiteRequestWrapper> { }
    public class GetTestSuiteInfoEvent : PubSubEvent<TestSuiteRequest> { }

    public class TestCaseIDDoubleClickEvent : PubSubEvent<TestCaseRequestWrapper> { }
    public class GetTestCaseInfoEvent : PubSubEvent<TestCaseRequest> { }


    public class AddTestPlanPivotSummaryEvent : PubSubEvent { }

    #endregion

}
