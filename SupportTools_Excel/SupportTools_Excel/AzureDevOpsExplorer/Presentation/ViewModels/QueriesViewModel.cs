using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Xml;
using System.Xml.Linq;

using Prism.Commands;

using SupportTools_Excel.AzureDevOpsExplorer.Presentation.ModelWrappers;
using SupportTools_Excel.AzureDevOpsExplorer.Presentation.Views;
using SupportTools_Excel.Core.Presentation.ViewModels;

using VNC;
using VNC.Core.Mvvm;

namespace SupportTools_Excel.AzureDevOpsExplorer.Presentation.ViewModels
{
    public class QueriesViewModel : ViewModelBase, IAZDOQueriesViewModel
    {
        #region Constructors and Load

        // View First
        // View creates new ViewModel in code or Xaml
        // or ViewModel passed into View constructor

        public QueriesViewModel()
        {
            long startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            // TODO(crhodes)
            // Decide if we want defaults
            //AZDOQueries = new AZDOQueriesWrapper(new Domain.AZDOQueries());

            InitializeViewModel();

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // ViewModel First
        // Calling base(view) wires this ViewModel into the View

        public QueriesViewModel(Queries view) : base(view)
        {
            long startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            InitializeViewModel();

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void InitializeViewModel()
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            //RunQueryCommand = new DelegateCommand(OnRunQueryExecute, OnRunQueryCanExecute);
            //RunTeamProjectQueryCommand = new DelegateCommand(OnRunTeamProjectQueryExecute, OnRunTeamProjectQueryCanExecute);
            //RunTeamProjectQueriesCommand = new DelegateCommand(OnRunTeamProjectQueriesExecute, OnRunTeamProjectQueriesCanExecute);
            QueryChangedCommand = new DelegateCommand(OnQueryChangedExecute, OnQueryChangedCanExecute);
            QueryDoubleClickCommand = new DelegateCommand(OnQueryDoubleClickExecute, OnQueryDoubleClickCanExecute);

            PopulateWorkItemQueries();
            PopulateWorkItemFields();

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void PopulateWorkItemFields()
        {
            WorkItemFields = new ObservableCollection<string>();

            using (XmlTextReader xtr = new XmlTextReader(Common.cCONFIG_FILE))
            {
                XDocument xDocument = XDocument.Load(xtr, LoadOptions.PreserveWhitespace);

                var fields = xDocument.Descendants("TFSQueries").Descendants("Fields");

                foreach (var field in fields.Elements())
                {
                    WorkItemFields.Add(field.Attribute("Name").Value);
                }
            }

            // HACK(crhodes)
            // Should either retrieve all the fields or drive this in the XML file
            // For now just hard code some fields that will likely be interesting.
            // NB.  Can use short name or ref name, eg. BugReason or Custom.BugReason
            // But, be mindful of spaces in names.

            //WorkItemFields.Add("Created By");
            //WorkItemFields.Add("Created Date");
            //WorkItemFields.Add("Changed By");
            //WorkItemFields.Add("Changed Date");

            //WorkItemFields.Add("CVSS Vector String");
            //WorkItemFields.Add("CVSSScore");

            //WorkItemFields.Add("DeferNote");
            //WorkItemFields.Add("DeferNoteHTML");
            //WorkItemFields.Add("Deferred By");
            //WorkItemFields.Add("Deferred Date");
            //WorkItemFields.Add("Deferred Fix Rating");

            //WorkItemFields.Add("Field Issue");
            //WorkItemFields.Add("Project ID");
            //WorkItemFields.Add("StoryPoints");

            //WorkItemFields.Add("BugReason");
            //WorkItemFields.Add("FeatureReason");
            //WorkItemFields.Add("IssueReason");
            //WorkItemFields.Add("ProductionIssueReason");
            //WorkItemFields.Add("ReleaseReason");
            //WorkItemFields.Add("RequestReason");
            //WorkItemFields.Add("TaskReason");
            //WorkItemFields.Add("TestCaseReason");
            //WorkItemFields.Add("TestPlanReason");
            //WorkItemFields.Add("TestSuiteReason");
            //WorkItemFields.Add("UserNeedsReason");
            //WorkItemFields.Add("UserStoryReason");

            //// NOTE(crhodes)
            //// Fields for PlainText (multiline) to HTML fix

            //// Bug

            //WorkItemFields.Add("CF.VSTS.HazardAnalysisHHS");
            //WorkItemFields.Add("Custom.STS_HazardAnalysisHHS");

            //WorkItemFields.Add("CF.VSTS.MilestoneChanges");
            //WorkItemFields.Add("Custom.STS_MilestoneChanges");

            //WorkItemFields.Add("CF.VSTS.MilestoneWithCompletionDate");
            //WorkItemFields.Add("Custom.STS_MilestoneWithCompletionDate");

            //WorkItemFields.Add("CF.VSTS.OfficialPOAMComments");
            //WorkItemFields.Add("Custom.STS_OfficialPOAMComments");

            //WorkItemFields.Add("CF.VSTS.RemediationDescription");
            //WorkItemFields.Add("Custom.STS_RemediationDescription");

            //// Release

            //WorkItemFields.Add("CF.VSTS.PRD.ChangeHistory");
            //WorkItemFields.Add("Custom.STS_ChangeHistoryPRD");

            //WorkItemFields.Add("CF.VSTS.PRD.Scope");
            //WorkItemFields.Add("Custom.STS_PRDScope");

            //WorkItemFields.Add("DevCustom.ReferenceDocumentsPRD");
            //WorkItemFields.Add("Custom.STS_ReferenceDocumentsPRD");

            //WorkItemFields.Add("CF.VSTS.PRD.UserDefinition");
            //WorkItemFields.Add("Custom.STS_UserDefinition");

            //// Request 

            //WorkItemFields.Add("CF.VSTS.RequestJustification");
            //WorkItemFields.Add("Custom.STS_RequestJustification");

            //// Test Case

            //WorkItemFields.Add("Custom.TestCase.TestData");
            //WorkItemFields.Add("Custom.STS_Data");

            //WorkItemFields.Add("Custom.TestCase.ExecutionSteps");
            //WorkItemFields.Add("Custom.STS_Execution");

            //WorkItemFields.Add("Custom.TestCase.SetupSteps");
            //WorkItemFields.Add("Custom.STS_Setup");
        }

        private void PopulateWorkItemQueries()
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            WorkItemQueries = new List<WorkItemQueryWrapper>();

            XmlTextReader xtr = new XmlTextReader(Common.cCONFIG_FILE);

            XDocument xDocument = XDocument.Load(xtr, LoadOptions.PreserveWhitespace);

            var queries = xDocument.Descendants("TFSQueries").Descendants("Queries");

            WorkItemQueries.Add(
                new WorkItemQueryWrapper(new Domain.WorkItemQuery()
                {
                    Name = "Default",
                    QueryWithTokens = "SELECT @FIELDS FROM WorkItems WHERE [System.TeamProject] = '@PROJECT'"
                }));

            SelectedQuery = WorkItemQueries[0];

            foreach (var query in queries.Elements())
            {
                //var nameV = query.Attribute("Name").Value;
                //var queryV = query.Attribute("Query").Value;

                WorkItemQueries.Add(
                    new WorkItemQueryWrapper(new Domain.WorkItemQuery()
                    {
                        Name = query.Attribute("Name").Value,
                        QueryWithTokens = query.Attribute("Query").Value
                    }));
            }

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region Fields

        #endregion

        #region Properties

        public ObservableCollection<string> WorkItemFields
        {
            get;
            set;
        }

        private List<WorkItemQueryWrapper> _workItemQueries;
        public List<WorkItemQueryWrapper> WorkItemQueries
        {
            get => _workItemQueries;
            set
            {
                if (_workItemQueries == value)
                    return;
                _workItemQueries = value;
                OnPropertyChanged();
            }
        }

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

        WorkItemQueryWrapper _selectedQuery;
        public WorkItemQueryWrapper SelectedQuery
        {
            get
            {
                return _selectedQuery;
            }
            set
            {
                _selectedQuery = value;
                OnPropertyChanged();
            }
        }

        #endregion

        #region Commands

        #region QueryChanged Command

        public DelegateCommand QueryChangedCommand { get; set; }

        public void OnQueryChangedExecute()
        {
            // TODO(crhodes)
            // Do something amazing.
            Message = "Query Changed";
        }

        public bool OnQueryChangedCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion

        #region QueryDoubleClick Command

        public DelegateCommand QueryDoubleClickCommand { get; set; }

        public void OnQueryDoubleClickExecute()
        {
            // TODO(crhodes)
            // Do something amazing.
            Message = "Query DoubleClick";
        }

        public bool OnQueryDoubleClickCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion


        #endregion Commands

    }
}
