using Prism.Commands;

using SupportTools_Excel.AzureDevOpsExplorer.Domain;
using SupportTools_Excel.AzureDevOpsExplorer.Presentation.ModelWrappers;
using SupportTools_Excel.AzureDevOpsExplorer.Presentation.Views;
using SupportTools_Excel.Core.Presentation.ViewModels;

using VNC;
using VNC.Core.Mvvm;

namespace SupportTools_Excel.AzureDevOpsExplorer.Presentation.ViewModels
{
    public class TestManagementActionsViewModel : ViewModelBase, IAZDOTestManagementActionsViewModel
    {
        #region Constructors and Load

        // View First

        public TestManagementActionsViewModel()
        {
            long startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            // TODO(crhodes)
            // Decide if we want defaults
            //AZDOTestManagementActions = new AZDOTestManagementActionsWrapper(new Domain.AZDOTestManagementActions());

            InitializeViewModel();

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // ViewModel First

        public TestManagementActionsViewModel(TestManagementActions view) : base(view)
        {
            long startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            InitializeViewModel();

            //View = view;

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void InitializeViewModel()
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            GetTestPlanInfoCommand = new DelegateCommand(OnGetTestPlanInfoExecute, OnGetTestPlanInfoCanExecute);
            GetTestSuiteInfoCommand = new DelegateCommand(OnGetTestSuiteInfoExecute, OnGetTestSuiteInfoCanExecute);
            GetTestCaseInfoCommand = new DelegateCommand(OnGetTestCaseInfoExecute, OnGetTestCaseInfoCanExecute);

            AddPivotSummaryCommand = new DelegateCommand(OnAddPivotSummaryExecute, OnAddPivotSummaryCanExecute);

            TestPlanID_DoubleClickCommand = new DelegateCommand(OnTestPlanID_DoubleClick, OnTestPlanID_DoubleClickCanExecute);
            TestSuiteID_DoubleClickCommand = new DelegateCommand(OnTestSuiteID_DoubleClick, OnTestSuiteID_DoubleClickCanExecute);
            TestCaseID_DoubleClickCommand = new DelegateCommand(OnTestCaseID_DoubleClick, OnTestCaseID_DoubleClickCanExecute);

            TestPlanRequest = new TestPlanRequestWrapper(new TestPlanRequest());
            TestSuiteRequest = new TestSuiteRequestWrapper(new TestSuiteRequest());
            TestCaseRequest = new TestCaseRequestWrapper(new TestCaseRequest());

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region Fields

        #endregion

        #region Properties

        private TestPlanRequestWrapper _testPlanRequest;

        public TestPlanRequestWrapper TestPlanRequest
        {
            get => _testPlanRequest;
            set
            {
                if (_testPlanRequest == value)
                    return;
                _testPlanRequest = value;
                OnPropertyChanged();
            }
        }

        private TestSuiteRequestWrapper _testSuiteRequest;
        public TestSuiteRequestWrapper TestSuiteRequest
        {
            get => _testSuiteRequest;
            set
            {
                if (_testSuiteRequest == value)
                    return;
                _testSuiteRequest = value;
                OnPropertyChanged();
            }
        }

        private TestCaseRequestWrapper _testCaseRequest;
        public TestCaseRequestWrapper TestCaseRequest
        {
            get => _testCaseRequest;
            set
            {
                if (_testCaseRequest == value)
                    return;
                _testCaseRequest = value;
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

        // TODO(crhodes)
        // This is for a Grid or List

        // public System.Collections.ObjectModel.ObservableCollection<AZDOTestManagementActionsWrapper> Rows { get; set; }

        // // and the SelectedItem in the Grid or List

        // AZDOTestManagementActionsWrapper _selectedItem;
        // public AZDOTestManagementActionsWrapper SelectedItem
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

        // Don't forget to uncomment InitializeRows in Constructors

        // void InitializeRows()
        // {
        // Rows = new System.Collections.ObjectModel.ObservableCollection<AZDOTestManagementActionsWrapper>();
        // Rows.Add(new AZDOTestManagementActionsWrapper(new Domain.AZDOTestManagementActions(){ StringProperty ="Red", IntProperty = 1}));
        // Rows.Add(new AZDOTestManagementActionsWrapper(new Domain.AZDOTestManagementActions(){ StringProperty = "Green", IntProperty = 2 }));
        // Rows.Add(new AZDOTestManagementActionsWrapper(new Domain.AZDOTestManagementActions(){ StringProperty = "Blue", IntProperty = 3 }));

        // OnPropertyChanged("Rows");
        // }		

        #endregion

        #region Commands

        #region TestPlanID_DoubleClick Command

        public DelegateCommand TestPlanID_DoubleClickCommand { get; set; }

        public void OnTestPlanID_DoubleClick()
        {
            // Need to pass wrapper so PropertyChanged gets handled
            Common.EventAggregator.GetEvent<TestPlanIDDoubleClickEvent>().Publish(TestPlanRequest);
            //Common.EventAggregator.GetEvent<TestPlanIDDoubleClickEvent>().Publish(TestPlanRequest.Model);
        }

        public bool OnTestPlanID_DoubleClickCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion

        #region TestSuiteID_DoubleClick Command

        public DelegateCommand TestSuiteID_DoubleClickCommand { get; set; }

        public void OnTestSuiteID_DoubleClick()
        {
            Common.EventAggregator.GetEvent<TestSuiteIDDoubleClickEvent>().Publish(TestSuiteRequest);
        }

        public bool OnTestSuiteID_DoubleClickCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion

        #region TestCaseID_DoubleClick Command

        public DelegateCommand TestCaseID_DoubleClickCommand { get; set; }

        public void OnTestCaseID_DoubleClick()
        {
            Common.EventAggregator.GetEvent<TestCaseIDDoubleClickEvent>().Publish(TestCaseRequest);
        }

        public bool OnTestCaseID_DoubleClickCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion
        
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

        #region GetTestPlanInfo Command

        public DelegateCommand GetTestPlanInfoCommand { get; set; }
        public string GetTestPlanInfoContent { get; set; } = "Get TestPlan Info";
        public string GetTestPlanInfoToolTip { get; set; } = "Get TestPlan Info Tooltip";

        public void OnGetTestPlanInfoExecute()
        {
            Common.EventAggregator.GetEvent<GetTestPlanInfoEvent>().Publish(TestPlanRequest.Model);
        }

        public bool OnGetTestPlanInfoCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion

        #region GetTestSuiteInfo Command

        public DelegateCommand GetTestSuiteInfoCommand { get; set; }
        public string GetTestSuiteInfoContent { get; set; } = "Get TestSuite Info";
        public string GetTestSuiteInfoToolTip { get; set; } = "Get TestSuite Info Tooltip";

        public void OnGetTestSuiteInfoExecute()
        {
            Common.EventAggregator.GetEvent<GetTestSuiteInfoEvent>().Publish(TestSuiteRequest.Model);
        }

        public bool OnGetTestSuiteInfoCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion

        #region GetTestCaseInfo Command

        public DelegateCommand GetTestCaseInfoCommand { get; set; }
        public string GetTestCaseInfoContent { get; set; } = "Get TestCase Info";
        public string GetTestCaseInfoToolTip { get; set; } = "Get TestCase Info Tooltip";

        public void OnGetTestCaseInfoExecute()
        {
            Common.EventAggregator.GetEvent<GetTestCaseInfoEvent>().Publish(TestCaseRequest.Model);
        }

        public bool OnGetTestCaseInfoCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion

        #region AddPivotSummary Command

        public DelegateCommand AddPivotSummaryCommand { get; set; }
        public string AddPivotSummaryContent { get; set; } = "Add Pivot Summary";
        public string AddPivotSummaryToolTip { get; set; } = "Add Pivot Summary ToolTip";

        public void OnAddPivotSummaryExecute()
        {
            Common.EventAggregator.GetEvent<AddTestPlanPivotSummaryEvent>().Publish();
        }

        public bool OnAddPivotSummaryCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion

        #endregion Commands

    }
}
