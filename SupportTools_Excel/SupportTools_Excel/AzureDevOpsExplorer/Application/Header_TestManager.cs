
using VNC.AddinHelper;

using XlHlp = VNC.AddinHelper.Excel;

namespace SupportTools_Excel.AzureDevOpsExplorer.Application
{
    public class Header_TestManager
    {
        #region TestManager (TM)

        internal static void Add_Queries(Excel.XlLocation insertAt)
        {
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "TestTeamProject");

            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Name");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Owner");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "QueryText");

            insertAt.IncrementRows();
        }

        internal static void Add_TestCases(Excel.XlLocation insertAt)
        {
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "TestTeamProject");

            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "ID");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Title");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Description");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "DateCreated");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "DateModified");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Revision");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "State");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Links");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "TestSuiteID");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "TestSuiteName");

            insertAt.IncrementRows();
        }

        internal static void Add_TestConfigurations(Excel.XlLocation insertAt)
        {
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "TestTeamProject");

            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "ID");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Name");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "AreaPath");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Description");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "IsDefault");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "LastUpdated");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "LastUpdatedBy");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Revision");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "State");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Values");

            insertAt.IncrementRows();
        }

        internal static void Add_TestEnvironments(Excel.XlLocation insertAt)
        {
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "TestTeamProject");

            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "ID");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Name");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "ControllerDisplayName");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "ControllerEnvironmentId");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "ControllerName");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "DateCreate");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Description");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "EnvironmentType");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Error");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "InvalidProperties");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "IsDirty");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "LabEnvironmentUri");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "LabServerUri");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "MachineRoles");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Owner");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "TeamProject.Name");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "TestController.Name");

            insertAt.IncrementRows();
        }

        internal static void Add_TestFailureTypes(Excel.XlLocation insertAt)
        {
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "TeamProject");

            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "ID");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Name");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Project.TeamProjectName");

            insertAt.IncrementRows();
        }

        internal static void Add_TestPlans(Excel.XlLocation insertAt)
        {
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "TestTeamProject");

            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "ID");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Name");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Description");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "State");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "AreaPath");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Iteration");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "StartDate");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "EndDate");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "OwnerName");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "LastUpdated");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "LastUpdatedByName");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Revision");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Links");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "RootSuiteId");

            insertAt.IncrementRows();
        }

        internal static void Add_TestPoints(Excel.XlLocation insertAt)
        {
            // TODO(crhodes)
            // Commented out rows show up in watch window

            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Id");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "AssignedToName");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Comment");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "ConfigurationId");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "ConfigurationName");
            //XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "FailureType");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "HasCachedProperties");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "History.Count");
            //XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "InvalidProperties");
            //XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "IsDirty");
            //XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "IsSaved");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "IsTestCaseAutomated");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "LastUpdated");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "LastUpdatedBy");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "MostRecentFailureType");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "MostRecentFailureTypeId");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "MostRecentResolutionStateId");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "MostRecentResult");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "MostRecentResultId");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "MostRecentResultOutcome");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "MostRecentResultState");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "MostRecentRunId");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Plan.Id");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Revision");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "State");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "SuiteId");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "TestCaseExists");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "TestCaseId");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "TestCaseWorkItem.Id");
            //XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "TotalInvalidProperties");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "UserData");

            insertAt.IncrementRows();
        }

        internal static void Add_TestResolutionStates(Excel.XlLocation insertAt)
        {
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "TestTeamProject");

            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Id");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Name");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Project.TeamProjectName");

            insertAt.IncrementRows();
        }

        internal static void Add_TestResults(Excel.XlLocation insertAt)
        {
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "TestTeamProject");

            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Attachments");

            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "CollectorsEnabled");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Comment");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "DateCreated");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "DateStarted");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "DateCompleted");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Duration");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "ErrorMessage");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Outcome");

            insertAt.IncrementRows();
        }

        internal static void Add_TestRuns(Excel.XlLocation insertAt)
        {
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "TestTeamProject");

            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Title");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Attachments");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "BuildConfigurationId");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "BuildDirectory");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "BuildFlavor");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "BuildNumber");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "BuildPlatform");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "BuildUri");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "IsBvt");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Iteration");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "LastUpdated");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "LastUpdatedBy");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "LinkedWorkItemCount");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "NotApplicableTests");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "OwnerName");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "PassedTests");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "PostProcessState");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "TeamProjectName");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Revision");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "State");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Statistics.CompletedTests");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Statistics.FailedTests");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Statistics.InconclusiveTests");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Statistics.InProgressTests");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Statistics.PassedTests");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Statistics.PendingTests");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Statistics.TotalTests");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "TestEnvironmentId");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "TestMessageLogEntries");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "TestPlanId");

            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "TestSettings.Name");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "TestSettings.Id");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "TestSettingsId");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "TotalTests");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Type");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "UnanalyzedTests");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Version");

            insertAt.IncrementRows();
        }

        internal static void Add_TestSettings(Excel.XlLocation insertAt)
        {
            XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 20, "TeamProject");

            XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 20, "Id");
            XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 20, "IsAutomated");
            XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 20, "LastUpdated");
            XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 20, "LastUpdatedBy");
            XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 20, "MachineRoles.Count");
            XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 20, "Name");
            XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 20, "Revision");
            XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 20, "Settings");

            insertAt.IncrementRows();
        }

        internal static void Add_TestSuites(Excel.XlLocation insertAt)
        {
            XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 20, "TeamProject");

            XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 20, "ID");
            XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 20, "Title");
            XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 20, "Description");
            XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 20, "State");
            XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 20, "LastUpdated");
            XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 20, "LastUpdatedBy");
            XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 20, "TestCaseCount");
            XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 20, "TestSuiteType");
            XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 20, "TestCasesCount");
            XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 20, "Revision");
            XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 20, "PlanID");
            XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 20, "PlanName");

            insertAt.IncrementRows();
        }

        internal static void Add_TestSuite_Entry(Excel.XlLocation insertAt)
        {
            XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 20, "ID");
            XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 20, "Title");
            XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 20, "Configurations");
            XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 20, "EntryType");
            XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 20, "PointAssignments");
            XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 20, "TestCase");
            XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 20, "TestObject");
            XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 20, "TestSuite");

            insertAt.IncrementRows();
        }

        internal static void Add_TestSuite_TestSuite(Excel.XlLocation insertAt)
        {
            XlHlp.AddLabeledInfo(insertAt.AddRow(), "Test Suite Entry", "Test Suites");

            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "ID");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Title");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "AllTestCases");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "DefaultConfigurations");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Description");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Error");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "InvalidProperties");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "IsDirty");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "IsRoot");

            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "LastUpdated");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "LastUpdatedByName");

            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Parent.Id");

            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Plan.Id");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Plan.Name");

            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Revision");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "State");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "TestCaseCount");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "TestSuiteType");

            insertAt.IncrementRows();
        }

        internal static void Add_TestSuite_TestCase(Excel.XlLocation insertAt)
        {
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "ID");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Title");

            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Actions");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Area");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Attachments");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "DateCreated");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "DateModified");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Error");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Exists");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Implementation");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "IsAutomated");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "IsDirty");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Links");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "OwnerName");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Priority");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "TeamProjectName");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Reason");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Revision");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "State");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "WorkItemId");

            insertAt.IncrementRows();
        }

        internal static void Add_TestVariables(Excel.XlLocation insertAt)
        {
            XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 20, "TeamProject");

            XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 20, "ID");
            XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 20, "AllowedValues");
            XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 20, "Description");
            XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 20, "Name");
            XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 20, "Revision");

            insertAt.IncrementRows();
        }

        #endregion
    }
}
