using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Controls;

using Microsoft.TeamFoundation.TestManagement.Client;
using SupportTools_Excel.Domain;
using VNC.AddinHelper;

using XlHlp = VNC.AddinHelper.Excel;
using SupportTools_Excel.AzureDevOpsExplorer.Application;
using SupportTools_Excel.AzureDevOpsExplorer.Domain;
using VNC;

namespace SupportTools_Excel.AzureDevOpsExplorer.Application
{
    public class Body_TestManager
    {
        #region TestManager (TM)

        internal static int Add_Queries(
            Excel.XlLocation insertAt,
            Options_AZDO_TFS options,
            ITestManagementTeamProject testManagementTeamProject)
        {
            Int64 startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            int itemCount = 0;
            int totalItems = testManagementTeamProject.Queries.Count;

            XlHlp.DisplayInWatchWindow($"Processing ({ totalItems } Queries");

            foreach (ITestCaseQuery query in testManagementTeamProject.Queries)
            {
                insertAt.ClearOffsets();

                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), testManagementTeamProject.TeamProjectName);

                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), query.Name);
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), query.Owner);
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), query.QueryText);

                insertAt.IncrementRows();
                itemCount++;

                AZDOHelper.ProcessItemDelay(options);
                AZDOHelper.DisplayLoopUpdates(startTicks, options, totalItems, itemCount);
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);

            return itemCount;
        }

        internal static int Add_TestCases(Excel.XlLocation insertAt,
            Options_AZDO_TFS options,
            ITestManagementTeamProject testManagementTeamProject)
        {
            Int64 startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            int count = 0;

            string query = String.Format(
                "SELECT [System.Id], [System.Title]"
                + " FROM WorkItems"
                + " WHERE [System.WorkItemType] = 'Test Case'"
                + " AND [Team Project] = '{0}'", testManagementTeamProject.TeamProjectName);

            //string query = String.Format(
            //    "SELECT *"
            //    + " FROM TestCase");

            // NOTE(\)
            // Seems like all fields get populated.

            //int countQueries = testManagementTeamProject.Queries.Count;

            foreach (var testCase in testManagementTeamProject.TestCases.Query(query))
            {
                insertAt.ClearOffsets();
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), testManagementTeamProject.TeamProjectName);

                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), testCase.Id.ToString());
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), testCase.Title);
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), testCase.Description);
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), testCase.DateCreated.ToString());
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), testCase.DateModified.ToString());
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), testCase.Revision.ToString());
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), testCase.State);
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), testCase.Links.Count.ToString());
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), testCase.TestSuiteEntry.Id.ToString());
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), testCase.TestSuiteEntry.Title);

                insertAt.IncrementRows();
                count++;
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);

            return count;
        }

        internal static int Add_TestConfigurations(Excel.XlLocation insertAt,
            Options_AZDO_TFS options,
            ITestManagementTeamProject testManagementTeamProject)
        {
            Int64 startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            int itemCount = 0;

            string query = String.Format(
                "SELECT *"
                + " FROM TestConfiguration");

            ITestConfigurationCollection testConfigurations = testManagementTeamProject.TestConfigurations.Query(query);
            int totalItems = testConfigurations.Count;

            XlHlp.DisplayInWatchWindow($"Processing ({ totalItems } testConfigurations");

            foreach (ITestConfiguration testConfiguration in testConfigurations)
            {
                insertAt.ClearOffsets();
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testManagementTeamProject.TeamProjectName}");

                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testConfiguration.Id}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testConfiguration.Name}");

                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testConfiguration.AreaPath}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testConfiguration.Description}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testConfiguration.IsDefault}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testConfiguration.LastUpdated}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testConfiguration.LastUpdatedByName}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testConfiguration.Revision}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testConfiguration.State}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testConfiguration.Values.Count}");

                insertAt.IncrementRows();
                itemCount++;

                //ProcessItemDelay(options);
                AZDOHelper.DisplayLoopUpdates(startTicks, options, totalItems, itemCount);
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);

            return itemCount;
        }

        internal static int Add_TestEnvironments(
            Excel.XlLocation insertAt,
            Options_AZDO_TFS options,
            ITestManagementTeamProject testManagementTeamProject)
        {
            Int64 startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            int itemCount = 0;

            string query = String.Format(
                "SELECT *"
                + " FROM TestEnvironment");

            var testEnvironments = testManagementTeamProject.TestEnvironments.Query();
            var totalItems = testEnvironments.Count();

            XlHlp.DisplayInWatchWindow($"Processing ({ totalItems } testEnvironments");

            foreach (ITestEnvironment testEnvironment in testEnvironments)
            {
                insertAt.ClearOffsets();
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), testManagementTeamProject.TeamProjectName);

                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), testEnvironment.Id.ToString());
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testEnvironment.Name}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testEnvironment.ControllerDisplayName}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testEnvironment.ControllerEnvironmentId}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testEnvironment.ControllerName}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testEnvironment.DateCreated}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testEnvironment.Description}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testEnvironment.DisplayName}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testEnvironment.EnvironmentType}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testEnvironment.Error}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testEnvironment.InvalidProperties.Count()}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testEnvironment.IsDirty}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testEnvironment.LabEnvironmentUri}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testEnvironment.LabServerUri}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testEnvironment.MachineRoles.Count}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testEnvironment.Owner}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testEnvironment.TeamProject.TeamProjectName}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testEnvironment.TestController.Name}");

                insertAt.IncrementRows();
                itemCount++;

                //ProcessItemDelay(options);
                AZDOHelper.DisplayLoopUpdates(startTicks, options, totalItems, itemCount);
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);

            return itemCount;
        }

        internal static int Add_TestFailureTypes(
            Excel.XlLocation insertAt,
            Options_AZDO_TFS options,
            ITestManagementTeamProject testManagementTeamProject)
        {
            Int64 startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            int itemCount = 0;

            string query = String.Format(
                "SELECT *"
                + " FROM TestFailureType");

            IEnumerable<ITestFailureType> testFailureTypes = testManagementTeamProject.TestFailureTypes.Query();
            int totalItems = testFailureTypes.Count();

            XlHlp.DisplayInWatchWindow($"Processing ({ totalItems } testFailureTypes");

            foreach (ITestFailureType testFailureType in testFailureTypes)
            {
                insertAt.ClearOffsets();

                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testManagementTeamProject.TeamProjectName}");

                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testFailureType.Id}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testFailureType.Name}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testFailureType.Project.TeamProjectName}");

                insertAt.IncrementRows();
                itemCount++;

                //ProcessItemDelay(options);
                AZDOHelper.DisplayLoopUpdates(startTicks, options, totalItems, itemCount);
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);

            return itemCount;
        }

        internal static int Add_TestPlans(
            Excel.XlLocation insertAt,
            Options_AZDO_TFS options,
            ITestManagementTeamProject testManagementTeamProject)
        {
            Int64 startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            int itemCount = 0;

            //string query = String.Format(
            //    "SELECT [System.Id]"
            //    + " FROM TestPlan"
            //    + " WHERE [Team Project] = '{0}'", testManagementTeamProject.TeamProjectName);

            string query = String.Format(
                "SELECT *"
                + " FROM TestPlan");

            ITestPlanCollection testPlans = testManagementTeamProject.TestPlans.Query(query);
            int totalItems = testPlans.Count;

            XlHlp.DisplayInWatchWindow($"Processing ({ totalItems } testPlans");

            foreach (var testPlan in testPlans)
            {
                insertAt.ClearOffsets();

                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testManagementTeamProject.TeamProjectName}");

                try
                {
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testPlan.Id}");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testPlan.Name}");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testPlan.Description}");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testPlan.State}");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testPlan.AreaPath}");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testPlan.Iteration}");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testPlan.StartDate}");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testPlan.EndDate}");

                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testPlan.OwnerName}");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testPlan.LastUpdated}");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testPlan.LastUpdatedByName}");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testPlan.Revision}");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testPlan.Links.Count}");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testPlan.RootSuite.Id}");
                }
                catch (Exception ex)
                {
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ex}");
                }

                insertAt.IncrementRows();
                itemCount++;

                AZDOHelper.ProcessItemDelay(options);
                AZDOHelper.DisplayLoopUpdates(startTicks, options, totalItems, itemCount);
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);

            return itemCount;
        }

        internal static int Add_TestPoints(
            Excel.XlLocation insertAt,
            Options_AZDO_TFS options,
            ITestManagementTeamProject testManagementTeamProject)
        {
            Int64 startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            int itemCount = 0;

            string query = String.Format(
                "SELECT *"
                + " FROM TestPoint");

            XlHlp.AddLabeledInfo(insertAt.AddRow(), "Not Implemented", "Yet");

            //// TODO(crhodes)
            //// Learn what TestPoints are.  testManagementTeamProject.TestPoints returns a helper with methods

            //ITestPointHelper testPoints = testManagementTeamProject.TestPoints;
            //testPoints.

            //foreach (ITestPoint testPoint in testManagementTeamProject.TestPoints)
            //{
            //    insertAt.ClearOffsets();

            //    XlHlp.AddContentToCell(insertAt.AddOffsetColumn(), $"{testManagementTeamProject.TeamProjectName}");

            //    XlHlp.AddContentToCell(insertAt.AddOffsetColumn(), $"{testFailureType.Id}");
            //    XlHlp.AddContentToCell(insertAt.AddOffsetColumn(), $"{testFailureType.Name}");
            //    XlHlp.AddContentToCell(insertAt.AddOffsetColumn(), $"{testFailureType.Project.TeamProjectName}");

            //    insertAt.IncrementRows();
            //    itemCount++;
            //    ProcessLoopDelay(options);
            //    DisplayLoopUpdates(options, totalItems, itemCount);
            //}

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);

            return itemCount;
        }

        internal static int Add_TestResolutionStates(
            Excel.XlLocation insertAt,
            Options_AZDO_TFS options,
            ITestManagementTeamProject testManagementTeamProject)
        {
            Int64 startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            int itemCount = 0;

            IEnumerable<ITestResolutionState> testResolutionStates = testManagementTeamProject.TestResolutionStates.Query();
            int totalItems = testResolutionStates.Count();

            XlHlp.DisplayInWatchWindow($"Processing ({ totalItems } testResolutionStates");

            foreach (ITestResolutionState testResolutionState in testManagementTeamProject.TestResolutionStates.Query())
            {
                insertAt.ClearOffsets();

                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), testManagementTeamProject.TeamProjectName);

                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testResolutionState.Id}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testResolutionState.Name}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testResolutionState.Project.TeamProjectName}");

                insertAt.IncrementRows();
                itemCount++;

                //ProcessItemDelay(options);
                AZDOHelper.DisplayLoopUpdates(startTicks, options, totalItems, itemCount);
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);

            return itemCount;
        }

        internal static int Add_TestResults(
            Excel.XlLocation insertAt,
            Options_AZDO_TFS options,
            ITestManagementTeamProject testManagementTeamProject)
        {
            Int64 startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            int itemCount = 0;

            string query = String.Format(
                "SELECT *"
                + " FROM TestResult"
                + " WHERE [DateCreated] > '1/1/2021'");

            ITestCaseResultCollection testCaseResults = testManagementTeamProject.TestResults.Query(query);
            int totalItems = testCaseResults.Count;

            int[] associatedWIs = testCaseResults.QueryAssociatedWorkItems();

            XlHlp.DisplayInWatchWindow($"Processing ({ totalItems }) testCaseResults");

            foreach (ITestResult testResult in testCaseResults)
            {
                insertAt.ClearOffsets();

                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), testManagementTeamProject.TeamProjectName);
                
                // IAttachmentOwner

                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testResult.Attachments.Count}");

                // ITestResult
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testResult.CollectorsEnabled.Count}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testResult.Comment}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testResult.DateCreated}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testResult.DateStarted}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testResult.DateCompleted}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testResult.Duration}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testResult.ErrorMessage}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testResult.Outcome}");

                insertAt.IncrementRows();
                itemCount++;

                AZDOHelper.ProcessItemDelay(options);
                AZDOHelper.DisplayLoopUpdates(startTicks, options, totalItems, itemCount);
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);

            return itemCount;
        }

        internal static int Add_TestRuns(
            Excel.XlLocation insertAt,
            Options_AZDO_TFS options,
            ITestManagementTeamProject testManagementTeamProject)
        {
            Int64 startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            int itemCount = 0;

            string query = String.Format(
                "SELECT *"
                + " FROM TestRun");
                //+ " WHERE [DateCreated] > '1/1/2021'");

            IEnumerable<ITestRun> testRuns = testManagementTeamProject.TestRuns.Query(query);
            int totalItems = testRuns.Count();

            XlHlp.DisplayInWatchWindow($"Processing ({ totalItems } testRuns");

            foreach (ITestRun testRun in testRuns)
            {
                insertAt.ClearOffsets();

                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), testManagementTeamProject.TeamProjectName);

                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testRun.Title}");

                // IAttachmentOwner

                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testRun.Attachments.Count}");

                // ITestRun

                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testRun.BuildConfigurationId}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testRun.BuildDirectory}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testRun.BuildFlavor}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testRun.BuildNumber}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testRun.BuildPlatform}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testRun.BuildUri}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testRun.IsBvt}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testRun.Iteration}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testRun.LastUpdated}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testRun.LastUpdatedByName}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testRun.LinkedWorkItemCount}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testRun.NotApplicableTests}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testRun.OwnerName}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testRun.PassedTests}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testRun.PostProcessState}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testRun.Project.TeamProjectName}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testRun.Revision}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testRun.State}");

                try
                {
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testRun.Statistics.CompletedTests}");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testRun.Statistics.FailedTests}");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testRun.Statistics.InconclusiveTests}");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testRun.Statistics.InProgressTests}");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testRun.Statistics.PassedTests}");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testRun.Statistics.PendingTests}");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testRun.Statistics.TotalTests}");
                }
                catch (Exception ex)
                {
                    var message = ex.ToString();
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), "<Exception>");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), "<Exception>");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), "<Exception>");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), "<Exception>");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), "<Exception>");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), "<Exception>");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), "<Exception>");
                }

                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testRun.TestEnvironmentId}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testRun.TestMessageLogEntries.Count}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testRun.TestPlanId}");

                string testSettingsName = "<null>";
                string testSettingsId = "<null>";

                if (testRun.TestSettings != null)
                {
                    testSettingsName = $"{testRun.TestSettings.Name}";
                    testSettingsId = $"{testRun.TestSettings.Id}";
                }

                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), testSettingsName);
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), testSettingsId);

                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testRun.TestSettingsId}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testRun.TotalTests}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testRun.Type}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testRun.UnanalyzedTests}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testRun.Version}");

                insertAt.IncrementRows();
                itemCount++;

                AZDOHelper.ProcessItemDelay(options);
                AZDOHelper.DisplayLoopUpdates(startTicks, options, totalItems, itemCount);
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);

            return itemCount;
        }

        internal static int Add_TestSettings(
            Excel.XlLocation insertAt,
            Options_AZDO_TFS options,
            ITestManagementTeamProject testManagementTeamProject)
        {
            Int64 startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            int itemCount = 0;

            string query = String.Format(
                "SELECT *"
                + " FROM TestSettings");

            IEnumerable<ITestSettings> testSettings = testManagementTeamProject.TestSettings.Query(query);
            int totalItems = testSettings.Count();

            XlHlp.DisplayInWatchWindow($"Processing ({ totalItems } testRuns");

            foreach (ITestSettings testSetting in testSettings)
            {
                insertAt.ClearOffsets();

                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testManagementTeamProject.TeamProjectName}");

                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testSetting.Id}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testSetting.IsAutomated}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testSetting.LastUpdated}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testSetting.LastUpdatedBy}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testSetting.MachineRoles.Count}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testSetting.Name}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testSetting.Revision}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testSetting.Settings}");

                insertAt.IncrementRows();
                itemCount++;

                AZDOHelper.ProcessItemDelay(options);
                AZDOHelper.DisplayLoopUpdates(startTicks, options, totalItems, itemCount);
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);

            return itemCount;
        }

        internal static int Add_TestSuites(
            Excel.XlLocation insertAt,
            Options_AZDO_TFS options,
            ITestManagementTeamProject testManagementTeamProject)
        {
            Int64 startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            int itemCount = 0;

            //string query = String.Format(
            //    "SELECT [System.Id], [System.Title]"
            //    + " FROM WorkItems"
            //    + " WHERE [System.WorkItemType] = 'Test Suite'"
            //    + " AND [Team Project] = '{0}'", testManagementTeamProject.TeamProjectName);

            string query = String.Format(
                "SELECT *"
                + " FROM TestSuite"
            );

            ITestSuiteCollection testSuites = testManagementTeamProject.TestSuites.Query(query);
            int totalItems = testSuites.Count;

            XlHlp.DisplayInWatchWindow($"Processing ({ totalItems } testSuites");

            foreach (var testSuite in testSuites)
            {
                insertAt.ClearOffsets();

                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testManagementTeamProject.TeamProjectName}");

                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testSuite.Id}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testSuite.Title}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testSuite.Description}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testSuite.State}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testSuite.LastUpdated}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testSuite.LastUpdatedByName}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testSuite.TestCaseCount}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testSuite.TestSuiteType}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testSuite.AllTestCases.Count}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testSuite.Revision}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testSuite.Plan.Id}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testSuite.Plan.Name}");

                insertAt.IncrementRows();
                itemCount++;

                AZDOHelper.ProcessItemDelay(options);
                AZDOHelper.DisplayLoopUpdates(startTicks, options, totalItems, itemCount);
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);

            return itemCount;
        }

        internal static int Add_TestVariables(
            Excel.XlLocation insertAt,
            Options_AZDO_TFS options,
            ITestManagementTeamProject testManagementTeamProject)
        {
            Int64 startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            int itemCount = 0;

            string query = String.Format(
                "SELECT *"
                + " FROM TestVariable");

            IEnumerable<ITestVariable> testVariables = testManagementTeamProject.TestVariables.Query();
            int totalItems = testVariables.Count();

            XlHlp.DisplayInWatchWindow($"Processing ({ totalItems } testVariables");

            foreach (ITestVariable testVariable in testVariables)
            {
                insertAt.ClearOffsets();

                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testManagementTeamProject.TeamProjectName}");

                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testVariable.Id}");

                StringBuilder allowedValues = new StringBuilder();

                foreach (var item in testVariable.AllowedValues)
                {
                    if (allowedValues.Length == 0)
                    {
                        allowedValues.Append($"{item.Value}");
                    }
                    else
                    {
                        allowedValues.Append($"; {item.Value}");
                    }
                }

                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{allowedValues}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testVariable.Description}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testVariable.Name}");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testVariable.Revision}");

                insertAt.IncrementRows();
                itemCount++;

                AZDOHelper.ProcessItemDelay(options);
                AZDOHelper.DisplayLoopUpdates(startTicks, options, totalItems, itemCount);
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);

            return itemCount;
        }

        #endregion
    }
}
