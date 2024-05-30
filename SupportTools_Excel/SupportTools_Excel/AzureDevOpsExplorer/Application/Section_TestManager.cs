using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;

using Microsoft.Office.Interop.Excel;
using Microsoft.TeamFoundation.TestManagement.Client;

using SupportTools_Excel.AzureDevOpsExplorer.Domain;

using VNC.AddinHelper;

using XlHlp = VNC.AddinHelper.Excel;

namespace SupportTools_Excel.AzureDevOpsExplorer.Application
{
    class Section_TestManager
    {
        #region TestManager (TM)

        public delegate int ProcessAddBodyCommand_TM(
            XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options,
            ITestManagementTeamProject testManagementTeamProject);

        internal static XlHlp.XlLocation AddSections_TM(
            XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options,
            ITestManagementTeamProject testManagementTeamProject,
            List<string> sectionsToDisplay)
        {
            long startTicks = XlHlp.DisplayInWatchWindow(insertAt);

            //CellFormatSpecification contentFormatSpecfication = new CellFormatSpecification();
            //contentFormatSpecification.Font.Size = 14;
            //contentFormatSpecification.Font.Bold = true;

            if (sectionsToDisplay.Count != 0)
            {
                try
                {
                    if (insertAt.OrientVertical)
                    {
                        XlHlp.AddSectionInfo(insertAt.AddRow(), "Test Manager Information", "");
                        insertAt.IncrementRows();
                    }
                    else
                    {
                        //contentFormatSpecification.Orientation = XlOrientation.xlUpward;

                        XlHlp.AddSectionInfo(insertAt.AddRow(), "Test Manager Information", "",
                            orientation: XlOrientation.xlHorizontal);
                        insertAt.DecrementRows();   // AddRow bumped it.
                        insertAt.IncrementColumns();
                    }

                    if (sectionsToDisplay.Contains("Info"))
                    {
                        insertAt = Add_Info(insertAt, options, testManagementTeamProject).IncrementPosition(insertAt.OrientVertical);
                    }

                    if (sectionsToDisplay.Contains("Queries"))
                    {
                        insertAt = Add_Queries(insertAt, options, testManagementTeamProject).IncrementPosition(insertAt.OrientVertical);
                    }

                    // Plans contains Suites contain Cases

                    if (sectionsToDisplay.Contains("TestPlans"))
                    {
                        insertAt = Add_TestPlans(insertAt, options, testManagementTeamProject).IncrementPosition(insertAt.OrientVertical);
                    }

                    if (sectionsToDisplay.Contains("TestSuites"))
                    {
                        insertAt = Add_TestSuites(insertAt, options, testManagementTeamProject).IncrementPosition(insertAt.OrientVertical);
                    }

                    if (sectionsToDisplay.Contains("TestCases"))
                    {
                        insertAt = Add_TestCases(insertAt, options, testManagementTeamProject).IncrementPosition(insertAt.OrientVertical);
                    }

                    // TODO(crhodes)
                    // Decide if there is a sensible order to these

                    if (sectionsToDisplay.Contains("TestConfigurations"))
                    {
                        insertAt = Add_TestConfigurations(insertAt, options, testManagementTeamProject).IncrementPosition(insertAt.OrientVertical);
                    }

                    if (sectionsToDisplay.Contains("TestEnvironments"))
                    {
                        insertAt = Add_TestEnvironments(insertAt, options, testManagementTeamProject).IncrementPosition(insertAt.OrientVertical);
                    }

                    if (sectionsToDisplay.Contains("TestFailureTypes"))
                    {
                        insertAt = Add_TestFailureTypes(insertAt, options, testManagementTeamProject).IncrementPosition(insertAt.OrientVertical);
                    }

                    if (sectionsToDisplay.Contains("TestPoints"))
                    {
                        insertAt = Add_TestPoints(insertAt, options, testManagementTeamProject).IncrementPosition(insertAt.OrientVertical);
                    }

                    if (sectionsToDisplay.Contains("TestResolutionStates"))
                    {
                        insertAt = Add_TestResolutionStates(insertAt, options, testManagementTeamProject).IncrementPosition(insertAt.OrientVertical);
                    }

                    if (sectionsToDisplay.Contains("TestResults"))
                    {
                        insertAt = Add_TestResults(insertAt, options, testManagementTeamProject);
                    }

                    if (sectionsToDisplay.Contains("TestRuns"))
                    {
                        insertAt = Add_TestRuns(insertAt, options, testManagementTeamProject).IncrementPosition(insertAt.OrientVertical);
                    }

                    if (sectionsToDisplay.Contains("TestSettings"))
                    {
                        insertAt = Add_TestSettings(insertAt, options, testManagementTeamProject).IncrementPosition(insertAt.OrientVertical);
                    }

                    if (sectionsToDisplay.Contains("TestVariables"))
                    {
                        insertAt = Add_TestVariables(insertAt, options, testManagementTeamProject);
                    }


                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }

            XlHlp.DisplayInWatchWindow(insertAt, startTicks, "End");

            return insertAt;
        }

        internal static XlHlp.XlLocation Add_Info(
            XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options,
            ITestManagementTeamProject testManagementTeamProject)
        {
            {
                long startTicks = XlHlp.DisplayInWatchWindow(insertAt);

                try
                {
                    XlHlp.AddLabeledInfo(insertAt.AddRow(), "Test Plan Info", "Not Implemented Yet");
                    //var v0 = testManagementTeamProject.AreaRootPath; // Internal property
                    var v1 = testManagementTeamProject.AttachmentDownloadUri;
                    var v2 = testManagementTeamProject.IsValid;
                    //var v2a = testManagementTeamProject.IteratonRootPath; // Internal Property
                    var v3 = testManagementTeamProject.Queries.Count;
                    var v4 = testManagementTeamProject.ServerVersion;
                    var v4a = testManagementTeamProject.TestManagementService;
                    var v4b = testManagementTeamProject.DefaultPageSizeForTestPoints;
                    var v4c = testManagementTeamProject.FetchIdentitiesInline;
                    //var v4d = testManagementTeamProject.Sessions.Count();
                    //var v4e = testManagementTeamProject.SharedSteps.Count();

                    var v4e = testManagementTeamProject.TeamProjectName;

                    //var v4d = testManagementTeamProject.TestCases.Count();
                    //var v4d = testManagementTeamProject.TestConfigurations.Count();
                    //var v4f = testManagementTeamProject.TestEnvironments;
                    //var v4g = testManagementTeamProject.TestFailureTypes;
                    //var v4h = testManagementTeamProject.TestPlans;
                    //var v4i = testManagementTeamProject.TestPoints;
                    //var v4j = testManagementTeamProject.TestResolutionStates;
                    //var v4k = testManagementTeamProject.TestResults;
                    //var v4l = testManagementTeamProject.TestRuns;
                    //var v4m = testManagementTeamProject.TestSettings;
                    //var v4n = testManagementTeamProject.TestSuites;
                    //var v4o = testManagementTeamProject.TestVariables;
                    var v4p = testManagementTeamProject.TfsIdentityStore;


                    //var v4b = testManagementTeamProject.TestManagementServiceVersion; // ??

                    // NOTE(crhodes)
                    // This may already be covered in WIT stuff


                    //var v8 = testManagementTeamProject.WitProject.AreaRootNodes.Count;
                    //var v11 = testManagementTeamProject.WitProject.Categories.Count;
                    //var v14 = testManagementTeamProject.WitProject.HasWorkItemReadRights;
                    //var v17 = testManagementTeamProject.WitProject.HasWorkItemReadRightsRecursive;
                    //var v20 = testManagementTeamProject.WitProject.HasWorkItemWriteRights;
                    //var v23 = testManagementTeamProject.WitProject.HasWorkItemWriteRightsRecursive;
                    //var v26 = testManagementTeamProject.WitProject.Id;
                    //var v29 = testManagementTeamProject.WitProject.IterationRootNodes.Count;
                    //var v32 = testManagementTeamProject.WitProject.Store.FieldDefinitions.Count;
                    //var v35 = testManagementTeamProject.WitProject.Store.IsIdentityFieldSupported;
                    //var v35a = testManagementTeamProject.WitProject.Store.Projects.Count;
                    //var v36 = testManagementTeamProject.WitProject.Store.ServerInfo;

                    //var v32a = testManagementTeamProject.WitProject.StoredQueries.Count;
                    //var v28 = testManagementTeamProject.WitProject.WorkItemTypes.Count;

                    //var v6 = testManagementTeamProject.TestCases.Count();



                    //insertAt = Add_TP_Areas(insertAt, project);

                    //insertAt.IncrementPosition(insertAt.OrientVertical);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }

                XlHlp.DisplayInWatchWindow(insertAt, startTicks, "End");

                return insertAt;
            }
        }

        // TODO(crhodes)
        // Change this to use ProcessTestManagementSection

        internal static XlHlp.XlLocation Add_Queries(
            XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options,
            ITestManagementTeamProject testManagementTeamProject)
        {
            long startTicks = XlHlp.DisplayInWatchWindow(insertAt);

            try
            {
                int currentRows = insertAt.RowsAdded;
                // Save the location of the count so we can update later after have traversed all items.

                Range rngCount = insertAt.GetCurrentRange();

                int count = 0;

                if (insertAt.OrientVertical)
                {
                    XlHlp.AddSectionInfo(insertAt.AddRow(), "Queries", count.ToString());
                }
                else
                {
                    XlHlp.AddSectionInfo(insertAt.AddRow(), "Queries", count.ToString(), orientation: XlOrientation.xlUpward);
                    insertAt.IncrementColumns();
                }

                insertAt.MarkStart(XlHlp.MarkType.GroupTable);

                Header_TestManager.Add_Queries(insertAt);

                count = Body_TestManager.Add_Queries(insertAt, options, testManagementTeamProject);

                insertAt.MarkEnd(XlHlp.MarkType.GroupTable, string.Format("tblWorkSpaces_{0}", insertAt.workSheet.Name));

                insertAt.Group(insertAt.OrientVertical, hide: true);

                // Update 

                if (insertAt.OrientVertical)
                {
                    XlHlp.AddLabeledInfoX(rngCount, "Queries", count.ToString());
                }
                else
                {
                    XlHlp.AddLabeledInfoX(rngCount, "Queries", count.ToString(), orientation: XlOrientation.xlUpward);
                }

                if (!insertAt.OrientVertical)
                {
                    // Skip past the info just added.  Use Group end because we have indented
                    insertAt.SetLocation(insertAt.RowStart, insertAt.GroupEndColumn + 1);
                }

                insertAt.EndSectionAndSetNextLocation(insertAt.OrientVertical);

                XlHlp.DisplayInWatchWindow(insertAt, startTicks, "End");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            return insertAt;
        }

        internal static XlHlp.XlLocation Add_TestCases(
            XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options,
            ITestManagementTeamProject testManagementTeamProject)
        {
            return ProcessAddSectionTestManager(insertAt, options, testManagementTeamProject,
                "TestCases", Header_TestManager.Add_TestCases,
                (insertAt, options, testManagementTeamProject) => Body_TestManager.Add_TestCases(insertAt, options, testManagementTeamProject),
                "tblTestCases");
        }

        internal static XlHlp.XlLocation Add_TestConfigurations(
            XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options,
            ITestManagementTeamProject testManagementTeamProject)
        {
            return ProcessAddSectionTestManager(insertAt, options, testManagementTeamProject,
                "TestConfigurations", Header_TestManager.Add_TestConfigurations,
                (insertAt, options, testManagementTeamProject) => Body_TestManager.Add_TestConfigurations(insertAt, options, testManagementTeamProject),
                "tblTestConfigurations");
        }

        internal static XlHlp.XlLocation Add_TestEnvironments(
            XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options,
            ITestManagementTeamProject testManagementTeamProject)
        {
            return ProcessAddSectionTestManager(insertAt, options, testManagementTeamProject,
                "TestEnvironments", Header_TestManager.Add_TestEnvironments,
                (insertAt, options, testManagementTeamProject) => Body_TestManager.Add_TestEnvironments(insertAt, options, testManagementTeamProject),
                "tblTestEnvironments");
        }

        internal static XlHlp.XlLocation Add_TestFailureTypes(
            XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options,
            ITestManagementTeamProject testManagementTeamProject)
        {
            return ProcessAddSectionTestManager(insertAt, options, testManagementTeamProject,
                "TestFailureTypes", Header_TestManager.Add_TestFailureTypes,
                (insertAt, options, testManagementTeamProject) => Body_TestManager.Add_TestFailureTypes(insertAt, options, testManagementTeamProject),
                "tblTestFailureTypes");
        }

        internal static XlHlp.XlLocation Add_TestPlans(
            XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options,
            ITestManagementTeamProject testManagementTeamProject)
        {
            return ProcessAddSectionTestManager(insertAt, options, testManagementTeamProject,
                "TestPlans", Header_TestManager.Add_TestPlans, Body_TestManager.Add_TestPlans, "tblTestPlans");
        }

        internal static XlHlp.XlLocation Add_TestPoints(
            XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options,
            ITestManagementTeamProject testManagementTeamProject)
        {
            return ProcessAddSectionTestManager(insertAt, options, testManagementTeamProject,
                "TestPoints", Header_TestManager.Add_TestPoints,
                (insertAt, options, testManagementTeamProject) => Body_TestManager.Add_TestPoints(insertAt, options, testManagementTeamProject),
                "tblTestPoints");
        }

        internal static XlHlp.XlLocation Add_TestResolutionStates(
            XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options,
            ITestManagementTeamProject testManagementTeamProject)
        {
            return ProcessAddSectionTestManager(insertAt, options, testManagementTeamProject,
                "TestResolutionStates", Header_TestManager.Add_TestResolutionStates,
                (insertAt, options, testManagementTeamProject) => Body_TestManager.Add_TestResolutionStates(insertAt, options, testManagementTeamProject),
                "tblTestResolutionStates");
        }

        internal static XlHlp.XlLocation Add_TestResults(
            XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options,
            ITestManagementTeamProject testManagementTeamProject)
        {
            return ProcessAddSectionTestManager(insertAt, options, testManagementTeamProject,
                "TestResults", Header_TestManager.Add_TestResults,
                (insertAt, options, testManagementTeamProject) => Body_TestManager.Add_TestResults(insertAt, options, testManagementTeamProject),
                "tblTestResults");
        }

        internal static XlHlp.XlLocation Add_TestRuns(
            XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options,
            ITestManagementTeamProject testManagementTeamProject)
        {
            return ProcessAddSectionTestManager(insertAt, options, testManagementTeamProject,
                "TestRuns", Header_TestManager.Add_TestRuns,
                (insertAt, options, testManagementTeamProject) => Body_TestManager.Add_TestRuns(insertAt, options, testManagementTeamProject),
                "tblTestRuns");
        }

        internal static XlHlp.XlLocation Add_TestSettings(
            XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options,
            ITestManagementTeamProject testManagementTeamProject)
        {
            return ProcessAddSectionTestManager(insertAt, options, testManagementTeamProject,
                "TestSettings", Header_TestManager.Add_TestSettings,
                (insertAt, options, testManagementTeamProject) => Body_TestManager.Add_TestSettings(insertAt, options, testManagementTeamProject),
                "tblTestSettings");
        }

        internal static XlHlp.XlLocation Add_TestSuites(
            XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options,
            ITestManagementTeamProject testManagementTeamProject)
        {
            return ProcessAddSectionTestManager(insertAt, options, testManagementTeamProject,
                "TestSuites", Header_TestManager.Add_TestSuites, Body_TestManager.Add_TestSuites, "tblTestSuites");
        }

        internal static XlHlp.XlLocation Add_TestVariables(
            XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options,
            ITestManagementTeamProject testManagementTeamProject)
        {
            return ProcessAddSectionTestManager(insertAt, options, testManagementTeamProject,
               "TestVariables", Header_TestManager.Add_TestVariables,
               (insertAt, options, testManagementTeamProject) => Body_TestManager.Add_TestVariables(insertAt, options, testManagementTeamProject),
               "tblTestVariables");
        }

        internal static XlHlp.XlLocation Add_TestPlan_Info(
            XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options,
            ITestPlan testPlan)
        {
            long startTicks = XlHlp.DisplayInWatchWindow(insertAt);

            try
            {
                if (insertAt.OrientVertical)
                {
                    XlHlp.AddSectionInfo(insertAt.AddRow(), "TestPlan Info", $"{testPlan.Id}");
                }
                else
                {
                    XlHlp.AddSectionInfo(insertAt.AddRow(), "TestPlan Info", $"{testPlan.Id}",
                        orientation: XlOrientation.xlUpward);
                    insertAt.IncrementColumns();
                }

                CellFormatSpecification redContent = options.FormatSpecs.RedContent;
                CellFormatSpecification dateLabel = options.FormatSpecs.DateLabel;
                CellFormatSpecification dateContent = options.FormatSpecs.DateContent;

                insertAt.MarkStart(XlHlp.MarkType.Group);

                XlHlp.AddLabeledInfo(insertAt.AddRow(), "Id", $"{ testPlan.Id }", contentFormat: redContent);
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "Name", $"{ testPlan.Name }");

                XlHlp.AddLabeledInfo(insertAt.AddRow(), "AreaPath", $"{ testPlan.AreaPath }");

                XlHlp.AddLabeledInfo(insertAt.AddRow(), "AutomatedTestEnvironmentId", $"{ testPlan.AutomatedTestEnvironmentId }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "AutomatedTestSettings", $"{ testPlan.AutomatedTestSettingsId }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "BuildDirectory", $"{ testPlan.BuildDirectory }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "BuildFilter", $"{ testPlan.BuildFilter }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "BuildNumber", $"{ testPlan.BuildNumber }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "BuildTakenDate", $"{ testPlan.BuildTakenDate }", labelFormat: dateLabel, contentFormat: dateContent);

                XlHlp.AddLabeledInfo(insertAt.AddRow(), "BuildUri", $"{ testPlan.BuildUri }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "Description", $"{ testPlan.Description }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "EndDate", $"{ testPlan.EndDate }", labelFormat: dateLabel, contentFormat: dateContent);
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "Error", $"{ testPlan.Error }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "InvalidProperties", $"{ testPlan.InvalidProperties }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "IsDirty", $"{ testPlan.IsDirty }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "Iteration", $"{ testPlan.Iteration }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "LastUpdated", $"{ testPlan.LastUpdated }", labelFormat: dateLabel, contentFormat: dateContent);
                //XlHlp.AddTitledIn3fo(insertAt.AddR2ow(), "LastUpdateBy", $"{ testPlan.LastUpdatedBy }", fontSize: 10);
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "LastUpdatedByName", $"{ testPlan.LastUpdatedByName }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "Links", $"{ testPlan.Links.Count }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "ManualTestEnvironmentId", $"{ testPlan.ManualTestEnvironmentId }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "ManualTestSettingsId", $"{ testPlan.ManualTestSettingsId }");

                XlHlp.AddLabeledInfo(insertAt.AddRow(), "NewBuildStatistics.BuildCount", $"{ testPlan.NewBuildStatistics.BuildCount }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "NewBuildStatistics.CompletedRequirementCount", $"{ testPlan.NewBuildStatistics.CompletedRequirementCount }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "NewBuildStatistics.ResolvedBugCount", $"{ testPlan.NewBuildStatistics.ResolvedBugCount }");

                //XlHlp.AddTitledIn3fo(insertAt.AddR2ow(), "Owner", $"{ testPlan.Owner }", fontSize: 10);
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "OwnerName", $"{ testPlan.OwnerName }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "OwnerTeamFoundationId", $"{ testPlan.OwnerTeamFoundationId }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "PreviousBuildUri", $"{ testPlan.PreviousBuildUri }");

                XlHlp.AddLabeledInfo(insertAt.AddRow(), "TeamProjectName", $"{ testPlan.Project.TeamProjectName }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "Revision", $"{ testPlan.Revision }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "RootSuite.Title", $"{ testPlan.RootSuite.Title }");

                XlHlp.AddLabeledInfo(insertAt.AddRow(), "StartDate", $"{ testPlan.StartDate }", labelFormat: dateLabel, contentFormat: dateContent);
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "State", $"{ testPlan.State }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "UserData", $"{ testPlan.UserData}");

                string tpQuery = "SELECT *"
                    + " FROM TestPoint"
                    + " GROUP BY SuiteId, TestCaseId";

                ITestPointCollection testPoints = testPlan.QueryTestPoints(tpQuery);

                insertAt.IncrementColumns();

                XlHlp.AddLabeledInfo(insertAt.AddRow(), "TestPlan.QueryTestPoints", $"{ testPoints.Count }");

                insertAt.MarkStart(XlHlp.MarkType.GroupTable);

                Header_TestManager.Add_TestPoints(insertAt);

                var ty = testPoints[0].GetType();

                XlHlp.DisplayInWatchWindow($"Adding TestPoints");

                int totalItems = testPoints.Count;
                int itemCount = 0;

                foreach (ITestPoint testPoint in testPoints)
                {
                    insertAt.ClearOffsets();

                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ testPoint.Id }", cellFormat: redContent);
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ testPoint.AssignedToName }");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ testPoint.Comment }");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ testPoint.ConfigurationId }");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ testPoint.ConfigurationName }");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ testPoint.HasCachedProperties }");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ testPoint.History.Count }");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ testPoint.IsTestCaseAutomated }");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ testPoint.LastUpdated }", cellFormat: dateContent);
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ testPoint.LastUpdatedByName }");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ testPoint.MostRecentFailureType }");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ testPoint.MostRecentFailureTypeId }");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ testPoint.MostRecentResolutionStateId }");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ testPoint.MostRecentResult }");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ testPoint.MostRecentResultId }");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ testPoint.MostRecentResultOutcome }");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ testPoint.MostRecentResultState }");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ testPoint.MostRecentRunId }");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ testPoint.Plan.Id }", cellFormat: redContent);
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ testPoint.Revision }");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ testPoint.State }");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ testPoint.SuiteId }", cellFormat: redContent);
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ testPoint.TestCaseExists }");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ testPoint.TestCaseId }", cellFormat: redContent);
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ testPoint.TestCaseWorkItem.Id }", cellFormat: redContent);
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ testPoint.UserData }");

                    insertAt.IncrementRows();

                    itemCount++;

                    if (itemCount % options.LoopUpdateInterval == 0)
                    {
                        XlHlp.DisplayInWatchWindow($"Added {itemCount} out of {totalItems}");
                    }
                }

                insertAt.MarkEnd(XlHlp.MarkType.GroupTable, string.Format("{0}_{1}", "tblTPlan", insertAt.workSheet.Name));

                insertAt.DecrementColumns();

                insertAt.MarkEnd(XlHlp.MarkType.Group);

                insertAt.Group(insertAt.OrientVertical);

                insertAt.EndSectionAndSetNextLocation(insertAt.OrientVertical);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            XlHlp.DisplayInWatchWindow(insertAt, startTicks, "End");
            return insertAt;
        }

        internal static XlHlp.XlLocation Add_TestSuite_Info(
            XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options,
            ITestSuiteBase testSuiteBase)
        {
            long startTicks = XlHlp.DisplayInWatchWindow(insertAt);
            int totalItems;

            try
            {
                if (insertAt.OrientVertical)
                {
                    XlHlp.AddSectionInfo(insertAt.AddRow(), "TestSuite Info", $"{testSuiteBase.Id}");
                }
                else
                {
                    XlHlp.AddSectionInfo(insertAt.AddRow(), "TestPlan Info", $"{testSuiteBase.Id}",
                        orientation: XlOrientation.xlUpward);
                    insertAt.IncrementColumns();
                }

                insertAt.MarkStart(XlHlp.MarkType.Group);

                // Get or create any format specs we want to use.

                CellFormatSpecification redContent = options.FormatSpecs.RedContent;
                CellFormatSpecification dateLabel = options.FormatSpecs.DateLabel;
                CellFormatSpecification dateContent = options.FormatSpecs.DateContent;

                XlHlp.AddLabeledInfo(insertAt.AddRow(), "Id", $"{ testSuiteBase.Id }", contentFormat: redContent);
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "TestSuiteType", $"{ testSuiteBase.TestSuiteType }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "Title", $"{ testSuiteBase.Title }");

                XlHlp.AddLabeledInfo(insertAt.AddRow(), "AllTestCases", $"{ testSuiteBase.AllTestCases.Count }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "DefaultConfigurations", $"{ testSuiteBase.DefaultConfigurations.Count }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "Description", $"{ testSuiteBase.Description }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "Error", $"{ testSuiteBase.Error }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "InvalidProperties", $"{ testSuiteBase.InvalidProperties }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "IsDirty", $"{ testSuiteBase.IsDirty }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "IsRoot", $"{ testSuiteBase.IsRoot }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "LastUpdated", $"{ testSuiteBase.LastUpdated }", labelFormat: dateLabel, contentFormat: dateContent);
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "LastUpdatedByname", $"{ testSuiteBase.LastUpdatedByName }");

                string parentId = "null";

                if (testSuiteBase.Parent != null)
                {
                    parentId = testSuiteBase.Parent.Id.ToString();
                }

                XlHlp.AddLabeledInfo(insertAt.AddRow(), "Parent.Id", $"{ parentId }", contentFormat: redContent);

                XlHlp.AddLabeledInfo(insertAt.AddRow(), "Plan.Id", $"{ testSuiteBase.Plan.Id }", contentFormat: redContent);
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "Project.TeamProjectName", $"{ testSuiteBase.Project.TeamProjectName }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "Revision", $"{ testSuiteBase.Revision }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "State", $"{ testSuiteBase.State }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "Status", $"{ ((ITestSuiteBase2)testSuiteBase).Status }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "TestCaseCount", $"{ testSuiteBase.TestCaseCount }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "TestCases.Count", $"{ testSuiteBase.TestCases.Count }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "TestSuiteEntry.Configuration.Count", $"{ testSuiteBase.TestSuiteEntry.Configurations.Count }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "TestSuiteEntry.Entrytype", $"{ testSuiteBase.TestSuiteEntry.EntryType }");

                string parentTestSuiteId = "null";

                if (testSuiteBase.TestSuiteEntry.ParentTestSuite != null)
                {
                    parentTestSuiteId = testSuiteBase.TestSuiteEntry.ParentTestSuite.Id.ToString();
                }

                XlHlp.AddLabeledInfo(insertAt.AddRow(), "TestSuiteEntry.ParentTestSuite.Id", $"{ parentTestSuiteId }", contentFormat: redContent);

                switch (testSuiteBase.TestSuiteType)
                {
                    case TestSuiteType.DynamicTestSuite:
                        IDynamicTestSuite dynamicTestSuite = (IDynamicTestSuite)testSuiteBase;

                        XlHlp.AddLabeledInfo(insertAt.AddRow(), "(Dynamic)LastError", $"{ dynamicTestSuite.LastError }");
                        XlHlp.AddLabeledInfo(insertAt.AddRow(), "(Dynamic)LastPopulated", $"{ dynamicTestSuite.LastPopulated}", labelFormat: dateLabel, contentFormat: dateContent);

                        break;

                    case TestSuiteType.None:

                        break;

                    case TestSuiteType.RequirementTestSuite:
                        IRequirementTestSuite requirementTestSuite = (IRequirementTestSuite)testSuiteBase;

                        XlHlp.AddLabeledInfo(insertAt.AddRow(), "(Requirement)RequirementId", $"{ requirementTestSuite.RequirementId }", contentFormat: redContent);

                        break;

                    case TestSuiteType.StaticTestSuite:
                        IStaticTestSuite staticTestSuite = (IStaticTestSuite)testSuiteBase;

                        XlHlp.AddLabeledInfo(insertAt.AddRow(), "(StaticTestSuite)Entities", $"{ staticTestSuite.Entries.Count }");

                        Header_TestManager.Add_TestSuite_Entry(insertAt);

                        foreach (ITestSuiteEntry entry in staticTestSuite.Entries.ToList().OrderBy(x => x.Id))
                        {
                            insertAt.ClearOffsets();

                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{entry.Id}", cellFormat: redContent);
                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{entry.Title}");

                            string configurationCount = "null";
                            string pointAssignments = "null";
                            string testCaseStuff = "null";
                            string testSuiteStuff = "null";
                            string testObjectStuff = "null";

                            if (entry.Configurations != null)
                            {
                                configurationCount = entry.Configurations.Count.ToString();
                            }

                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), configurationCount);
                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{entry.EntryType}");
                            //XlHlp.AddContentToCell(insertAt.AddOffsetColumn(), entry.InvalidProperties.ToString());
                            //XlHlp.AddContentToCell(insertAt.AddOffsetColumn(), ((IPropertyOwner)entry).IsDirty.ToString());
                            //XlHlp.AddContentToCell(insertAt.AddOffsetColumn(), entry.Par

                            if (entry.PointAssignments != null)
                            {
                                pointAssignments = $"{entry.PointAssignments}";
                            }

                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), pointAssignments);

                            if (entry.TestCase != null)
                            {
                                testCaseStuff = $"{entry.TestCase}";
                            }

                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), testCaseStuff);

                            if (entry.TestObject != null)
                            {
                                testObjectStuff = $"{entry.TestObject}";
                            }

                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), testObjectStuff);

                            if (entry.TestSuite != null)
                            {
                                testSuiteStuff = $"{entry.TestSuite}";
                            }

                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), testSuiteStuff);

                            insertAt.IncrementRows();
                        }

                        XlHlp.AddLabeledInfo(insertAt.AddRow(), "(Static)EntriesException", $"{ staticTestSuite.EntriesException }");
                        XlHlp.AddLabeledInfo(insertAt.AddRow(), "(Static)SubSuites", $"{ staticTestSuite.SubSuites.Count }");

                        insertAt.IncrementColumns();

                        Header_TestManager.Add_TestSuite_TestSuite(insertAt);

                        foreach (var testSuite in staticTestSuite.SubSuites.ToList().OrderBy(x => x.Id))
                        {
                            XlHlp.DisplayInWatchWindow($"Adding TestSuite");

                            insertAt.ClearOffsets();

                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testSuite.Id}", cellFormat: redContent);
                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testSuite.Title}");


                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testSuite.AllTestCases.Count}");

                            string defaultConfigurations = "null";
                            string error = "null";
                            string lastUpdatedByName = "null";

                            if (testSuite.DefaultConfigurations != null)
                            {
                                defaultConfigurations = $"{testSuite.DefaultConfigurations.Count}";
                            }
                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), defaultConfigurations);

                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), testSuite.Description);

                            if (testSuite.Error != null)
                            {
                                error = testSuite.Error;
                            }

                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), error);
                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testSuite.InvalidProperties}");
                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testSuite.IsDirty}");
                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testSuite.IsRoot}");

                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testSuite.LastUpdated}", cellFormat: dateContent);

                            if (testSuite.LastUpdatedByName != null)
                            {
                                lastUpdatedByName = testSuite.LastUpdatedByName;
                            }

                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), lastUpdatedByName);

                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testSuite.Parent.Id}", cellFormat: redContent);
                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testSuite.Plan.Id}", cellFormat: redContent);
                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testSuite.Plan.Name}");
                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testSuite.Revision}");
                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testSuite.State}");
                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testSuite.TestCaseCount}");
                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{testSuite.TestSuiteType}");

                            insertAt.IncrementRows();

                            insertAt.ClearOffsets();

                            insertAt.IncrementColumns();

                            XlHlp.AddLabeledInfo(insertAt.AddRow(), "TestSuite.AllTestCases", $"{ testSuite.AllTestCases.Count }");

                            Header_TestManager.Add_TestSuite_TestCase(insertAt);

                            string implementation = "manual";

                            int itemCount = 0;

                            totalItems = testSuite.AllTestCases.Count;

                            XlHlp.DisplayInWatchWindow($"Adding {totalItems} TestCases");

                            foreach (ITestCase testCase in testSuite.AllTestCases.ToList().OrderBy(x => x.Id))
                            {
                                insertAt.ClearOffsets();

                                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ testCase.Id }", cellFormat: redContent);
                                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ testCase.Title  }");
                                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ testCase.Actions.Count }");
                                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ testCase.Area }");
                                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ testCase.Attachments.Count }");
                                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ testCase.DateCreated }", cellFormat: dateContent);
                                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ testCase.DateModified }", cellFormat: dateContent);
                                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ testCase.Error }");
                                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ testCase.Exists }");

                                if (testCase.Implementation != null)
                                {
                                    implementation = testCase.Implementation.DisplayText;
                                }

                                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), implementation);

                                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ testCase.IsAutomated }");
                                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ testCase.IsDirty }");
                                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ testCase.Links.Count }");
                                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ testCase.OwnerName }");
                                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ testCase.Priority }");
                                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ testCase.Project.TeamProjectName }");
                                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ testCase.Reason }");
                                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ testCase.Revision }");
                                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ testCase.State }");
                                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ testCase.WorkItem.Id }", cellFormat: redContent);

                                insertAt.IncrementRows();

                                itemCount++;

                                if (itemCount % options.LoopUpdateInterval == 0)
                                {
                                    XlHlp.DisplayInWatchWindow($"Added {itemCount} out of {totalItems}");
                                }
                            }

                            insertAt.DecrementColumns();    // Test Cases
                        }

                        insertAt.DecrementColumns();    // Test Suites

                        break;
                }

                insertAt.MarkEnd(XlHlp.MarkType.Group);

                insertAt.Group(insertAt.OrientVertical);

                insertAt.EndSectionAndSetNextLocation(insertAt.OrientVertical);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            XlHlp.DisplayInWatchWindow(insertAt, startTicks, "End");
            return insertAt;
        }

        internal static XlHlp.XlLocation Add_TestCase_Info(
            XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options,
            ITestCase testCase)
        {
            long startTicks = XlHlp.DisplayInWatchWindow(insertAt);

            try
            {
                if (insertAt.OrientVertical)
                {
                    XlHlp.AddSectionInfo(insertAt.AddRow(), "TestCase Info", $"{testCase.Id}");
                }
                else
                {
                    XlHlp.AddSectionInfo(insertAt.AddRow(), "TestCase Info", $"{testCase.Id}",
                        orientation: XlOrientation.xlUpward);
                    insertAt.IncrementColumns();
                }

                insertAt.MarkStart(XlHlp.MarkType.Group);

                // TODO(crhodes)
                // Figure out what does here

                CellFormatSpecification redContent = options.FormatSpecs.RedContent;
                CellFormatSpecification dateLabel = options.FormatSpecs.DateLabel;
                CellFormatSpecification dateContent = options.FormatSpecs.DateContent;


                XlHlp.AddLabeledInfo(insertAt.AddRow(), "Id", $"{ testCase.Id }", contentFormat: redContent);
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "Title", $"{ testCase.Title }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "Actions", $"{ testCase.Actions.Count }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "Area", $"{ testCase.Area }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "Attachments", $"{ testCase.Attachments.Count }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "CustomFields", $"{ testCase.CustomFields.Count }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "Data.DataSetName", $"{ testCase.Data.DataSetName }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "DataReadOnly.DataSetName", $"{ testCase.DataReadOnly.DataSetName }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "DateCreated", $"{ testCase.DateCreated }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "DateModified", $"{ testCase.DateModified }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "DefaultTable", $"{ testCase.DefaultTable }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "DefaultTableReadOnly", $"{ testCase.DefaultTableReadOnly }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "Description", $"{ testCase.Description }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "Error", $"{ testCase.Error }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "Exists", $"{ testCase.Exists }");
                // This can be null
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "Id", $"{ testCase.Implementation }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "InvalidProperties", $"{ testCase.InvalidProperties.Count() }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "IsAutomated", $"{ testCase.IsAutomated }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "IsDirty", $"{ testCase.IsDirty }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "Links", $"{ testCase.Links.Count }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "OwnerName", $"{ testCase.OwnerName }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "Priority", $"{ testCase.Priority }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "Project.TeamProjectName", $"{ testCase.Project.TeamProjectName }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "Reason", $"{ testCase.Reason }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "Revision", $"{ testCase.Revision }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "State", $"{ testCase.State }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "TestParameters", $"{ testCase.TestParameters.Count }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "TestSuiteEntry.Id", $"{ testCase.TestSuiteEntry.Id }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "TestSuiteEntry.Title", $"{ testCase.TestSuiteEntry.Title }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "UserData", $"{ testCase.UserData }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "WorkItem.Id", $"{ testCase.WorkItem.Id }");

                insertAt.MarkEnd(XlHlp.MarkType.Group);

                insertAt.Group(insertAt.OrientVertical);

                insertAt.EndSectionAndSetNextLocation(insertAt.OrientVertical);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            XlHlp.DisplayInWatchWindow(insertAt, startTicks, "End");
            return insertAt;
        }

        #endregion

        internal static XlHlp.XlLocation ProcessAddSectionTestManager(
        XlHlp.XlLocation insertAt,
        Options_AZDO_TFS options,
        ITestManagementTeamProject testManagementTeamProject,
        string sectionTitle,
        RequestHandlers.ProcessAddHeaderCommand addHeaderCommand,
        ProcessAddBodyCommand_TM addBodyCommand,
        string tablePrefix)
        {
            long startTicks = XlHlp.DisplayInWatchWindow(insertAt);

            try
            {
                // Save the location of the count so we can update later after have traversed all items.

                Range rngCount = insertAt.GetCurrentRange();

                int count = 0;

                if (insertAt.OrientVertical)
                {
                    XlHlp.AddSectionInfo(insertAt.AddRow(), sectionTitle, "");
                }
                else
                {
                    XlHlp.AddSectionInfo(insertAt.AddRow(), sectionTitle, "", orientation: XlOrientation.xlUpward);
                    insertAt.IncrementColumns();
                }

                insertAt.MarkStart(XlHlp.MarkType.GroupTable);

                addHeaderCommand(insertAt);

                count = addBodyCommand(insertAt, options, testManagementTeamProject);

                insertAt.MarkEnd(XlHlp.MarkType.GroupTable, string.Format("{0}_{1}", tablePrefix, insertAt.workSheet.Name));

                insertAt.Group(insertAt.OrientVertical);

                if (insertAt.OrientVertical)
                {
                    XlHlp.AddLabeledInfoX(rngCount, sectionTitle, count.ToString());
                }
                else
                {
                    XlHlp.AddLabeledInfoX(rngCount, sectionTitle, count.ToString(), orientation: XlOrientation.xlUpward);
                }

                if (!insertAt.OrientVertical)
                {
                    // Skip past the info just added.  Use Group end because we have indented
                    insertAt.SetLocation(insertAt.RowStart, insertAt.GroupEndColumn + 1);
                }

                insertAt.EndSectionAndSetNextLocation(insertAt.OrientVertical);

                XlHlp.DisplayInWatchWindow(insertAt, startTicks, "End");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            XlHlp.DisplayInWatchWindow("End: " + DateTime.Now);
            XlHlp.DisplayInWatchWindow(insertAt, startTicks, "End");

            return insertAt;
        }
    }
}
