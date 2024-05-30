using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;

using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.Framework.Client;
using Microsoft.TeamFoundation.Framework.Common;
using Microsoft.TeamFoundation.Server;
using Microsoft.TeamFoundation.TestManagement.Client;
using Microsoft.TeamFoundation.VersionControl.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;

using SupportTools_Excel.AzureDevOpsExplorer.Application;
using SupportTools_Excel.AzureDevOpsExplorer.Domain;

using VNC;

using VNCTFS = VNC.TFS;
using XlHlp = VNC.AddinHelper.Excel;

namespace SupportTools_Excel.User_Interface.User_Controls
{
    public partial class wucTaskPane_TFS : UserControl
    {
        #region CreateWS_*

        #region ConfigurationServer

        private void CreateWS_ConfigurationServer_Info(
            Options_AZDO_TFS options,
            TfsConfigurationServer configurationServer)
        {
            long startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            XlHlp.XlLocation insertAt = CreateNewWorksheet($"CS_{configurationServer.Name}", options);

            XlHlp.AddLabeledInfoX(insertAt.AddRowX(2), "Configuration Server Info:", configurationServer.Name);

            insertAt = Section_ConfigurationServer.AddSection_OperationalDatabaseNames(insertAt);

            insertAt = Section_ConfigurationServer.AddSection_ConfigurationServer_Info(insertAt, AzureDevOpsExplorer.Presentation.Views.Server.ConfigurationServer);

            //startingRow++;

            //// TODO: Decide if want to display something about a Catalog Node

            // null = all types

            //ReadOnlyCollection<CatalogNode> childNodes =
            //    Server.ConfigurationServer.CatalogNode.QueryChildren(null, false, CatalogQueryOptions.None);

            // None of these returned any nodes except for  ProjectCollection.

            //ReadOnlyCollection<CatalogNode> childNodes =
            //    Server.ConfigurationServer.CatalogNode.QueryChildren(
            //        new[] { CatalogResourceTypes.AnalysisDatabase }, false, CatalogQueryOptions.None);

            ////AddSection_ChildNodes(insertAt.AddRow(), "AnalysisDatabase", childNodes);

            //// Obsolete
            ////childNodes =
            ////    configurationServer.CatalogNode.QueryChildren(
            ////        new[] { CatalogResourceTypes.ApplicationDatabase }, false, CatalogQueryOptions.None);

            ////startingRow += AddSection_ChildNodes(insertAt, "ApplicationDatabase", childNodes);

            //childNodes =
            //    Server.ConfigurationServer.CatalogNode.QueryChildren(
            //        new[] { CatalogResourceTypes.DataCollector }, false, CatalogQueryOptions.None);

            ////startingRow += AddSection_ChildNodes(insertAt, "DataCollector", childNodes);

            //childNodes =
            //    Server.ConfigurationServer.CatalogNode.QueryChildren(
            //        new[] { CatalogResourceTypes.GenericLink }, false, CatalogQueryOptions.None);

            ////startingRow += AddSection_ChildNodes(insertAt, "GenericLink", childNodes);

            //childNodes =
            //    Server.ConfigurationServer.CatalogNode.QueryChildren(
            //        new[] { CatalogResourceTypes.InfrastructureRoot }, false, CatalogQueryOptions.None);

            ////startingRow += AddSection_ChildNodes(insertAt, "InfrastructureRoot", childNodes);

            //childNodes =
            //    Server.ConfigurationServer.CatalogNode.QueryChildren(
            //        new[] { CatalogResourceTypes.Machine }, false, CatalogQueryOptions.None);

            ////startingRow += AddSection_ChildNodes(insertAt, "Machine", childNodes);

            //childNodes =
            //    Server.ConfigurationServer.CatalogNode.QueryChildren(
            //        new[] { CatalogResourceTypes.OrganizationalRoot }, false, CatalogQueryOptions.None);

            ////startingRow += AddSection_ChildNodes(insertAt, "OrganizationalRoot", childNodes);

            //childNodes =
            //    Server.ConfigurationServer.CatalogNode.QueryChildren(
            //        new[] { CatalogResourceTypes.ProcessGuidanceSite }, false, CatalogQueryOptions.None);

            ////startingRow += AddSection_ChildNodes(insertAt, "ProcessGuidanceSite", childNodes);

            //childNodes =
            //    Server.ConfigurationServer.CatalogNode.QueryChildren(new[] { CatalogResourceTypes.ProjectCollection }, false, CatalogQueryOptions.None);

            ////startingRow += AddSection_ChildNodes(insertAt, "ProjectCollection", childNodes);

            ////childNodes =
            ////     configurationServer.CatalogNode.QueryChildren(
            ////         new[] { CatalogResourceTypes.ProjectCollectionDatabase }, false, CatalogQueryOptions.None);

            ////startingRow += AddSection_ChildNodes(insertAt, "ProjectCollectionDatabase", childNodes);

            //childNodes =
            //     Server.ConfigurationServer.CatalogNode.QueryChildren(
            //         new[] { CatalogResourceTypes.ProjectPortal }, false, CatalogQueryOptions.None);

            ////startingRow += AddSection_ChildNodes(insertAt, "ProjectPortal", childNodes);

            //childNodes =
            //     Server.ConfigurationServer.CatalogNode.QueryChildren(
            //         new[] { CatalogResourceTypes.ProjectServerMapping }, false, CatalogQueryOptions.None);

            ////startingRow += AddSection_ChildNodes(insertAt, "ProjectServerMapping", childNodes);

            //childNodes =
            //     Server.ConfigurationServer.CatalogNode.QueryChildren(
            //         new[] { CatalogResourceTypes.ProjectServerRegistration }, false, CatalogQueryOptions.None);

            ////startingRow += AddSection_ChildNodes(insertAt, "ProjectServerRegistration", childNodes);

            //childNodes =
            //     Server.ConfigurationServer.CatalogNode.QueryChildren(
            //         new[] { CatalogResourceTypes.ReportingConfiguration }, false, CatalogQueryOptions.None);

            ////startingRow += AddSection_ChildNodes(insertAt, "ReportingConfiguration", childNodes);

            //childNodes =
            //     Server.ConfigurationServer.CatalogNode.QueryChildren(
            //         new[] { CatalogResourceTypes.ReportingFolder }, false, CatalogQueryOptions.None);

            ////startingRow += AddSection_ChildNodes(insertAt, "ReportingFolder", childNodes);

            //childNodes =
            //     Server.ConfigurationServer.CatalogNode.QueryChildren(
            //         new[] { CatalogResourceTypes.ReportingServer }, false, CatalogQueryOptions.None);

            ////startingRow += AddSection_ChildNodes(insertAt, "ReportingServer", childNodes);

            //childNodes =
            //     Server.ConfigurationServer.CatalogNode.QueryChildren(
            //         new[] { CatalogResourceTypes.ResourceFolder }, false, CatalogQueryOptions.None);

            ////startingRow += AddSection_ChildNodes(insertAt, "ResourceFolder", childNodes);

            ////childNodes =
            ////     configurationServer.CatalogNode.QueryChildren(
            ////         new[] { CatalogResourceTypes.SharePointSiteCreationXlLocation }, false, CatalogQueryOptions.None);

            ////startingRow += AddSection_ChildNodes(insertAt, "SharePointSiteCreationXlLocation", childNodes);

            //childNodes =
            //     Server.ConfigurationServer.CatalogNode.QueryChildren(new[] { CatalogResourceTypes.SharePointWebApplication }, false, CatalogQueryOptions.None);

            ////startingRow += AddSection_ChildNodes(insertAt, "SharePointWebApplication", childNodes);

            //childNodes =
            //     Server.ConfigurationServer.CatalogNode.QueryChildren(
            //         new[] { CatalogResourceTypes.SqlAnalysisInstance }, false, CatalogQueryOptions.None);

            ////startingRow += AddSection_ChildNodes(insertAt, "SqlAnalysisInstance", childNodes);

            ////childNodes =
            ////     configurationServer.CatalogNode.QueryChildren(
            ////         new[] { CatalogResourceTypes.SqlDatabaseInstance }, false, CatalogQueryOptions.None);

            ////startingRow += AddSection_ChildNodes(insertAt, "SqlDatabaseInstance", childNodes);

            //childNodes =
            //     Server.ConfigurationServer.CatalogNode.QueryChildren(
            //         new[] { CatalogResourceTypes.SqlReportingInstance }, false, CatalogQueryOptions.None);

            ////startingRow += AddSection_ChildNodes(insertAt, "SqlReportingInstance", childNodes);

            //childNodes =
            //     Server.ConfigurationServer.CatalogNode.QueryChildren(
            //         new[] { CatalogResourceTypes.TeamFoundationServerInstance }, false, CatalogQueryOptions.None);

            ////startingRow += AddSection_ChildNodes(insertAt, "TeamFoundationServerInstance", childNodes);

            //childNodes =
            //     Server.ConfigurationServer.CatalogNode.QueryChildren(
            //         new[] { CatalogResourceTypes.TeamFoundationWebApplication }, false, CatalogQueryOptions.None);

            ////startingRow += AddSection_ChildNodes(insertAt, "TeamFoundationWebApplication", childNodes);

            //childNodes =
            //     Server.ConfigurationServer.CatalogNode.QueryChildren(
            //         new[] { CatalogResourceTypes.TeamProject }, false, CatalogQueryOptions.None);

            ////startingRow += AddSection_ChildNodes(insertAt, "TeamProject", childNodes);

            //childNodes =
            //     Server.ConfigurationServer.CatalogNode.QueryChildren(
            //         new[] { CatalogResourceTypes.TeamSystemWebAccess }, false, CatalogQueryOptions.None);

            ////startingRow += AddSection_ChildNodes(insertAt, "TeamSystemWebAccess", childNodes);

            //childNodes =
            //     Server.ConfigurationServer.CatalogNode.QueryChildren(
            //         new[] { CatalogResourceTypes.TestController }, false, CatalogQueryOptions.None);

            ////startingRow += AddSection_ChildNodes(insertAt, "TestController", childNodes);

            //childNodes =
            //     Server.ConfigurationServer.CatalogNode.QueryChildren(
            //         new[] { CatalogResourceTypes.TestEnvironment }, false, CatalogQueryOptions.None);

            ////startingRow += AddSection_ChildNodes(insertAt, "TestEnvironment", childNodes);

            //childNodes =
            //     Server.ConfigurationServer.CatalogNode.QueryChildren(
            //         new[] { CatalogResourceTypes.WarehouseDatabase }, false, CatalogQueryOptions.None);

            ////startingRow += AddSection_ChildNodes(insertAt, "WarehouseDatabase", childNodes);

            // Get the Team Project Collections

            ReadOnlyCollection<CatalogNode> projectCollectionNodes = VNCTFS.Helper.Get_TeamProjectCollectionNodes(AzureDevOpsExplorer.Presentation.Views.Server.ConfigurationServer);

            // Add a sheet for each TeamProjectCollection

            foreach (CatalogNode teamProjectCollectionNode in projectCollectionNodes)
            {
                //TfsTeamProjectCollection teamProjectCollection = VNCTFS.Helper.Get_TeamProjectCollection(configurationServer, teamProjectCollectionNode);

                //CreateWS_TPC_Info(teamProjectCollectionNode, teamProjectCollection, false, orientVertical);
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion ConfigurationServer

        #region All_TPC

        private void CreateWS_All_TPC_AreaCheck(string areasToCheck, Options_AZDO_TFS options)
        {
            long startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            try
            {
                XlHlp.XlLocation insertAt = CreateNewWorksheet(string.Format("{0}_{1}", "All_WIS", "AreaCheck"),
                    options);

                foreach (Project project in AzureDevOpsExplorer.Presentation.Views.Server.WorkItemStore.Projects)
                {
                    long loopTicks = Log.Trace($"Processing {project.Name}", Common.LOG_CATEGORY);

                    Globals.ThisAddIn.Application.StatusBar = $"Processing {project.Name}";

                    insertAt = Section_WorkItemStore.Add_TP_AreaCheck(insertAt, project, areasToCheck);

                    Log.Trace($"EndProcessing {project.Name}", Common.LOG_CATEGORY, loopTicks);

                    AZDOHelper.ProcessLoopDelay(options);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void CreateWS_All_TPC_Areas(Options_AZDO_TFS options)
        {
            long startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            try
            {
                XlHlp.XlLocation insertAt = CreateNewWorksheet(string.Format("{0}_{1}", "All_TPC", "Areas"),
                    options);

                XlHlp.AddLabeledInfoX(insertAt.AddRowX(), "Areas All Team Projects", AzureDevOpsExplorer.Presentation.Views.Server.TfsTeamProjectCollection.Name);

                insertAt.MarkStart(XlHlp.MarkType.GroupTable);

                Header_WorkItemStore.Add_TP_Areas(insertAt);

                foreach (var teamProjectName in options.TeamProjects)
                {
                    long loopTicks = Log.Trace($"Processing {teamProjectName}", Common.LOG_CATEGORY);

                    Globals.ThisAddIn.Application.StatusBar = $"Processing {teamProjectName}";

                    Project project = AzureDevOpsExplorer.Presentation.Views.Server.WorkItemStore.Projects[teamProjectName];

                    insertAt = Body_WorkItemStore.Add_TP_Areas(insertAt, options, AzureDevOpsExplorer.Presentation.Views.Server.CommonStructureService, project);

                    // Save with each loop so we loose less when things crash ;(

                    Globals.ThisAddIn.Application.ActiveWorkbook.Save();

                    Log.Trace($"EndProcessing {teamProjectName}", Common.LOG_CATEGORY, loopTicks);

                    AZDOHelper.ProcessLoopDelay(options);
                }

                insertAt.MarkEnd(XlHlp.MarkType.GroupTable, string.Format("tbl_{0}", insertAt.workSheet.Name));

                insertAt.Group(insertAt.OrientVertical, hide: true);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void CreateWS_All_TPC_BuildDefinitions(Options_AZDO_TFS options)
        {
            long startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            try
            {
                XlHlp.XlLocation insertAt = CreateNewWorksheet(string.Format("{0}_{1}", "All_TPC", "BuildDefinitions"),
                    options);

                XlHlp.AddLabeledInfoX(insertAt.AddRowX(), "Build Definitions - All TeamProjects", AzureDevOpsExplorer.Presentation.Views.Server.TfsTeamProjectCollection.Name);

                insertAt.MarkStart(XlHlp.MarkType.GroupTable);

                Header_BuildServer.Add_BuildDefinitions(insertAt);

                foreach (var teamProjectName in options.TeamProjects)
                {
                    long loopTicks = Log.Trace($"Processing {teamProjectName}", Common.LOG_CATEGORY);

                    Globals.ThisAddIn.Application.StatusBar = $"Processing {teamProjectName}";

                    Body_BuildServer.Add_BuildDefinitions(insertAt, options, AzureDevOpsExplorer.Presentation.Views.Server.BuildServer, teamProjectName);

                    Log.Trace($"EndProcessing {teamProjectName}", Common.LOG_CATEGORY, loopTicks);

                    AZDOHelper.ProcessLoopDelay(options);
                }

                insertAt.MarkEnd(XlHlp.MarkType.GroupTable, string.Format("tbl_{0}", insertAt.workSheet.Name));

                insertAt.Group(insertAt.OrientVertical, hide: true);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void CreateWS_All_TPC_Developers(Options_AZDO_TFS options)
        {
            long startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            try
            {
                XlHlp.XlLocation insertAt = CreateNewWorksheet(string.Format("{0}_{1}({2})", "All_TPC", "Devs", options.GoBackDays),
                    options);

                XlHlp.AddLabeledInfoX(insertAt.AddRowX(), "Developers - All TeamProjects", AzureDevOpsExplorer.Presentation.Views.Server.TfsTeamProjectCollection.Name);

                insertAt.MarkStart(XlHlp.MarkType.GroupTable);

                Header_VersionControlServer.Add_TP_Developers(insertAt);

                foreach (var teamProjectName in options.TeamProjects)
                {
                    TeamProject teamProject = VNCTFS.Helper.Get_TeamProject(AzureDevOpsExplorer.Presentation.Views.Server.VersionControlServer, teamProjectName.Trim());

                    long loopTicks = Log.Trace($"Processing {teamProject.Name}", Common.LOG_CATEGORY);

                    Globals.ThisAddIn.Application.StatusBar = $"Processing {teamProject.Name}";

                    insertAt.ClearOffsets();

                    XlHlp.AddContentToCell(insertAt.AddRowX(), teamProject.Name);

                    insertAt = Section_VersionControlServer.Add_TP_Developers(insertAt, options, AzureDevOpsExplorer.Presentation.Views.Server.VersionControlServer, teamProject, true, teamProject.Name);

                    Log.Trace($"EndProcessing {teamProject.Name}", Common.LOG_CATEGORY, loopTicks);

                    AZDOHelper.ProcessLoopDelay(options);
                }

                insertAt.MarkEnd(XlHlp.MarkType.GroupTable, string.Format("tbl_{0}", insertAt.workSheet.Name));

                insertAt.Group(insertAt.OrientVertical, hide: true);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void CreateWS_All_TPC_Teams(Options_AZDO_TFS options)
        {
            long startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            try
            {
                XlHlp.XlLocation insertAt = CreateNewWorksheet(string.Format("{0}_{1}", "All_TPC", "Teams"),
                    options);

                XlHlp.AddLabeledInfoX(insertAt.AddRowX(), "Teams All Team Projects", AzureDevOpsExplorer.Presentation.Views.Server.TfsTeamProjectCollection.Name);

                insertAt.MarkStart(XlHlp.MarkType.GroupTable);

                Header_VersionControlServer.Add_TP_Teams(insertAt);

                //foreach (var teamProject in Server.VersionControlServer.GetAllTeamProjects(refresh: true))
                foreach (var teamProjectName in options.TeamProjects)
                {
                    TeamProject teamProject = VNCTFS.Helper.Get_TeamProject(AzureDevOpsExplorer.Presentation.Views.Server.VersionControlServer, teamProjectName.Trim());

                    Globals.ThisAddIn.Application.StatusBar = $"Processing  {teamProjectName}";

                    if (teamProject is null)
                    {
                        // TODO(crhodes)
                        // Add something to indicate no VersionControl stuff
                    }
                    else
                    {
                        long loopTicks = Log.Trace($"Processing    {teamProjectName}", Common.LOG_CATEGORY);

                        ProjectInfo teamProjectInfo = AzureDevOpsExplorer.Presentation.Views.Server.CommonStructureService.GetProjectFromName(teamProject.Name);
                        var tpUri = teamProjectInfo.Uri;

                        TfsTeamService teamService = new TfsTeamService();

                        teamService.Initialize(teamProject.TeamProjectCollection);

                        var defaultTeam = teamService.GetDefaultTeam(tpUri, new List<String>());

                        IEnumerable<TeamFoundationTeam> allTeams = teamService.QueryTeams(tpUri);

                        Body_VersionControlServer.Add_TP_Teams(insertAt, options, AzureDevOpsExplorer.Presentation.Views.Server.VersionControlServer, teamProject, allTeams, defaultTeam);

                        Log.Trace($"EndProcessing {teamProject.Name}", Common.LOG_CATEGORY, loopTicks);

                        AZDOHelper.ProcessLoopDelay(options);
                    }
                }

                insertAt.MarkEnd(XlHlp.MarkType.GroupTable, string.Format("tbl_{0}", insertAt.workSheet.Name));

                insertAt.Group(insertAt.OrientVertical, hide: true);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void CreateWS_All_TPC_TestCases(Options_AZDO_TFS options)
        {
            long startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            try
            {
                XlHlp.XlLocation insertAt = CreateNewWorksheet(string.Format("{0}_{1}", "All_TPC", "Test Cases"),
                    options);

                XlHlp.AddLabeledInfoX(insertAt.AddRowX(), "Test Cases - All Team Projects", AzureDevOpsExplorer.Presentation.Views.Server.TfsTeamProjectCollection.Name);

                insertAt.MarkStart(XlHlp.MarkType.GroupTable);

                Header_TestManager.Add_TestCases(insertAt);

                //foreach (var teamProject in Server.VersionControlServer.GetAllTeamProjects(refresh: true))
                foreach (var teamProjectName in options.TeamProjects)
                {
                    TeamProject teamProject = VNCTFS.Helper.Get_TeamProject(AzureDevOpsExplorer.Presentation.Views.Server.VersionControlServer, teamProjectName.Trim());

                    long loopTicks = Log.Trace($"Processing    {teamProject.Name}", Common.LOG_CATEGORY);

                    Globals.ThisAddIn.Application.StatusBar = $"Processing    {teamProject.Name}";

                    ITestManagementTeamProject testManagementTeamProject = AzureDevOpsExplorer.Presentation.Views.Server.TestManagementService.GetTeamProject(teamProject.Name);

                    Body_TestManager.Add_TestCases(insertAt, options, testManagementTeamProject);

                    Log.Trace($"EndProcessing {teamProject.Name}", Common.LOG_CATEGORY, loopTicks);

                    AZDOHelper.ProcessLoopDelay(options);
                }

                insertAt.MarkEnd(XlHlp.MarkType.GroupTable, string.Format("tbl_{0}", insertAt.workSheet.Name));

                insertAt.Group(insertAt.OrientVertical, hide: true);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void CreateWS_All_TPC_TestPlans(Options_AZDO_TFS options)
        {
            long startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            try
            {
                XlHlp.XlLocation insertAt = CreateNewWorksheet(string.Format("{0}_{1}", "All_TPC", "Test Plans"),
                    options);

                XlHlp.AddLabeledInfoX(insertAt.AddRowX(), "Test Plans - All Team Projects", AzureDevOpsExplorer.Presentation.Views.Server.TfsTeamProjectCollection.Name);

                insertAt.MarkStart(XlHlp.MarkType.GroupTable);

                Header_TestManager.Add_TestPlans(insertAt);

                //foreach (var teamProject in Server.VersionControlServer.GetAllTeamProjects(refresh: true))
                foreach (var teamProjectName in options.TeamProjects)
                {
                    TeamProject teamProject = VNCTFS.Helper.Get_TeamProject(AzureDevOpsExplorer.Presentation.Views.Server.VersionControlServer, teamProjectName.Trim());

                    long loopTicks = Log.Trace($"Processing    {teamProject.Name}", Common.LOG_CATEGORY);

                    Globals.ThisAddIn.Application.StatusBar = $"Processing    {teamProject.Name}";

                    ITestManagementTeamProject testManagementTeamProject = AzureDevOpsExplorer.Presentation.Views.Server.TestManagementService.GetTeamProject(teamProject.Name);

                    Body_TestManager.Add_TestPlans(insertAt, options, testManagementTeamProject);

                    Log.Trace($"EndProcessing {teamProject.Name}", Common.LOG_CATEGORY, loopTicks);

                    AZDOHelper.ProcessLoopDelay(options);
                }

                insertAt.MarkEnd(XlHlp.MarkType.GroupTable, string.Format("tbl_{0}", insertAt.workSheet.Name));

                insertAt.Group(insertAt.OrientVertical, hide: true);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void CreateWS_All_TPC_TestSuites(Options_AZDO_TFS options)
        {
            long startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            try
            {
                XlHlp.XlLocation insertAt = CreateNewWorksheet(string.Format("{0}_{1}", "All_TPC", "TestSuites"),
                    options);

                XlHlp.AddLabeledInfoX(insertAt.AddRowX(), "Test Suites - All Team Projects", AzureDevOpsExplorer.Presentation.Views.Server.TfsTeamProjectCollection.Name);

                insertAt.MarkStart(XlHlp.MarkType.GroupTable);

                Header_TestManager.Add_TestSuites(insertAt);

                //foreach (var teamProject in Server.VersionControlServer.GetAllTeamProjects(refresh: true))
                foreach (var teamProjectName in options.TeamProjects)
                {
                    TeamProject teamProject = VNCTFS.Helper.Get_TeamProject(AzureDevOpsExplorer.Presentation.Views.Server.VersionControlServer, teamProjectName.Trim());

                    long loopTicks = Log.Trace($"Processing    {teamProject.Name}", Common.LOG_CATEGORY);

                    Globals.ThisAddIn.Application.StatusBar = $"Processing    {teamProject.Name}";

                    ITestManagementTeamProject testManagementTeamProject = AzureDevOpsExplorer.Presentation.Views.Server.TestManagementService.GetTeamProject(teamProject.Name);

                    Body_TestManager.Add_TestSuites(insertAt, options, testManagementTeamProject);

                    Log.Trace($"EndProcessing {teamProject.Name}", Common.LOG_CATEGORY, loopTicks);

                    AZDOHelper.ProcessLoopDelay(options);
                }

                insertAt.MarkEnd(XlHlp.MarkType.GroupTable, string.Format("tbl_{0}", insertAt.workSheet.Name));

                insertAt.Group(insertAt.OrientVertical, hide: true);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void CreateWS_All_TPC_WorkItemActivity(Options_AZDO_TFS options)
        {
            long startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            try
            {
                XlHlp.XlLocation insertAt = CreateNewWorksheet("All_TPC_WorkItemActivity", options);

                XlHlp.AddLabeledInfoX(insertAt.AddRowX(), "Work Item Activity All Team Projects",
                    AzureDevOpsExplorer.Presentation.Views.Server.TfsTeamProjectCollection.Name);

                insertAt.MarkStart(XlHlp.MarkType.GroupTable);

                Header_WorkItemStore.Add_TP_WorkItemActivity(insertAt);

                foreach (var teamProjectName in options.TeamProjects)
                {
                    Project project = AzureDevOpsExplorer.Presentation.Views.Server.WorkItemStore.Projects[teamProjectName];

                    long loopTicks = Log.Trace($"Processing  {project.Name}", Common.LOG_CATEGORY);

                    Globals.ThisAddIn.Application.StatusBar = $"Processing  {project.Name}";

                    DateTime maxLastCreatedDate = DateTime.MinValue;
                    DateTime maxLastChangedDate = DateTime.MinValue;
                    DateTime maxLastRevisedDate = DateTime.MinValue;

                    Body_WorkItemStore.Add_TP_WorkItemActivity(insertAt,
                        options, AzureDevOpsExplorer.Presentation.Views.Server.WorkItemStore,
                        project,
                        out maxLastCreatedDate, out maxLastChangedDate, out maxLastRevisedDate);

                    // Save with each loop so we loose less when things crash ;(

                    Globals.ThisAddIn.Application.ActiveWorkbook.Save();

                    Log.Trace($"EndProcessing {project.Name}", Common.LOG_CATEGORY, loopTicks);

                    AZDOHelper.ProcessLoopDelay(options);
                }

                insertAt.MarkEnd(XlHlp.MarkType.GroupTable, string.Format("tbl_{0}", insertAt.workSheet.Name));

                insertAt.Group(insertAt.OrientVertical, hide: true);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void CreateWS_All_TPC_WorkItemDetails(Options_AZDO_TFS options)
        {
            long startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            try
            {
                XlHlp.XlLocation insertAt = CreateNewWorksheet("All_TPC_WorkItemDetails", options);

                XlHlp.AddLabeledInfoX(insertAt.AddRowX(), "Work Item Details All Team Projects", AzureDevOpsExplorer.Presentation.Views.Server.TfsTeamProjectCollection.Name);

                XlHlp.AddSectionInfo(insertAt.AddRow(),
                    "WorkItem Details", options.WorkItemQuerySpec.Query);

                insertAt.MarkStart(XlHlp.MarkType.GroupTable);

                WorkItemCollection queryResults;

                Header_WorkItemStore.Add_TP_WorkItemDetails(insertAt, options);

                // HACK(crhodes)
                // If the Query contains @PROJECT looping across makes sense.
                // If the Query does not contain @PROJECT
                // then it may make sense to allow the Query.ReplaceQueryTokens method
                // to add the IN (list of projects) clause to the WHERE condition.

                if (options.WorkItemQuerySpec.QueryWithTokens.Contains("PROJECT"))
                {
                    foreach (var teamProjectName in options.TeamProjects.OrderBy(n => n))
                    {
                        Project project = AzureDevOpsExplorer.Presentation.Views.Server.WorkItemStore.Projects[teamProjectName];
                        string projectName = project.Name;

                        long loopTicks = Log.Trace($"Processing {projectName}", Common.LOG_CATEGORY);

                        Globals.ThisAddIn.Application.StatusBar = $"Processing {projectName}";

                        //string parsedQuery = AZDOHelper.ParseQueryTokens(query, options, project);

                        options.WorkItemQuerySpec.ReplaceQueryTokens(options, project.Name);

                        queryResults = AzureDevOpsExplorer.Presentation.Views.Server.WorkItemStore.Query(options.WorkItemQuerySpec.Query);

                        Body_WorkItemStore.Add_TP_WorkItemDetails(insertAt, options, queryResults);

                        Globals.ThisAddIn.Application.ActiveWorkbook.Save();

                        Log.Trace($"EndProcessing {project.Name}", Common.LOG_CATEGORY, loopTicks);

                        AZDOHelper.ProcessLoopDelay(options);
                    }
                }
                else
                {
                    options.WorkItemQuerySpec.ReplaceQueryTokens(options);

                    queryResults = AzureDevOpsExplorer.Presentation.Views.Server.WorkItemStore.Query(options.WorkItemQuerySpec.Query);

                    Body_WorkItemStore.Add_TP_WorkItemDetails(insertAt, options, queryResults);

                    Globals.ThisAddIn.Application.ActiveWorkbook.Save();
                }

                insertAt.MarkEnd(XlHlp.MarkType.GroupTable, string.Format("tbl_{0}", insertAt.workSheet.Name));

                insertAt.Group(insertAt.OrientVertical, hide: true);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void CreateWS_All_TPC_WorkItemFields(Options_AZDO_TFS options)
        {
            long startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            try
            {
                XlHlp.XlLocation insertAt = CreateNewWorksheet(string.Format("{0}_{1}", "All_TPC", "WorkItemFields"),
                    options);

                XlHlp.AddLabeledInfoX(insertAt.AddRowX(), "Work Item Fields - All Team Projects", AzureDevOpsExplorer.Presentation.Views.Server.TfsTeamProjectCollection.Name);

                insertAt.MarkStart(XlHlp.MarkType.GroupTable);

                Header_WorkItemStore.Add_TP_WorkItemFields(insertAt);

                foreach (var teamProjectName in options.TeamProjects)
                {
                    Project project = AzureDevOpsExplorer.Presentation.Views.Server.WorkItemStore.Projects[teamProjectName];

                    long loopTicks = Log.Trace($"Processing    {project.Name}", Common.LOG_CATEGORY);

                    Globals.ThisAddIn.Application.StatusBar = $"Processing    {project.Name}";

                    Body_WorkItemStore.Add_TP_WorkItemFields(insertAt, options, project);

                    // Save with each loop so we loose less when things crash ;(

                    Globals.ThisAddIn.Application.ActiveWorkbook.Save();

                    Log.Trace($"EndProcessing {project.Name}", Common.LOG_CATEGORY, loopTicks);

                    AZDOHelper.ProcessLoopDelay(options);
                }

                insertAt.MarkEnd(XlHlp.MarkType.GroupTable, string.Format("tbl_{0}", insertAt.workSheet.Name));

                insertAt.Group(insertAt.OrientVertical, hide: true);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void CreateWS_All_TPC_WorkItemTypes(Options_AZDO_TFS options)
        {
            long startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            try
            {
                XlHlp.XlLocation insertAt = CreateNewWorksheet("All_TPC_WorkItemTypes", options);

                XlHlp.AddLabeledInfoX(insertAt.AddRowX(), "Work Item Types All Team Projects",
                    AzureDevOpsExplorer.Presentation.Views.Server.TfsTeamProjectCollection.Name);

                insertAt.MarkStart(XlHlp.MarkType.GroupTable);

                Header_WorkItemStore.Add_TP_WorkItemTypes(insertAt);

                foreach (var teamProjectName in options.TeamProjects)
                {
                    Project project = AzureDevOpsExplorer.Presentation.Views.Server.WorkItemStore.Projects[teamProjectName];

                    long loopTicks = Log.Trace($"Processing    {project.Name}", Common.LOG_CATEGORY);

                    Globals.ThisAddIn.Application.StatusBar = $"Processing    {project.Name}";

                    DateTime maxLastCreatedDate = DateTime.MinValue;
                    DateTime maxLastChangedDate = DateTime.MinValue;
                    DateTime maxLastRevisedDate = DateTime.MinValue;

                    Body_WorkItemStore.Add_TP_WorkItemTypes(insertAt,
                        options, AzureDevOpsExplorer.Presentation.Views.Server.WorkItemStore,
                        project,
                        out maxLastCreatedDate, out maxLastChangedDate, out maxLastRevisedDate);

                    // Save with each loop so we loose less when things crash ;(

                    Globals.ThisAddIn.Application.ActiveWorkbook.Save();

                    Log.Trace($"EndProcessing {project.Name}", Common.LOG_CATEGORY, loopTicks);

                    AZDOHelper.ProcessLoopDelay(options);
                }

                insertAt.MarkEnd(XlHlp.MarkType.GroupTable, string.Format("tbl_{0}", insertAt.workSheet.Name));

                insertAt.Group(insertAt.OrientVertical, hide: true);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void CreateWS_TPC_Info(
            CatalogNode teamProjectCollectionNode,
            TfsTeamProjectCollection teamProjectCollection,
            bool showDetails,
            Options_AZDO_TFS options)
        {
            long startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            try
            {
                XlHlp.XlLocation insertAt = CreateNewWorksheet("TPC_Info-" + GetTeamProjectCollectionName(teamProjectCollection),
                    options);

                insertAt = Section_TeamProjectCollection.Add_Info(insertAt, options, teamProjectCollection, showDetails);

                // Get a catalog of team projects for the collection

                ReadOnlyCollection<CatalogNode> teamProjects = teamProjectCollectionNode.QueryChildren(
                    new[] { CatalogResourceTypes.TeamProject }, false, CatalogQueryOptions.None);

                insertAt = Section_TeamProjectCollection.AddSection_TeamProjects(insertAt, teamProjects);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);
        }

        //void CreateWS_All_TPC_LastChangeset(Options_AZDO_TFS options,
        //    VersionControlServer versionControlServer)
        //{
        //    long startTicks = Log.Trace("Enter", Common.LOG_CATEGORY);

        //    try
        //    {
        //        XlHlp.XlLocation insertAt = CreateNewWorksheet(string.Format("{0}_{1}", "All_TPC", "LastChangeset"),
        //            options);

        //        XlHlp.AddLabeledInfoX(insertAt.AddRowX(), "Last Changeset All TeamProjects", AzureDevOpsExplorer.Presentation.Views.Server.TfsTeamProjectCollection.Name);

        //        insertAt.MarkStart(XlHlp.MarkType.GroupTable);

        //        Header_VersionControlServer.Add_Changesets(insertAt);

        //        //Body_VersionControlServer.Add_Changesets(insertAt, options, versionControlServer);

        //        foreach (var teamProjectName in options.TeamProjects)
        //        {
        //            TeamProject teamProject = VNCTFS.Helper.Get_TeamProject(versionControlServer, teamProjectName.Trim());

        //            long loopTicks = Log.Trace($"Processing {teamProject.Name}", Common.LOG_CATEGORY);

        //            Globals.ThisAddIn.Application.StatusBar = $"Processing {teamProject.Name}";

        //            insertAt.ClearOffsets();

        //            XlHlp.AddContentToCell(insertAt.AddRowX(), teamProject.Name);

        //            Body_VersionControlServer.Add_TP_Changesets(insertAt, options,
        //                AzureDevOpsExplorer.Presentation.Views.Server.VersionControlServer, teamProjectName);

        //            Log.Trace($"EndProcessing {teamProject.Name}", Common.LOG_CATEGORY, loopTicks);

        //            AZDOHelper.ProcessLoopDelay(options);
        //        }

        //        insertAt.MarkEnd(XlHlp.MarkType.GroupTable, string.Format("tbl_{0}", insertAt.workSheet.Name));

        //        insertAt.Group(insertAt.OrientVertical, hide: true);
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.ToString());
        //    }

        //    Log.Trace("Exit", Common.LOG_CATEGORY, startTicks);
        //}

        private void CreateWS_TPC_Members(string teamProjectCollectionName,
            Options_AZDO_TFS options)
        {
            long startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            try
            {
                XlHlp.XlLocation insertAt = CreateNewWorksheet(string.Format("{0}_{1}", "All_TPC", "Members"),
                    options);

                XlHlp.AddLabeledInfoX(insertAt.AddRowX(), "Name:", teamProjectCollectionName);

                insertAt = Section_TeamProjectCollection.Add_Members(insertAt, options);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion All_TPC

        private void CreateWS_TP(string teamProjectName,
            TeamProjectActionRequest request,
            Options_AZDO_TFS options)
        {
            long startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            try
            {
                XlHlp.XlLocation insertAt = CreateNewWorksheet($"TP_{teamProjectName}", options);

                // TODO(crhodes)
                // Figure out how to handle this more cleanly.  Bombs if no VCS entry, e.g. uses GIT

                try
                {
                    TeamProject teamProject = null;
                    ProjectInfo projectInfo = null;

                    if (request.TPSections.Count > 0 || request.VCSSections.Count > 0)
                    {
                        teamProject = VNCTFS.Helper.Get_TeamProject(AzureDevOpsExplorer.Presentation.Views.Server.VersionControlServer, teamProjectName.Trim());
                        projectInfo = AzureDevOpsExplorer.Presentation.Views.Server.CommonStructureService.GetProjectFromName(teamProject.Name);
                    }

                    if (request.TPSections.Count > 0)
                    {
                        insertAt = Section_TeamProject.AddSections(insertAt, teamProject, request.TPSections);

                        insertAt.IncrementRows();
                    }

                    if (request.VCSSections.Count > 0)
                    {
                        insertAt = Section_VersionControlServer.AddSections(insertAt,
                            options,
                            AzureDevOpsExplorer.Presentation.Views.Server.VersionControlServer,
                            teamProject,
                            projectInfo,
                            request.VCSSections);

                        insertAt.IncrementRows();
                    }
                }
                catch (Exception ex)
                {
                    XlHlp.AddContentToCell(insertAt.AddRowX(), "TeamProject (TP) Information Not Available");

                    insertAt = new XlHlp.XlLocation(insertAt.workSheet, row: 18, column: 1, options.OrientOutputVertically);
                }

                if (request.WISSections.Count > 0)
                {
                    Project project = AzureDevOpsExplorer.Presentation.Views.Server.WorkItemStore.Projects[teamProjectName];

                    insertAt = Section_WorkItemStore.AddSections(insertAt, options,
                        AzureDevOpsExplorer.Presentation.Views.Server.WorkItemStore,
                        AzureDevOpsExplorer.Presentation.Views.Server.CommonStructureService,
                        project,
                        request.WISSections);

                    insertAt.IncrementRows();
                }

                if (request.TMSections.Count > 0)
                {
                    Project project = AzureDevOpsExplorer.Presentation.Views.Server.WorkItemStore.Projects[teamProjectName];
                    ITestManagementTeamProject testManagementTeamProject = AzureDevOpsExplorer.Presentation.Views.Server.TestManagementService.GetTeamProject(project);

                    insertAt = Section_TestManager.AddSections_TM(insertAt,
                        options,
                        testManagementTeamProject,
                        request.TMSections);
                }

                if (request.BSSections.Count > 0)
                {
                    insertAt = Section_BuildServer.AddSections(insertAt,
                        options,
                        AzureDevOpsExplorer.Presentation.Views.Server.BuildServer,
                        teamProjectName,
                        request.BSSections);

                    insertAt.IncrementRows();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void CreateWS_TP_Queries(
            WorkItemStore workItemStore,
            Project project,
            Dictionary<string, string> queries,
            Options_AZDO_TFS options)
        {
            long startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            try
            {
                XlHlp.XlLocation insertAt = CreateNewWorksheet($"TPQ_{project.Name}", options);

                // TODO(crhodes)
                // We don't support multiple queries yet.
                // Hope I don't regret not passing in the query stuff.
                //foreach (string queryName in queries.Keys)
                //{
                //    insertAt = Section_WorkItemStore.Add_TP_Query(insertAt, options, project, queries[queryName], queryName);
                //}
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void CreateWS_TP_TemplateType(Options_AZDO_TFS options)
        {
            long startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            try
            {
                XlHlp.XlLocation insertAt = CreateNewWorksheet(string.Format("{0}_{1}", "TPC", "Templates"),
                    options);

                ReadOnlyCollection<CatalogNode> teamProjectCollectionNodes = VNCTFS.Helper.Get_TeamProjectCollectionNodes(AzureDevOpsExplorer.Presentation.Views.Server.ConfigurationServer);

                foreach (CatalogNode tpcNode in teamProjectCollectionNodes)
                {
                    TfsTeamProjectCollection tpc = VNCTFS.Helper.Get_TeamProjectCollection(AzureDevOpsExplorer.Presentation.Views.Server.ConfigurationServer, tpcNode);

                    foreach (Microsoft.TeamFoundation.WorkItemTracking.Client.Project project in AzureDevOpsExplorer.Presentation.Views.Server.WorkItemStore.Projects)
                    {
                        CategoryCollection categories = AzureDevOpsExplorer.Presentation.Views.Server.WorkItemStore.Projects[project.Name].Categories;

                        XlHlp.AddLabeledInfoX(insertAt.AddRowX(), "TP Name", string.Format("{0} ({1})", project.Name, categories.Count()));

                        DisplayListOf_Categories(insertAt, categories);

                        AZDOHelper.ProcessLoopDelay(options);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion CreateWS_*

        #region VersionControlServer (VCS)

        private void CreateWS_VCS_Branches(string teamProjectCollectionName,
            Options_AZDO_TFS options)
        {
            long startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            try
            {
                XlHlp.XlLocation insertAt = CreateNewWorksheet(string.Format("{0}_{1}", "All_TP", "Root Branches"),
                    options);

                XlHlp.AddLabeledInfoX(insertAt.AddRowX(), "VCS Branches All Team Projects", AzureDevOpsExplorer.Presentation.Views.Server.TfsTeamProjectCollection.Name);

                insertAt = Section_VersionControlServer.Add_RootBranches(insertAt, options, AzureDevOpsExplorer.Presentation.Views.Server.VersionControlServer);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void CreateWS_VCS_ChangeSetInfo(int changesetId,
            string sectionsToDisplay,
            Options_AZDO_TFS options,
            ICommonStructureService commonStructureService)
        {
            long startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            try
            {
                Changeset changeSet = AzureDevOpsExplorer.Presentation.Views.Server.VersionControlServer.GetChangeset(changesetId);

                XlHlp.XlLocation insertAt = CreateNewWorksheet(string.Format("CS_{0}", changesetId),
                    options);

                // TODO(crhodes)
                // Change this to write to Sheet
                //
                // Use first thing to get code to add to Version Control Server (VCS) Information sheet

                if (sectionsToDisplay.Contains("Info"))
                {
                    Section_VersionControlServer.Display_VCS_Changeset_Info(changeSet);
                }

                if (sectionsToDisplay.Contains("Changes"))
                {
                    Section_VersionControlServer.Display_VCS_ChangeSet_Changes(changeSet);
                }

                if (sectionsToDisplay.Contains("Associated WorkItems"))
                {
                    Section_VersionControlServer.Display_VCS_ChangeSet_AssociatedWorkItems(changeSet);
                }

                if (sectionsToDisplay.Contains("WorkItems"))
                {
                    Section_VersionControlServer.Display_VCS_Changeset_WorkItems(changeSet, AzureDevOpsExplorer.Presentation.Views.Server.WorkItemStore, commonStructureService);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void CreateWS_VCS_ShelveSets(string teamProjectCollectionName,
            Options_AZDO_TFS options)
        {
            long startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            try
            {
                XlHlp.XlLocation insertAt = CreateNewWorksheet(string.Format("{0}_{1}", "All_TP", "ShelveSets"),
                    options);

                XlHlp.AddLabeledInfoX(insertAt.AddRowX(), "Name:", teamProjectCollectionName);

                insertAt = Section_VersionControlServer.Add_Shelvesets(insertAt, options, AzureDevOpsExplorer.Presentation.Views.Server.VersionControlServer);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void CreateWS_VCS_Workspaces(string tfsUri,
            string teamProjectCollection,
            Options_AZDO_TFS options)
        {
            long startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            try
            {
                XlHlp.XlLocation insertAt = CreateNewWorksheet(string.Format("{0}_{1}", "All_TP", "WorkSpaces"),
                    options);

                insertAt = Section_VersionControlServer.Add_Workspaces(insertAt, options, AzureDevOpsExplorer.Presentation.Views.Server.VersionControlServer);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion VersionControlServer (VCS)

        #region WorkItemStore (WIS)

        private void CreateWS_WIS_WorkItemInfo(int workItemID,
            WorkItemActionRequest request,
            Options_AZDO_TFS options)
        {
            long startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            try
            {
                WorkItem workItem = VNCTFS.Helper.RetrieveWorkItem(workItemID, AzureDevOpsExplorer.Presentation.Views.Server.WorkItemStore);

                if (workItem is null)
                {
                    MessageBox.Show($"{workItemID} not found.");
                }

                XlHlp.XlLocation insertAt = CreateNewWorksheet($"{workItem.Type.Name}_{workItem.Id}", options);

                if (request.WorkItemSections.Contains("Info"))
                {
                    insertAt = Section_WorkItemStore.Add_WorkItem_Info(insertAt, options, workItem);
                    insertAt.IncrementRows();
                }

                if (request.WorkItemSections.Contains("Fields"))
                {
                    insertAt = Section_WorkItemStore.Add_WorkItem_Fields(insertAt, options, workItem, request);
                    insertAt.IncrementRows();
                }

                if (request.WorkItemSections.Contains("TestFields"))
                {
                    insertAt = Section_WorkItemStore.Test_WorkItem_Fields(insertAt, options, workItem, request);
                    insertAt.IncrementRows();
                }

                if (request.WorkItemSections.Contains("PlainLinks"))
                {
                    insertAt = Section_WorkItemStore.Add_WorkItem_Links(insertAt, options,
                        AzureDevOpsExplorer.Presentation.Views.Server.WorkItemStore, workItemID);
                    insertAt.IncrementRows();
                }

                if (request.WorkItemSections.Contains("WorkItemLinks"))
                {
                    WorkItem workitem = VNCTFS.Helper.RetrieveWorkItem(workItemID, AzureDevOpsExplorer.Presentation.Views.Server.WorkItemStore);

                    // TODO(crhodes)
                    // Check for null

                    insertAt = Section_WorkItemStore.Add_WorkItem_WorkItemLinks(insertAt, options,
                        AzureDevOpsExplorer.Presentation.Views.Server.WorkItemStore, workitem);
                    insertAt.IncrementRows();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion WorkItemStore (WIS)

        #region Test Manager (TM)

        private void CreateWS_TM_TestCaseInfo(int testCaseId, List<string> sectionsToDisplay, Options_AZDO_TFS options)
        {
            long startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            // Get Test Case Info

            WorkItem wi = VNCTFS.Helper.RetrieveWorkItem(testCaseId, AzureDevOpsExplorer.Presentation.Views.Server.WorkItemStore);

            // Get Team Project for Test Plan

            string teamProjectName = wi.Project.Name;

            Project project = AzureDevOpsExplorer.Presentation.Views.Server.WorkItemStore.Projects[teamProjectName];

            TestManagementService testManagementService = AzureDevOpsExplorer.Presentation.Views.Server.TestManagementService;

            ITestManagementTeamProject testManagementTeamProject = AzureDevOpsExplorer.Presentation.Views.Server.TestManagementService.GetTeamProject(project);

            ITestCase testCase = testManagementTeamProject.TestCases.Find(testCaseId);

            XlHlp.XlLocation insertAt = CreateNewWorksheet($"TestCaseInfo_{testCaseId}", options);

            if (sectionsToDisplay.Contains("Info"))
            {
                insertAt = Section_TestManager.Add_TestCase_Info(insertAt, options, testCase);
                insertAt.IncrementRows();
            }

            WorkItem workItem = null;

            if (sectionsToDisplay.Contains("WorkItemInfo"))
            {
                workItem = VNCTFS.Helper.RetrieveWorkItem(testCaseId, AzureDevOpsExplorer.Presentation.Views.Server.WorkItemStore);

                insertAt = Section_WorkItemStore.Add_WorkItem_Info(insertAt, options, workItem);
                insertAt.IncrementRows();
            }

            if (sectionsToDisplay.Contains("WorkItemLinks"))
            {
                if (workItem == null)
                {
                    workItem = VNCTFS.Helper.RetrieveWorkItem(testCaseId, AzureDevOpsExplorer.Presentation.Views.Server.WorkItemStore);
                }

                insertAt = Section_WorkItemStore.Add_WorkItem_WorkItemLinks(insertAt, options, AzureDevOpsExplorer.Presentation.Views.Server.WorkItemStore, workItem);
                insertAt.IncrementRows();
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void CreateWS_TM_TestPlanInfo(int testPlanId, List<string> sectionsToDisplay, Options_AZDO_TFS options)
        {
            long startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            // Get Test Plan Info

            WorkItem wi = VNCTFS.Helper.RetrieveWorkItem(testPlanId, AzureDevOpsExplorer.Presentation.Views.Server.WorkItemStore);

            // Get Team Project for Test Plan

            string teamProjectName = wi.Project.Name;

            Project project = AzureDevOpsExplorer.Presentation.Views.Server.WorkItemStore.Projects[teamProjectName];

            TestManagementService testManagementService = AzureDevOpsExplorer.Presentation.Views.Server.TestManagementService;

            ITestManagementTeamProject testManagementTeamProject = AzureDevOpsExplorer.Presentation.Views.Server.TestManagementService.GetTeamProject(project);

            ITestPlan testPlan = testManagementTeamProject.TestPlans.Find(testPlanId);

            IStaticTestSuite staticTestSuite = testPlan.RootSuite;

            ITestSuiteCollection testSuites = staticTestSuite.SubSuites;

            ITestSuiteBase testSuiteBase = testSuites[0];

            ITestCaseCollection testCaseCollection = testSuiteBase.AllTestCases;

            XlHlp.XlLocation insertAt = CreateNewWorksheet($"TestPlanInfo_{testPlanId}", options);

            if (sectionsToDisplay.Contains("Info"))
            {
                insertAt = Section_TestManager.Add_TestPlan_Info(insertAt, options, testPlan);
                insertAt.IncrementRows();
            }

            WorkItem workItem = null;

            if (sectionsToDisplay.Contains("WorkItemInfo"))
            {
                workItem = VNCTFS.Helper.RetrieveWorkItem(testPlan.Id, AzureDevOpsExplorer.Presentation.Views.Server.WorkItemStore);

                insertAt = Section_WorkItemStore.Add_WorkItem_Info(insertAt, options, workItem);
                insertAt.IncrementRows();
            }

            if (sectionsToDisplay.Contains("WorkItemLinks"))
            {
                if (workItem == null)
                {
                    workItem = VNCTFS.Helper.RetrieveWorkItem(testPlanId, AzureDevOpsExplorer.Presentation.Views.Server.WorkItemStore);
                }

                insertAt = Section_WorkItemStore.Add_WorkItem_WorkItemLinks(insertAt, options, AzureDevOpsExplorer.Presentation.Views.Server.WorkItemStore, workItem);
                insertAt.IncrementRows();
            }

            if (sectionsToDisplay.Contains("RootTestSuite"))
            {
                insertAt = Section_TestManager.Add_TestSuite_Info(insertAt, options, testPlan.RootSuite);
                insertAt.IncrementRows();
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void CreateWS_TM_TestSuiteInfo(int testSuiteId, List<string> sectionsToDisplay, Options_AZDO_TFS options)
        {
            long startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            // Get Test Suite Info

            WorkItem wi = VNCTFS.Helper.RetrieveWorkItem(testSuiteId, AzureDevOpsExplorer.Presentation.Views.Server.WorkItemStore);

            // Get Team Project for Test Plan

            string teamProjectName = wi.Project.Name;

            Project project = AzureDevOpsExplorer.Presentation.Views.Server.WorkItemStore.Projects[teamProjectName];

            TestManagementService testManagementService = AzureDevOpsExplorer.Presentation.Views.Server.TestManagementService;

            ITestManagementTeamProject testManagementTeamProject = AzureDevOpsExplorer.Presentation.Views.Server.TestManagementService.GetTeamProject(project);

            ITestSuiteBase testSuite = testManagementTeamProject.TestSuites.Find(testSuiteId);

            if (testSuite is null)
            {
                MessageBox.Show($"Cannot find TestSuite: {testSuiteId}");
            }

            ITestCaseCollection testCaseCollection = testSuite.AllTestCases;

            XlHlp.XlLocation insertAt = CreateNewWorksheet($"TestSuiteInfo_{testSuiteId}", options);

            if (sectionsToDisplay.Contains("Info"))
            {
                insertAt = Section_TestManager.Add_TestSuite_Info(insertAt, options, testSuite);
                insertAt.IncrementRows();
            }

            WorkItem workItem = null;

            if (sectionsToDisplay.Contains("WorkItemInfo"))
            {
                workItem = VNCTFS.Helper.RetrieveWorkItem(testSuiteId, AzureDevOpsExplorer.Presentation.Views.Server.WorkItemStore);

                insertAt = Section_WorkItemStore.Add_WorkItem_Info(insertAt, options, workItem);
                insertAt.IncrementRows();
            }

            if (sectionsToDisplay.Contains("WorkItemLinks"))
            {
                if (workItem == null)
                {
                    workItem = VNCTFS.Helper.RetrieveWorkItem(testSuiteId, AzureDevOpsExplorer.Presentation.Views.Server.WorkItemStore);
                }

                insertAt = Section_WorkItemStore.Add_WorkItem_WorkItemLinks(insertAt, options, AzureDevOpsExplorer.Presentation.Views.Server.WorkItemStore, workItem);
                insertAt.IncrementRows();
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion Test Manager (TM)

        private void DisplayListOf_Categories(XlHlp.XlLocation insertAt, Microsoft.TeamFoundation.WorkItemTracking.Client.CategoryCollection categories)
        {
            long startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            insertAt.MarkStart(XlHlp.MarkType.Group);

            foreach (Category category in categories.OrderBy(n => n.Name))
            {
                StringBuilder sbWorkItemTypes = new StringBuilder();

                insertAt.ClearOffsets();

                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), category.Name);
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), category.ReferenceName);
                int count = category.WorkItemTypes.Count();
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), count.ToString());

                sbWorkItemTypes.Append("WorkItemTypes:(");

                if (count > 1)
                {
                    foreach (WorkItemType item in category.WorkItemTypes)
                    {
                        if (sbWorkItemTypes.Length > 0) sbWorkItemTypes.Append(";");
                        sbWorkItemTypes.Append(item.Name);
                    }
                }
                else
                {
                    sbWorkItemTypes.Append(category.DefaultWorkItemType.Name);
                }

                sbWorkItemTypes.Append(")");

                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), category.DefaultWorkItemType.Name);
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), category.DefaultWorkItemType.Description);
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), category.DefaultWorkItemType.Description);
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), sbWorkItemTypes.ToString());

                //string templateType = "???";

                //switch (category.Name)
                //{
                //    case "Requirement Category":
                //        XlHlp.AddContentToCell(insertAt.AddOffsetColumn(), category.Name);
                //        XlHlp.AddContentToCell(insertAt.AddOffsetColumn(), category.DefaultWorkItemType.Name);

                //        switch (category.DefaultWorkItemType.Name)
                //        {
                //            case "User Story":
                //                templateType = "Agile";
                //                break;

                //            case "Requirement":
                //                templateType = "CMMI";
                //                break;

                //            case "Product Backlog Item":
                //                templateType = "Scrum";
                //                break;

                //            default:
                //                break;
                //        }

                //        XlHlp.AddContentToCell(insertAt.AddOffsetColumn(), templateType);
                //        break;

                //    default:
                //        XlHlp.AddContentToCell(insertAt.AddOffsetColumn(), templateType);
                //        break;
                //}

                insertAt.IncrementRows();
            }

            insertAt.ClearOffsets();
            insertAt.IncrementRows();

            insertAt.MarkEnd(XlHlp.MarkType.Group);

            insertAt.Group(insertAt.OrientVertical, hide: true);

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);
        }
    }
}