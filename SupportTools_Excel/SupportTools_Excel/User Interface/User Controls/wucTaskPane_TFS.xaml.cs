using System;
using System.Data;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Windows;
using System.Windows.Controls;
//using System.Windows.Forms;
using System.Windows.Input;

using DevExpress.Xpf.Core;

using Microsoft.Office.Interop.Excel;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.VersionControl.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;

using SupportTools_Excel.AzureDevOpsExplorer.Application;
using SupportTools_Excel.AzureDevOpsExplorer.Domain;
using SupportTools_Excel.AzureDevOpsExplorer.Presentation.ModelWrappers;
using SupportTools_Excel.AzureDevOpsExplorer.Presentation.ViewModels;

using VNC;
using VNC.TFS.User_Interface.User_Controls;

using XlHlp = VNC.AddinHelper.Excel;

namespace SupportTools_Excel.User_Interface.User_Controls
{

    [System.Diagnostics.CodeAnalysis.SuppressMessage("Globalization", "CA1305:Specify IFormatProvider", Justification = "<Pending>")]
    public partial class wucTaskPane_TFS : UserControl
    {
        // This is updated with the TFS Server changes
        #region Constructors and Load

        public wucTaskPane_TFS()
        {
            long startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            //try
            //{
            //    var bootstrapper = new Bootstrapper();
            //    bootstrapper.Run();
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.ToString());
            //}

            InitializeComponent();
            LoadControlContents();

            // Server

            Common.EventAggregator.GetEvent<GetConfigurationServerInfoEvent>().Subscribe(GetConfigurationServerInfo);
            Common.EventAggregator.GetEvent<PopulateTeamProjectsEvent>().Subscribe(PopulateTeamProjects);
            Common.EventAggregator.GetEvent<EnableMainUIEvent>().Subscribe(EnableMainUI);

            //// Queries

            //Common.EventAggregator.GetEvent<RunQueryEvent>().Subscribe(RunQuery);
            //Common.EventAggregator.GetEvent<RunTeamProjectQueryEvent>().Subscribe(RunTeamProjectQuery);
            //Common.EventAggregator.GetEvent<RunTeamProjectQueriesEvent>().Subscribe(RunTeamProjectQueries);

            // WorkItems

            Common.EventAggregator.GetEvent<WorkItemIDDoubleClickEvent>().Subscribe(WorkItemIDDoubleClick);
            Common.EventAggregator.GetEvent<GetWorkItemInfoEvent>().Subscribe(GetWorkItemInfo);
            Common.EventAggregator.GetEvent<AddPivotSummaryEvent>().Subscribe(AddPivotSummary);

            // TeamProjectActions

            Common.EventAggregator.GetEvent<GetTeamProjectInfoEvent>().Subscribe(GetTeamProjectInfo);
            Common.EventAggregator.GetEvent<GetTeamProjectXMLEvent>().Subscribe(GetTeamProjectXML);

            // AZDO OrganizationActions

            Common.EventAggregator.GetEvent<GetTPCInfoEvent>().Subscribe(Get_TPC_Info);
            Common.EventAggregator.GetEvent<GetTPCAreasEvent>().Subscribe(Get_TPC_Areas);
            Common.EventAggregator.GetEvent<GetBranchesEvent>().Subscribe(GetBranches);
            Common.EventAggregator.GetEvent<GetAllTPDevelopersEvent>().Subscribe(Get_All_TPC_Developers);
            Common.EventAggregator.GetEvent<GetTPCMembersEvent>().Subscribe(Get_TPC_Members);
            Common.EventAggregator.GetEvent<GetTPCShelfsetsEvent>().Subscribe(Get_TPC_Shelfsets);
            Common.EventAggregator.GetEvent<GetTPCBuildDefinitionsEvent>().Subscribe(Get_TPC_BuildDefinitions);
            Common.EventAggregator.GetEvent<GetTPCTeamsEvent>().Subscribe(Get_TPC_Teams);
            Common.EventAggregator.GetEvent<GetTPCWorkItemTypesEvent>().Subscribe(Get_TPC_WorkItemTypes);
            Common.EventAggregator.GetEvent<GetTPCWorkItemFieldsEvent>().Subscribe(Get_TPC_WorkItemFields);
            Common.EventAggregator.GetEvent<GetTPCWorkItemDetailsEvent>().Subscribe(Get_TPC_WorkItemDetails);
            Common.EventAggregator.GetEvent<GetTPCWorkspacesEvent>().Subscribe(Get_TPC_Workspaces);
            Common.EventAggregator.GetEvent<GetTPCLastChangesetEvent>().Subscribe(Get_TPC_LastChangeset);
            Common.EventAggregator.GetEvent<GetTPCWorkItemActivityEvent>().Subscribe(Get_TPC_WorkItemActivity);
            Common.EventAggregator.GetEvent<GetTPCTestPlansEvent>().Subscribe(Get_TPC_TestPlans);
            Common.EventAggregator.GetEvent<GetTPCTestSuitesEvent>().Subscribe(Get_TPC_TestSuites);
            Common.EventAggregator.GetEvent<GetTPCTestCasesEvent>().Subscribe(Get_TPC_TestCases);

            // AZDOTestManagementActions

            Common.EventAggregator.GetEvent<TestPlanIDDoubleClickEvent>().Subscribe(TestPlanIdDoubleClick);
            Common.EventAggregator.GetEvent<GetTestPlanInfoEvent>().Subscribe(GetTestPlanInfo);

            Common.EventAggregator.GetEvent<TestSuiteIDDoubleClickEvent>().Subscribe(TestSuiteIdDoubleClick);
            Common.EventAggregator.GetEvent<GetTestSuiteInfoEvent>().Subscribe(GetTestSuiteInfo);

            Common.EventAggregator.GetEvent<TestCaseIDDoubleClickEvent>().Subscribe(TestCaseIdDoubleClick);
            Common.EventAggregator.GetEvent<GetTestCaseInfoEvent>().Subscribe(GetTestCaseInfo);

            Common.EventAggregator.GetEvent<AddTestPlanPivotSummaryEvent>().Subscribe(AddTestPlanPivotSummary);

            // Get Exception if do here.
            //try
            //{
            //    var bootstrapper = new Bootstrapper();
            //    bootstrapper.Run();
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.ToString());
            //}

            //RunQueryCommand = new DelegateCommand(RunQuery);

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void LoadControlContents()
        {
            long startTicks = Log.APPLICATION_INITIALIZE("Enter", Common.LOG_CATEGORY);

            try
            {
                ((OptionsViewModel)azdoOptions.ViewModel).Options.XMLTemplateFilePath = Common.cEXPORT_TEMPLATE_PATH;
                //AZDOOptions.teXMLTemplateFilePath.Text = Common.cEXPORT_TEMPLATE_PATH;
                //AZDOOptions.teGoBackDays.Text = Common.cGO_BACK_DAYS;

                // Set the UI to the initial state.  Xaml has everything expanded/visible

                lgMainGroup.Visibility = Visibility.Hidden;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            ApplicationThemeHelper.ApplicationThemeName = "MetropolisLight";

            Log.APPLICATION_INITIALIZE("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion Constructors and Load

        #region Prism Event Handlers

        #region AZDOTestManagementActions

        private void AddTestPlanPivotSummary()
        {
            long startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void TestPlanIdDoubleClick(TestPlanRequestWrapper request)
        {
            ProcessDoubleClick(request);
        }

        private void TestSuiteIdDoubleClick(TestSuiteRequestWrapper request)
        {
            ProcessDoubleClick(request);
        }

        private void TestCaseIdDoubleClick(TestCaseRequestWrapper request)
        {
            ProcessDoubleClick(request);
        }

        private static void ProcessDoubleClick(TestPlanRequestWrapper request)
        {
            long startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            string selectedCell = (string)Globals.ThisAddIn.Application.ActiveCell.Value.ToString();

            Microsoft.Office.Interop.Excel.Range selectedRange = (Range)Globals.ThisAddIn.Application.Selection;
            selectedRange.SpecialCells(XlCellType.xlCellTypeVisible).Select();
            Microsoft.Office.Interop.Excel.Range selectedVisibleRange = (Range)Globals.ThisAddIn.Application.Selection;

            StringBuilder selectedCellsText = new StringBuilder();

            if (selectedRange.Count > 1)
            {
                foreach (Range cell in selectedVisibleRange.Cells)
                {
                    if (selectedCellsText.Length > 0)
                    {
                        selectedCellsText.Append($", {cell.Value}");
                    }
                    else
                    {
                        selectedCellsText.Append($"{cell.Value}");
                    }
                }
            }
            else
            {
                selectedCellsText.Append($"{selectedRange.Value}");
            }

            request.TestID = selectedCellsText.ToString();

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private static void ProcessDoubleClick(TestSuiteRequestWrapper request)
        {
            long startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            string selectedCell = (string)Globals.ThisAddIn.Application.ActiveCell.Value.ToString();

            Range selectedRange = (Range)Globals.ThisAddIn.Application.Selection;
            selectedRange.SpecialCells(XlCellType.xlCellTypeVisible).Select();
            Range selectedVisibleRange = (Range)Globals.ThisAddIn.Application.Selection;

            StringBuilder selectedCellsText = new StringBuilder();

            if (selectedRange.Count > 1)
            {
                foreach (Microsoft.Office.Interop.Excel.Range cell in selectedVisibleRange.Cells)
                {
                    if (selectedCellsText.Length > 0)
                    {
                        selectedCellsText.Append($", {cell.Value}");
                    }
                    else
                    {
                        selectedCellsText.Append($"{cell.Value}");
                    }
                }
            }
            else
            {
                selectedCellsText.Append($"{selectedRange.Value}");
            }

            request.TestID = selectedCellsText.ToString();

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private static void ProcessDoubleClick(TestCaseRequestWrapper request)
        {
            long startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            string selectedCell = (string)Globals.ThisAddIn.Application.ActiveCell.Value.ToString();

            Microsoft.Office.Interop.Excel.Range selectedRange = (Range)Globals.ThisAddIn.Application.Selection;
            selectedRange.SpecialCells(XlCellType.xlCellTypeVisible).Select();
            Microsoft.Office.Interop.Excel.Range selectedVisibleRange = (Range)Globals.ThisAddIn.Application.Selection;

            StringBuilder selectedCellsText = new StringBuilder();

            if (selectedRange.Count > 1)
            {
                foreach (Microsoft.Office.Interop.Excel.Range cell in selectedVisibleRange.Cells)
                {
                    if (selectedCellsText.Length > 0)
                    {
                        selectedCellsText.Append($", {cell.Value}");
                    }
                    else
                    {
                        selectedCellsText.Append($"{cell.Value}");
                    }
                }
            }
            else
            {
                selectedCellsText.Append($"{selectedRange.Value}");
            }

            request.TestID = selectedCellsText.ToString();

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private static void ProcessDoubleClick(WorkItemActionRequestWrapper request)
        {
            long startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            string selectedCell = (string)Globals.ThisAddIn.Application.ActiveCell.Value.ToString();

            Range selectedRange = (Range)Globals.ThisAddIn.Application.Selection;
            selectedRange.SpecialCells(XlCellType.xlCellTypeVisible).Select();
            Range selectedVisibleRange = (Range)Globals.ThisAddIn.Application.Selection;

            StringBuilder selectedCellsText = new StringBuilder();

            // TODO(crhodes)
            // Is the outer if even needed if only one cell?

            if (selectedRange.Count > 1)
            {
                foreach (Range cell in selectedVisibleRange.Cells)
                {
                    if (selectedCellsText.Length > 0)
                    {
                        selectedCellsText.Append($", {cell.Value}");
                    }
                    else
                    {
                        selectedCellsText.Append($"{cell.Value}");
                    }
                }
            }
            else
            {
                selectedCellsText.Append($"{selectedRange.Value}");
            }

            request.WorkItemID =  selectedCellsText.ToString();

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

 
        private void GetTestPlanInfo(TestPlanRequest request)
        {
            long startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            try
            {
                RequestHandlers.SpeedUpStart();

                char[] splitChars = { ',' };

                int testPlanId = 0;

                Options_AZDO_TFS options = GetOptions();

                foreach (string testPlan in request.TestID.Split(splitChars, StringSplitOptions.None))
                {
                    if (int.TryParse(testPlan, out testPlanId))
                    {
                        CreateWS_TM_TestPlanInfo(testPlanId, request.TestSections, options);
                    }

                    AZDOHelper.ProcessLoopDelay(options);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                RequestHandlers.SpeedUpEnd();
            }

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void GetTestSuiteInfo(TestSuiteRequest request)
        {
            long startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            try
            {
                RequestHandlers.SpeedUpStart();

                char[] splitChars = { ',' };

                int testSuiteId = 0;

                Options_AZDO_TFS options = GetOptions();

                foreach (string testSuite in request.TestID.Split(splitChars, StringSplitOptions.None))
                {
                    if (int.TryParse(testSuite, out testSuiteId))
                    {
                        CreateWS_TM_TestSuiteInfo(testSuiteId, request.TestSections, options);
                    }

                    AZDOHelper.ProcessLoopDelay(options);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                RequestHandlers.SpeedUpEnd();
            }

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void GetTestCaseInfo(TestCaseRequest request)
        {
            long startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            try
            {
                RequestHandlers.SpeedUpStart();

                char[] splitChars = { ',' };

                int testCaseId = 0;

                Options_AZDO_TFS options = GetOptions();

                foreach (string testCase in request.TestID.Split(splitChars, StringSplitOptions.None))
                {
                    if (int.TryParse(testCase, out testCaseId))
                    {
                        CreateWS_TM_TestCaseInfo(testCaseId, request.TestSections, options);
                    }

                    AZDOHelper.ProcessLoopDelay(options);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                RequestHandlers.SpeedUpEnd();
            }

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region WorkItem Actions

        private void WorkItemIDDoubleClick(WorkItemActionRequestWrapper request)
        {
            long startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            ProcessDoubleClick(request);

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void GetWorkItemInfo(WorkItemActionRequest request)
        {
            long startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            try
            {
                RequestHandlers.SpeedUpStart();

                char[] splitChars = { ',' };

                int workItemId = 0;

                Options_AZDO_TFS options = GetOptions();

                foreach (string workItem in request.WorkItemID.Split(splitChars, StringSplitOptions.None))
                {
                    if (int.TryParse(workItem, out workItemId))
                    {
                        CreateWS_WIS_WorkItemInfo(workItemId, request, options);

                        Globals.ThisAddIn.Application.ActiveWorkbook.Save();
                    }

                    AZDOHelper.ProcessLoopDelay(options);
                }             
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                RequestHandlers.SpeedUpEnd();
            }

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        void AddPivotSummary(WorkItemActionRequest request)
        {
            long startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Worksheet ws = (Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            Microsoft.Office.Interop.Excel.Range activeCell = Globals.ThisAddIn.Application.ActiveCell;

            Options_AZDO_TFS options = new Options_AZDO_TFS();

            int workItemID = int.Parse(request.WorkItemID);

            XlHlp.XlLocation insertAt = CreateNewWorksheet(string.Format("P_WI_{0}", workItemID), GetOptions());

            // TODO(crhodes)
            // Figure out how to get the table name from the active cell.

            var tableName = activeCell.ListObject.Name;

            PivotCache pc = wb.PivotCaches().Create(
                SourceType: XlPivotTableSourceType.xlDatabase,
                SourceData: tableName,
                Version: 6);

            string insertRange = $"{insertAt.workSheet.Name}!R{options.StartingRow}C{options.StartingColumn}";

            PivotTable pt = pc.CreatePivotTable(TableDestination: insertRange);

            // this is from the macro recording.  Not all may be needed or desired.

            pt.ColumnGrand = true;
            pt.HasAutoFormat = true;
            pt.DisplayErrorString = false;
            pt.DisplayNullString = true;
            pt.EnableDrilldown = true;
            pt.ErrorString = "";
            pt.MergeLabels = false;
            pt.NullString = "";
            pt.PageFieldOrder = 2;
            pt.PageFieldWrapCount = 0;
            pt.PreserveFormatting = true;
            pt.RowGrand = true;
            pt.SaveData = true;
            pt.PrintTitles = false;
            pt.RepeatItemsOnEachPrintedPage = true;
            pt.TotalsAnnotation = false;
            pt.CompactRowIndent = 1;
            pt.InGridDropZones = false;
            pt.DisplayFieldCaptions = true;
            pt.DisplayMemberPropertyTooltips = false;
            pt.DisplayContextTooltips = true;
            pt.ShowDrillIndicators = true;
            pt.PrintDrillIndicators = false;
            pt.AllowMultipleFilters = false;
            pt.SortUsingCustomLists = true;
            pt.FieldListSortAscending = false;
            pt.ShowValuesRow = false;
            pt.CalculatedMembersInFilters = false;
            pt.RowAxisLayout(XlLayoutRowType.xlCompactRow);

            pt.PivotCache().RefreshOnFileOpen = false;
            pt.PivotCache().MissingItemsLimit = XlPivotTableMissingItems.xlMissingItemsDefault;

            PivotField pf1 = (PivotField)pt.PivotFields("Source.Type");
            PivotField pf2 = (PivotField)pt.PivotFields("Target.Type");

            pf1.Orientation = XlPivotFieldOrientation.xlRowField;
            pf1.Position = 1;

            pf2.Orientation = XlPivotFieldOrientation.xlRowField;
            pf2.Position = 2;

            //pt.AddDataField(pf1, "Count", XlConsolidationFunction.xlCount);

            //pf2.Orientation = XlPivotFieldOrientation.xlRowField;
            //pf2.Position = 2;

            //With ActiveSheet.PivotTables("PivotTable1").PivotFields("Target.Type")
            //    .Orientation = xlRowField
            //    .Position = 1
            //End With

            //ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables(_
            //    "PivotTable1").PivotFields("Target.Type"), "Count of Target.Type", xlCount
            //With ActiveSheet.PivotTables("PivotTable1").PivotFields("SourceId")
            //    .Orientation = xlRowField
            //    .Position = 2
            //End With

            insertAt.workSheet.Select();

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        //void RunQuery(WorkItemQuery workItemQuery)
        ////void RunQuery(wucTFSQuery_Picker queryPicker)
        //{
        //    long startTicks = XlHlp.DisplayInWatchWindow("Begin");

        //    try
        //    {
        //        RequestHandlers.SpeedUpStart();

        //        CreateWS_Query(AzureDevOpsExplorer.Presentation.Views.Server.WorkItemStore, GetOptions());
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.ToString());
        //    }
        //    finally
        //    {
        //        RequestHandlers.SpeedUpEnd();
        //    }

        //    XlHlp.DisplayInWatchWindow("End", startTicks);
        //}

        //private void RunTeamProjectQuery(WorkItemQuery workItemQuery)
        //{
        //    long startTicks = XlHlp.DisplayInWatchWindow("Begin");

        //    Options_AZDO_TFS options = GetOptions();

        //    try
        //    {
        //        RequestHandlers.SpeedUpStart();

        //        foreach (string teamProjectName in options.TeamProjects)
        //        {
        //            Globals.ThisAddIn.Application.StatusBar = "Processing " + teamProjectName;

        //            Project project = AzureDevOpsExplorer.Presentation.Views.Server.WorkItemStore.Projects[teamProjectName];

        //            CreateWS_TP_Query(options, project);

        //            Globals.ThisAddIn.Application.ActiveWorkbook.Save();

        //            AZDOHelper.ProcessLoopDelay(options);
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.ToString());
        //    }
        //    finally
        //    {
        //        RequestHandlers.SpeedUpEnd();
        //    }

        //    XlHlp.DisplayInWatchWindow("End", startTicks);
        //}

        //public DelegateCommand RunQueryCommand { get; set; }
        //public DelegateCommand RunTeamProjectQueryCommand { get; set; }
        //public DelegateCommand RunTeamProjectQueriesCommand { get; set; }

        //private void RunTeamProjectQueries(WorkItemQuery workItemQuery)
        //{
        //    // TODO(crhodes)
        //    // This needs to take a collection of queries
        //}

        #endregion

        void EnableMainUI(Visibility visibility)
        {
            long startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            lgMainGroup.Visibility = visibility;

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void GetConfigurationServerInfo(wucTFSProvider_Picker serverProvider)
        {
            long startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            try
            {
                RequestHandlers.SpeedUpStart();

                CreateWS_ConfigurationServer_Info(GetOptions(), AzureDevOpsExplorer.Presentation.Views.Server.ConfigurationServer);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                RequestHandlers.SpeedUpEnd();
            }

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void GetTeamProjectInfo(TeamProjectActionRequest request)
        {
            long startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Options_AZDO_TFS options = GetOptions();

            var priorStatusBar = Globals.ThisAddIn.Application.StatusBar;

            try
            {
                Globals.ThisAddIn.Application.DisplayStatusBar = true;

                RequestHandlers.SpeedUpStart();

                if (options.TeamProjects is null)
                {
                    MessageBox.Show("No Team Project(s) selected, aborting");
                }
                else
                {
                    options.TeamProjects.Reverse();

                    foreach (string teamProjectName in options.TeamProjects)
                    {
                        try
                        {
                            Globals.ThisAddIn.Application.StatusBar = "Processing " + teamProjectName;

                            CreateWS_TP(teamProjectName, request, options);

                            Globals.ThisAddIn.Application.ActiveWorkbook.Save();

                            AZDOHelper.ProcessLoopDelay(options);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.ToString());
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                Globals.ThisAddIn.Application.StatusBar = priorStatusBar;
                RequestHandlers.SpeedUpEnd();
            }

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void GetTeamProjectXML(TeamProjectActionRequest request)
        {
            long startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Options_AZDO_TFS options = GetOptions();

            var priorStatusBar = Globals.ThisAddIn.Application.StatusBar;

            try
            {
                Globals.ThisAddIn.Application.DisplayStatusBar = true;

                RequestHandlers.SpeedUpStart();

                foreach (string teamProjectName in options.TeamProjects)
                {
                    try
                    {
                        Globals.ThisAddIn.Application.StatusBar = "Processing " + teamProjectName;

                        Project project = AzureDevOpsExplorer.Presentation.Views.Server.WorkItemStore.Projects[teamProjectName];

                        Body_WorkItemStore.Get_TP_WorkItemTypesXML(options, project);

                        //Globals.ThisAddIn.Application.ActiveWorkbook.Save();

                        AZDOHelper.ProcessLoopDelay(options);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                Globals.ThisAddIn.Application.StatusBar = priorStatusBar;
                RequestHandlers.SpeedUpEnd();
            }

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region Event Handlers

        private void btnCodeChurn_Click(object sender, RoutedEventArgs e)
        {
            long startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Options_AZDO_TFS options = GetOptions();

            try
            {
                RequestHandlers.SpeedUpStart();
                //TeamFoundationServer tfs = TeamFoundationServerFactory.GetServer("http://WhateverServerUrl");
                //IBuildServer buildServer = (IBuildServer)tfs.GetService(typeof(IBuildServer));
                //VersionControlServer VsServer = (VersionControlServer)tfs.GetService(typeof(VersionControlServer));
                //IBuildDetail build = buildServer.GetAllBuildDetails(new Uri("http://WhateverBuildUrl"));

                //List<IChangesetsummary> associatedChangesets = InformationNodeConverters.GetAssociatedChangesets(build);

                //foreach (IChangesetsummary changeSetData in associatedChangesets)
                //{
                //    Changeset changeSet = VsServer.GetChangeset(changeSetData.ChangesetId);
                //    foreach (Change change in changeSet.Changes)
                //    {
                //        bool a = change.Item.IsContentDestroyed;
                //        long b = change.Item.ContentLength;
                //    }
                //}
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                RequestHandlers.SpeedUpEnd();
            }

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void btnGet_All_TP_AreaPathCheck_Click(object sender, RoutedEventArgs e)
        {
            long startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            //ProcessCreateWorkSheetSections(CreateWS_All_TP_AreaCheck, cbeAreas.Text, GetDisplayOrientation());
            //XlHlp.DisplayInWatchWindow(string.Format("{0}",
            //    MethodBase.GetCurrentMethod().Name));

            //if (!ValidUISelections()) { return; }

            //Options_AZDO_TFS options = GetDisplayOrientation();

            //try
            //{
            //    SpeedUpStart();

            //    CreateWS_All_TP_AreaCheck(cbeAreas.Text, orientVertical);
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.ToString());
            //}
            //finally
            //{
            //    SpeedUpEnd();
            //}

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void Get_All_TPC_Developers()
        {
            long startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            RequestHandlers.ProcessCreateWorkSheet(CreateWS_All_TPC_Developers, GetOptions());

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void GetBranches()
        {
            long startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Options_AZDO_TFS options = GetOptions();
            options.OrientOutputVertically = true;// This sheet works better vertically.
            MessageBox.Show("TBD - Not Implemented Yet");

            //ProcessCreateWorkSheetTeamProjectCollection(CreateWS_VCS_Branches, teTeamProjectCollection.Text, options);

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void btnGet_ChangeSetInfo_Click(object sender, RoutedEventArgs e)
        {
            long startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            //try
            //{
            //    SpeedUpStart();

            //    int changeSetID = int.Parse(teChangeSetID.Text);

            //    CreateWS_VCS_ChangeSetInfo(changeSetID, cbeChangeSetSections.Text, GetOptions());
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.ToString());
            //}
            //finally
            //{
            //    SpeedUpEnd();
            //}

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void btnGet_TemplateType_Click(object sender, RoutedEventArgs e)
        { 
            long startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            RequestHandlers.ProcessCreateWorkSheet(CreateWS_TP_TemplateType, GetOptions());

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void Get_TPC_BuildDefinitions()
        {
            long startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            RequestHandlers.ProcessCreateWorkSheet(CreateWS_All_TPC_BuildDefinitions, GetOptions());

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void Get_TPC_Areas()
        {
            long startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            RequestHandlers.ProcessCreateWorkSheet(CreateWS_All_TPC_Areas, GetOptions());

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void Get_TPC_Info()
        {
            long startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Options_AZDO_TFS options = GetOptions();

            try
            {
                RequestHandlers.SpeedUpStart();

                //// Get the Team Project Collections

                //ReadOnlyCollection<CatalogNode> projectCollectionNodes = VNCTFS.Helper.Get_TeamProjectCollectionNodes(AzureDevOpsExplorer.Presentation.Views.Server.ConfigurationServer);

                var tpcCatalogNode = AzureDevOpsExplorer.Presentation.Views.Server.TfsTeamProjectCollection.CatalogNode;

                CreateWS_TPC_Info(tpcCatalogNode, AzureDevOpsExplorer.Presentation.Views.Server.TfsTeamProjectCollection, false, options);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                RequestHandlers.SpeedUpEnd();
            }

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void Get_TPC_LastChangeset()
        {
            long startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            RequestHandlers.ProcessCreateWorkSheet((options) => Worksheet_Output.CreateWS_All_TPC_LastChangeset(options, AzureDevOpsExplorer.Presentation.Views.Server.VersionControlServer), GetOptions());

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void Get_TPC_WorkItemActivity()
        {
            long startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            RequestHandlers.ProcessCreateWorkSheet(CreateWS_All_TPC_WorkItemActivity, GetOptions());

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void Get_TPC_Members()
        {
            long startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            try
            {
                RequestHandlers.SpeedUpStart();

                string tpcName = AzureDevOpsExplorer.Presentation.Views.Server.VersionControlServer.TeamProjectCollection.Name;
                tpcName = tpcName.Substring(tpcName.IndexOf("\\") + 1);

                CreateWS_TPC_Members(tpcName, GetOptions());
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                RequestHandlers.SpeedUpEnd();
            }

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void Get_TPC_Shelfsets()
        {
            long startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Options_AZDO_TFS options = GetOptions();
            options.OrientOutputVertically = true;  // This sheet works better vertically.
                                                    //ProcessCreateWorkSheetTeamProjectCollection(CreateWS_ShelveSets, teTeamProjectCollection.Text, options);
            MessageBox.Show("TBD - Not Implemented Yet");

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void Get_TPC_Teams()
        {
            long startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            RequestHandlers.ProcessCreateWorkSheet(CreateWS_All_TPC_Teams, GetOptions());

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void Get_TPC_TestCases()
        {
            long startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            RequestHandlers.ProcessCreateWorkSheet(CreateWS_All_TPC_TestCases, GetOptions());

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void Get_TPC_TestPlans()
        {

            long startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            RequestHandlers.ProcessCreateWorkSheet(CreateWS_All_TPC_TestPlans, GetOptions());

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void Get_TPC_TestSuites()
        {
            RequestHandlers.ProcessCreateWorkSheet(CreateWS_All_TPC_TestSuites, GetOptions());
        }

        private void Get_TPC_WorkItemDetails()
        {
            long startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            RequestHandlers.ProcessCreateWorkSheet(CreateWS_All_TPC_WorkItemDetails, GetOptions());

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void Get_TPC_WorkItemFields()
        {
            long startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            RequestHandlers.ProcessCreateWorkSheet(CreateWS_All_TPC_WorkItemFields, GetOptions());

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void Get_TPC_WorkItemTypes()
        {
            long startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            RequestHandlers.ProcessCreateWorkSheet(CreateWS_All_TPC_WorkItemTypes, GetOptions());

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void Get_TPC_Workspaces()
        {
            long startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            try
            {
                RequestHandlers.SpeedUpStart();

                //CreateWS_Workspaces(Server.ConfigurationServer.Uri.ToString(), teTeamProjectCollection.Text, GetOptions());
                MessageBox.Show("TBD - Not Implemented Yet");   
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                RequestHandlers.SpeedUpEnd();
            }

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void btnSearchForFiles_Click(object sender, RoutedEventArgs e)
        {
            long startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            try
            {
                RequestHandlers.SpeedUpStart();

                SearchForFiles();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                RequestHandlers.SpeedUpEnd();
            }

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void btnUnmergedChanges_Click(object sender, RoutedEventArgs e)
        {
            long startTicks = XlHlp.DisplayInWatchWindow("Begin");

            try
            {
                RequestHandlers.SpeedUpStart();

                MergeCandidate[] mergeCandidates = AzureDevOpsExplorer.Presentation.Views.Server.VersionControlServer.GetMergeCandidates("$/Development", "$/Release", RecursionType.Full);

                foreach (var mergeCandidate in mergeCandidates)
                {
                    if (mergeCandidate.Changeset.Owner == @"DOMAIN\ChuckNorris")
                    {
                        //This is an unmerged changeset commited by Chuck
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                RequestHandlers.SpeedUpEnd();
            }

            XlHlp.DisplayInWatchWindow("End", startTicks);
        }

        #endregion Event Handlers

        #region Utility Routines

        private XlHlp.XlLocation CreateNewWorksheet(string sheetName,
            Options_AZDO_TFS options, [CallerMemberName] string callerName = "")
        {
            Int64 startTicks = Log.APPLICATION($"Begin: sheetName: {sheetName}", Common.LOG_CATEGORY);

            string safeSheetName = XlHlp.SafeSheetName(sheetName);
            Worksheet ws = XlHlp.NewWorksheet(safeSheetName, beforeSheetName: "FIRST");

            XlHlp.XlLocation insertAt = new XlHlp.XlLocation(ws, options.StartingRow, options.StartingColumn, options.OrientOutputVertically);
            XlHlp.AddLabeledInfoX(insertAt.AddRowX(), "Date Run", DateTime.Now.ToString());

            XlHlp.DisplayInWatchWindow("End", startTicks);

            if (!options.FormatSpecs.IsInitialized)
            {
                options.FormatSpecs.Initialize(insertAt);
            }

            // TODO(crhodes)
            // To ensure we have a useful workbook name, check each time we create a new Worksheet.
            // Arguably this should be done somewhere else, but this gives us a better chance of saving intermiediate
            // output as we work on large workbooks.

            string strOutputFile = Globals.ThisAddIn.Application.ActiveWorkbook.Name;

            if (strOutputFile.Equals("Book1") || strOutputFile.Equals("AzureDevOpsExplorer-XXX"))
            {
                // NOTE(crhodes)
                // Looks like this ia a brand now Workbook
                // or we were lazy and didn't change the default name.
                // Get get a better name

                using (System.Windows.Forms.SaveFileDialog saveFileDialog = new System.Windows.Forms.SaveFileDialog())
                {

                    try
                    {
                        saveFileDialog.FileName = "AzureDevOpsExplorer-XXX.xlsx";

                        //saveFileDialog.InitialDirectory = startingFolder;

                        if (System.Windows.Forms.DialogResult.Cancel != saveFileDialog.ShowDialog())
                        {
                            strOutputFile = saveFileDialog.FileName;
                            Globals.ThisAddIn.Application.ActiveWorkbook.SaveAs(strOutputFile);
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);

            return insertAt;
        }

        private static string GetTeamProjectCollectionName(TfsTeamProjectCollection teamProjectCollection)
        {
            long startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            string colName = teamProjectCollection.Name;
            // TFS in cloud send names back with back slashes
            colName = colName.Replace('\\', '/');
            // strip all all but TeamProjectCollection name part
            colName = colName.Substring(colName.LastIndexOf('/') + 1);

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);

            return colName;
        }

        Options_AZDO_TFS GetOptions()
        {
            // HACK(crhodes)
            // What is the proper way of getting data from ViewModel?
            // Should we use IAZDOOptionsViewModel - and put GetOptions in Interface?

            var options = ((OptionsViewModel)azdoOptions.ViewModel).GetOptions();
            return options;
            //return AZDOOptions.GetOptions();
        }

        private void PopulateTeamProjects()
        {
            long startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            try
            {
                var projectList = (from Project prj in AzureDevOpsExplorer.Presentation.Views.Server.WorkItemStore.Projects select prj.Name).ToList();

                var priorStatusBar = Globals.ThisAddIn.Application.StatusBar;

                // NOTE(crhodes)
                // May have already loaded Team Projects for a different collection - clear first.

                ((OptionsViewModel)azdoOptions.ViewModel).TeamProjects.Clear();

                foreach (var item in projectList)
                {
                    ((OptionsViewModel)azdoOptions.ViewModel).TeamProjects.Add(item);
                    Globals.ThisAddIn.Application.StatusBar = $"{item}";
                }

                Globals.ThisAddIn.Application.StatusBar = priorStatusBar;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                // Should we throw?
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private XlHlp.XlLocation SearchForFiles()
        {
            long startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            XlHlp.XlLocation insertAt = default;

            try
            {
                string sheetName = XlHlp.SafeSheetName("Files");
                Worksheet ws = XlHlp.NewWorksheet(sheetName, beforeSheetName: "FIRST");

                int startingRow = 2;
                int startingColumn = 1;

                insertAt = new XlHlp.XlLocation(ws, startingRow, startingColumn);

                // List all of the .sln files.
                //string searchPattern = teFilePattern.Text;

                //ItemSet items = Server.VersionControlServer.GetItems(searchPattern, RecursionType.Full);

                //XlHlp.AddLabeledInfoX(insertAt.AddRowX(), "SearchPattern", searchPattern);
                //XlHlp.AddLabeledInfoX(insertAt.AddRowX(), "Count", items.Items.Count().ToString());

                //foreach (Item item in items.Items)
                //{
                //    insertAt.ClearOffsets();

                //    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), item.ItemType.ToString());
                //    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), item.ServerItem.ToString());

                //    insertAt.IncrementRows();
                //}
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);

            return insertAt;
        }

        #endregion Utility Routines

        private void teChangeSetID_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            long startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            double changeSetID = 0;
            string changeSet = (string)Globals.ThisAddIn.Application.ActiveCell.Value.ToString();

            if (double.TryParse(changeSet, out changeSetID))
            {
                //teChangeSetID.Text = int.Parse(changeSetID.ToString()).ToString();
            }
            else
            {
                MessageBox.Show("ChangeSetID not a number", "Error");
            }

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

    }
}