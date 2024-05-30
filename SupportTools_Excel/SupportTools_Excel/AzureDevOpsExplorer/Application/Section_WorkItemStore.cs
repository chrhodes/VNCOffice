using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;
using System.Windows;

using Microsoft.Office.Interop.Excel;
using Microsoft.TeamFoundation.Server;
using Microsoft.TeamFoundation.WorkItemTracking.Client;

using SupportTools_Excel.AzureDevOpsExplorer.Domain;

using VNC;
using VNC.AddinHelper;

using XlHlp = VNC.AddinHelper.Excel;

namespace SupportTools_Excel.AzureDevOpsExplorer.Application
{
    class Section_WorkItemStore
    {
        internal static XlHlp.XlLocation AddSections(                                                                                                                                                           XlHlp.XlLocation insertAt,
        Options_AZDO_TFS options,
        WorkItemStore workItemStore,
        ICommonStructureService commonStructureService,
        Project project,
        List<string> sectionsToDisplay)
        {
            Int64 startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            XlHlp.DisplayInWatchWindow(insertAt);

            if (sectionsToDisplay.Count != 0)
            {
                try
                {
                    if (insertAt.OrientVertical)
                    {
                        XlHlp.AddSectionInfo(insertAt.AddRow(), "Work Item Store (WIS) Information", "");
                        insertAt.IncrementRows();
                    }
                    else
                    {
                        XlHlp.AddSectionInfo(insertAt.AddRow(), "Work Item Store (WIS) Information", "",
                            orientation: XlOrientation.xlUpward);
                        insertAt.DecrementRows();   // AddRow bumped it.
                        insertAt.IncrementColumns();
                    }

                    if (sectionsToDisplay.Contains("Info"))
                    {
                        insertAt = Add_Info(insertAt, options, workItemStore, project).IncrementPosition(insertAt.OrientVertical);
                        insertAt.IncrementRows();
                        Globals.ThisAddIn.Application.ActiveWorkbook.Save();
                    }

                    if (sectionsToDisplay.Contains("Areas"))
                    {
                        insertAt = Add_TP_Areas(insertAt, options, commonStructureService, project).IncrementPosition(insertAt.OrientVertical);
                        insertAt.IncrementRows();
                        Globals.ThisAddIn.Application.ActiveWorkbook.Save();
                    }

                    if (sectionsToDisplay.Contains("Iterations"))
                    {
                        insertAt = Add_TP_Iterations(insertAt, options, commonStructureService, project).IncrementPosition(insertAt.OrientVertical);
                        insertAt.IncrementRows();
                        Globals.ThisAddIn.Application.ActiveWorkbook.Save();
                    }

                    if (sectionsToDisplay.Contains("Stored Queries"))
                    {
                        insertAt = Add_TP_StoredQueries(insertAt, options, project).IncrementPosition(insertAt.OrientVertical);
                        insertAt.IncrementRows();
                        Globals.ThisAddIn.Application.ActiveWorkbook.Save();
                    }

                    if (sectionsToDisplay.Contains("Work Item Fields"))
                    {
                        insertAt = Add_TP_WorkItemFields(insertAt, options, project).IncrementPosition(insertAt.OrientVertical);
                        insertAt.IncrementRows();
                        Globals.ThisAddIn.Application.ActiveWorkbook.Save();
                    }

                    if (sectionsToDisplay.Contains("Work Item Types"))
                    {
                        insertAt = Add_TP_WorkItemTypes(insertAt, options, workItemStore, project);
                        insertAt.IncrementRows();
                        Globals.ThisAddIn.Application.ActiveWorkbook.Save();
                    }

                    if (sectionsToDisplay.Contains("Work Item Activity"))
                    {
                        insertAt = Add_TP_WorkItemActivity(insertAt, options, workItemStore, project);
                        insertAt.IncrementRows();
                        Globals.ThisAddIn.Application.ActiveWorkbook.Save();
                    }

                    if (sectionsToDisplay.Contains("Work Item Details"))
                    {
                        insertAt = Add_TP_WorkItemDetails(insertAt, options, workItemStore, project).IncrementPosition(insertAt.OrientVertical);
                        insertAt.IncrementRows();
                        Globals.ThisAddIn.Application.ActiveWorkbook.Save();
                    }

                    if (sectionsToDisplay.Contains("Work Item Field Mapping"))
                    {
                        insertAt = Add_TP_FieldMapping(insertAt, options, project).IncrementPosition(insertAt.OrientVertical);
                        insertAt.IncrementRows();
                        Globals.ThisAddIn.Application.ActiveWorkbook.Save();
                    }

                    // Put Work Item Categories last as it has odd output.  Too lazy to fix indent.
                    if (sectionsToDisplay.Contains("Work Item Categories"))
                    {
                        insertAt = Add_TP_WorkItemCategories(insertAt, options, workItemStore, project).IncrementPosition(insertAt.OrientVertical);
                        insertAt.IncrementRows();
                        Globals.ThisAddIn.Application.ActiveWorkbook.Save();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }

            XlHlp.DisplayInWatchWindow(insertAt, startTicks, "End");

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);

            return insertAt;
        }

        internal static XlHlp.XlLocation Add_Info(
            XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options,
            WorkItemStore workItemStore,
            Project project = null)
        {
            Int64 startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);
            XlHlp.DisplayInWatchWindow(insertAt);

            try
            {
                insertAt.MarkStart(XlHlp.MarkType.Group);

                XlHlp.AddLabeledInfoX(insertAt.AddRowX(2), $"BypassRules: {workItemStore.BypassRules}", "");
                XlHlp.AddLabeledInfoX(insertAt.AddRowX(2), $"CallingProcessIdentity: {workItemStore.CallingProcessIdentity}", "");
                XlHlp.AddLabeledInfoX(insertAt.AddRowX(2), $"CultureInfo: {workItemStore.CultureInfo}", "");
                XlHlp.AddLabeledInfoX(insertAt.AddRowX(2), $"Diagnostics.RoundTripCount: {workItemStore.Diagnostics.RoundTripCount}", "");
                XlHlp.AddLabeledInfoX(insertAt.AddRowX(2), $"Diagnostics.RoundTripTime: {workItemStore.Diagnostics.RoundTripTime}", "");
                XlHlp.AddLabeledInfoX(insertAt.AddRowX(2), $"IsIdentityFieldSupported: {workItemStore.IsIdentityFieldSupported}", "");
                XlHlp.AddLabeledInfoX(insertAt.AddRowX(2), $"Projects.Count: {workItemStore.Projects.Count}", "");
                XlHlp.AddLabeledInfoX(insertAt.AddRowX(2), $"ServerInfo.Features.Count: {workItemStore.ServerInfo.Features.Count()}", "");
                XlHlp.AddLabeledInfoX(insertAt.AddRowX(2), $"TimeZone: {workItemStore.TimeZone}", "");
                XlHlp.AddLabeledInfoX(insertAt.AddRowX(2), $"Projects.Count: {workItemStore.Projects.Count}", "");
                XlHlp.AddLabeledInfoX(insertAt.AddRowX(2), $"ClientService.WorkItemServerVersion: {workItemStore.ClientService.WorkItemServerVersion}", "");

                if (project is { })
                {
                    XlHlp.AddLabeledInfoX(insertAt.AddRowX(2), $"Project.Name: {project.Name}", "");
                    XlHlp.AddLabeledInfoX(insertAt.AddRowX(2), $"Project.Id: {project.Id}", "");
                }

                insertAt.MarkEnd(XlHlp.MarkType.Group);

                insertAt.Group(insertAt.OrientVertical, hide: true);

                insertAt.EndSectionAndSetNextLocation(insertAt.OrientVertical);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            XlHlp.DisplayInWatchWindow(insertAt, startTicks, "End");
            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);

            return insertAt;
        }

        //internal static XlHlp.XlLocation Add_Query(
        //    XlHlp.XlLocation insertAt,
        //    Options_AZDO_TFS options,
        //    string tableName)
        //{
        //    long startTicks = XlHlp.DisplayInWatchWindow(insertAt);

        //    DisplayQueryInfo(insertAt, options);

        //    try
        //    {
        //        WorkItemCollection queryResults = Presentation.Views.Server.WorkItemStore.Query(options.WorkItemQuerySpec.Query);

        //        int totalItems;

        //        if ((totalItems = queryResults.Count) > 0)
        //        {
        //            XlHlp.AddLabeledInfo(insertAt.AddRow(), "Matches", totalItems.ToString());

        //            insertAt.MarkStart(XlHlp.MarkType.GroupTable);

        //            Header_WorkItemStore.Add_TP_WorkItemDetails(insertAt);

        //            Body_WorkItemStore.Add_TP_WorkItemDetails(insertAt, options, queryResults);

        //            insertAt.MarkEnd(XlHlp.MarkType.GroupTable, tableName);

        //            insertAt.Group(insertAt.OrientVertical, hide: true);
        //        }
        //    }
        //    catch (ValidationException cex)
        //    {
        //        XlHlp.AddLabeledInfo(insertAt.AddRow(), "Matches", cex.ToString());
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.ToString());
        //        throw;
        //    }

        //    if (!insertAt.OrientVertical)
        //    {
        //        // Skip past the info just added.
        //        insertAt.SetLocation(insertAt.RowStart, insertAt.MarkEndColumn + 1);
        //    }

        //    XlHlp.DisplayInWatchWindow(insertAt, startTicks, "End");

        //    return insertAt;
        //}

        internal static void DisplayQueryInfo(XlHlp.XlLocation insertAt, Options_AZDO_TFS options)
        {
            Int64 startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            XlHlp.AddLabeledInfo(insertAt.AddRow(2), "Name", options.WorkItemQuerySpec.Name);

            XlHlp.AddLabeledInfo(insertAt.AddRow(2), "QueryWithTokens", options.WorkItemQuerySpec.QueryWithTokens);

            XlHlp.AddLabeledInfo(insertAt.AddRow(2), "Query", options.WorkItemQuerySpec.Query);

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);
        }

        internal static XlHlp.XlLocation Add_TP_AreaCheck(
            XlHlp.XlLocation insertAt,
            Project project,
            string areasToCheck)
        {
            Int64 startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            try
            {
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), project.Name);
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), project.AreaRootNodes.Count.ToString());

                //XlHlp.AddTitledInfo(insertAt.AddRow(), "Areas", project.AreaRootNodes.Count.ToString());

                //insertAt.MarkStart(XlHlp.MarkType.GroupTable);

                char[] splitChars = { ',' };

                if (project.AreaRootNodes.Count == 0)
                {
                    //XlHlp.AddContentToCell(insertAt.AddRow(), "None");
                }
                else
                {
                    foreach (string area in areasToCheck.Split(splitChars, StringSplitOptions.None))
                    {
                        try
                        {
                            var result = project.AreaRootNodes[area];
                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), area);
                        }
                        catch (Exception)
                        {
                            // No such area.  Need to learn how to use .Contains()
                            //throw;
                        }
                    }

                }

                insertAt.AddRowX();

                //insertAt.MarkEnd(XlHlp.MarkType.GroupTable);

                //insertAt.Group(insertAt.OrientVertical, hide: true);

                insertAt.EndSectionAndSetNextLocation(insertAt.OrientVertical);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);

            return insertAt;
        }

        internal static XlHlp.XlLocation Add_TP_Areas(
            XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options,
            ICommonStructureService commonStructureService,
            Project project)
        {
            Int64 startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);
            XlHlp.DisplayInWatchWindow(insertAt);

            try
            {
                int areaNodesCount = project.AreaRootNodes.Count;

                if (insertAt.OrientVertical)
                {
                    XlHlp.AddSectionInfo(insertAt.AddRow(), "Areas", $"{ areaNodesCount }");
                }
                else
                {
                    XlHlp.AddSectionInfo(insertAt.AddRow(), "Areas", $"{ areaNodesCount }",
                        orientation: XlOrientation.xlUpward);
                    insertAt.IncrementColumns();
                }

                insertAt.MarkStart(XlHlp.MarkType.GroupTable);

                Header_WorkItemStore.Add_TP_Areas(insertAt);

                insertAt = Body_WorkItemStore.Add_TP_Areas(insertAt, options, commonStructureService, project);

                insertAt.MarkEnd(XlHlp.MarkType.GroupTable);

                insertAt.Group(insertAt.OrientVertical);

                if (!insertAt.OrientVertical)
                {
                    // Skip past the info just added.  Use Group end because we have indented
                    insertAt.SetLocation(insertAt.RowStart, insertAt.GroupEndColumn + 1);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            XlHlp.DisplayInWatchWindow(insertAt, startTicks, "End");
            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);

            return insertAt;
        }

        internal static XlHlp.XlLocation Add_TP_FieldMapping(
            XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options,
            Project project)
        {
            Int64 startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);
            XlHlp.DisplayInWatchWindow(insertAt);

            try
            {
                if (insertAt.OrientVertical)
                {
                    XlHlp.AddSectionInfo(insertAt.AddRow(), "WorkItem Field Mappings", project.WorkItemTypes.Count.ToString());
                }
                else
                {
                    XlHlp.AddSectionInfo(insertAt.AddRow(), "WorkItem Field Mappings", project.WorkItemTypes.Count.ToString(),
                        orientation: XlOrientation.xlUpward);
                    insertAt.IncrementColumns();
                }

                insertAt.MarkStart(XlHlp.MarkType.GroupTable);

                Header_WorkItemStore.Add_TP_FieldMapping(insertAt);

                Body_WorkItemStore.Add_TP_FieldMapping(insertAt, options, project);

                insertAt.MarkEnd(XlHlp.MarkType.GroupTable, string.Format("tblWITFM_{0}", project.Name));

                insertAt.Group(insertAt.OrientVertical);

                insertAt.EndSectionAndSetNextLocation(insertAt.OrientVertical);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            XlHlp.DisplayInWatchWindow(insertAt, startTicks, "End");
            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);

            return insertAt;
        }

        internal static XlHlp.XlLocation Add_TP_Iterations(
            XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options,
            ICommonStructureService commonStructureService,
            Project project)
        {
            Int64 startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);
            XlHlp.DisplayInWatchWindow(insertAt);

            try
            {
                int iterationNodes = project.IterationRootNodes.Count;

                if (insertAt.OrientVertical)
                {
                    XlHlp.AddSectionInfo(insertAt.AddRow(), "Iterations", iterationNodes.ToString());
                }
                else
                {
                    XlHlp.AddSectionInfo(insertAt.AddRow(), "Iterations", iterationNodes.ToString(),
                        orientation: XlOrientation.xlUpward);
                    insertAt.IncrementColumns();
                }

                insertAt.MarkStart(XlHlp.MarkType.GroupTable);

                Header_WorkItemStore.Add_TP_Iterations(insertAt);

                insertAt = Body_WorkItemStore.Add_TP_Iterations(insertAt, options, commonStructureService, project);

                insertAt.MarkEnd(XlHlp.MarkType.GroupTable);

                insertAt.Group(insertAt.OrientVertical);

                if (!insertAt.OrientVertical)
                {
                    // Skip past the info just added.  Use GroupEnd because we have indented to show hierarchy.
                    insertAt.SetLocation(insertAt.RowStart, insertAt.GroupEndColumn + 1);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            XlHlp.DisplayInWatchWindow(insertAt, startTicks, "End");
            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);

            return insertAt;
        }

        //internal static XlHlp.XlLocation Add_TP_Query(
        //    XlHlp.XlLocation insertAt,
        //    Options_AZDO_TFS options,
        //    Project project)
        //{
        //    long startTicks = XlHlp.DisplayInWatchWindow(insertAt);

        //    options.WorkItemQuerySpec.ReplaceQueryTokens(options);
        //    string tableName = $"tblQuery_{project.Name}";

        //    insertAt = Add_Query(insertAt, options, tableName);

        //    XlHlp.DisplayInWatchWindow(insertAt, startTicks, "End");

        //    return insertAt;
        //}

        internal static XlHlp.XlLocation Add_TP_StoredQueries(
            XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options,
            Project project)
        {
            Int64 startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);
            XlHlp.DisplayInWatchWindow(insertAt);

            try
            {
                if (insertAt.OrientVertical)
                {
                    XlHlp.AddSectionInfo(insertAt.AddRow(), 
                        "StoredQueries", "");
                }
                else
                {
                    XlHlp.AddSectionInfo(insertAt.AddRow(), 
                        "StoredQueries", "", 
                        orientation: XlOrientation.xlUpward);
                    insertAt.IncrementColumns();
                }

                insertAt.MarkStart(XlHlp.MarkType.Group);

                insertAt = AddQueryNodes(insertAt, project.QueryHierarchy);

                insertAt.MarkEnd(XlHlp.MarkType.Group);

                insertAt.Group(insertAt.OrientVertical);

                if (!insertAt.OrientVertical)
                {
                    // Skip past the info just added.  Use Group end because we have indented
                    insertAt.SetLocation(insertAt.RowStart, insertAt.GroupEndColumn + 1);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            XlHlp.DisplayInWatchWindow(insertAt, startTicks, "End");
            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);

            return insertAt;
        }

        internal static XlHlp.XlLocation Add_TP_WorkItemCategories(
            XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options,
            WorkItemStore workItemStore,
            Project project)
        {
            Int64 startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);
            XlHlp.DisplayInWatchWindow(insertAt);

            try
            {
                Microsoft.TeamFoundation.WorkItemTracking.Client.CategoryCollection categories = 
                    workItemStore.Projects[project.Name].Categories;

                if (insertAt.OrientVertical)
                {
                    XlHlp.AddSectionInfo(insertAt.AddRow(), 
                        "WorkItem Categories", $"{categories.Count}");
                }
                else
                {
                    XlHlp.AddSectionInfo(insertAt.AddRow(), 
                        "WorkItem Categories", $"{categories.Count}",
                        orientation: XlOrientation.xlUpward);
                    insertAt.IncrementColumns();
                }

                insertAt.MarkStart(XlHlp.MarkType.GroupTable);

                foreach (Category category in categories.OrderBy(nnn => nnn.Name))
                {
                    insertAt.ClearOffsets();

                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), category.Name);
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), category.DefaultWorkItemType.Name);

                    foreach (WorkItemType wit in category.WorkItemTypes.OrderBy(nnn => nnn.Name))
                    {
                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), wit.Name);
                    }

                    insertAt.IncrementRows();

                    AZDOHelper.ProcessLoopDelay(options);
                }

                insertAt.MarkEnd(XlHlp.MarkType.GroupTable, string.Format("tblWIC_{0}", project.Name));

                insertAt.Group(insertAt.OrientVertical, hide: true);

                insertAt.EndSectionAndSetNextLocation(insertAt.OrientVertical);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            XlHlp.DisplayInWatchWindow(insertAt, startTicks, "End");
            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);

            return insertAt;
        }

        internal static XlHlp.XlLocation Add_TP_WorkItemDetails(
            XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options,
            WorkItemStore workItemStore,
            Project project)
        {
            Int64 startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);
            XlHlp.DisplayInWatchWindow(insertAt);

            WorkItemCollection queryResults;

            try
            {
                options.WorkItemQuerySpec.ReplaceQueryTokens(options, project.Name);

                if (insertAt.OrientVertical)
                {
                    //XlHlp.AddSectionInfo(insertAt.AddRow(), options.WorkItemQuerySpec.Query);
                    XlHlp.AddSectionInfo(insertAt.AddRow(),
                        "WorkItem Details", options.WorkItemQuerySpec.Query);
                }
                else
                {
                    XlHlp.AddSectionInfo(insertAt.AddRow(), 
                        "WorkItem Details", options.WorkItemQuerySpec.Query,
                        orientation: XlOrientation.xlUpward);
                    insertAt.IncrementColumns();
                }

                insertAt.MarkStart(XlHlp.MarkType.GroupTable);

                Header_WorkItemStore.Add_TP_WorkItemDetails(insertAt, options);

                queryResults = workItemStore.Query(options.WorkItemQuerySpec.Query);

                Body_WorkItemStore.Add_TP_WorkItemDetails(insertAt,
                        options, queryResults);

                insertAt.MarkEnd(XlHlp.MarkType.GroupTable, string.Format("tblWID_{0}", project.Name));

                insertAt.Group(insertAt.OrientVertical);

                insertAt.EndSectionAndSetNextLocation(insertAt.OrientVertical);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            XlHlp.DisplayInWatchWindow(insertAt, startTicks, "End");
            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);

            return insertAt;
        }

        internal static XlHlp.XlLocation Add_TP_WorkItemFields(
            XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options,
            Project project)
        {
            Int64 startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);
            XlHlp.DisplayInWatchWindow(insertAt);

            try
            {
                // Save the location of the title so we can update later after have traversed all items.

                Range rngTitle = insertAt.GetCurrentRange();

                if (insertAt.OrientVertical)
                {
                    XlHlp.AddSectionInfo(insertAt.AddRow(), "WorkItem Fields", project.WorkItemTypes.Count.ToString());
                }
                else
                {
                    XlHlp.AddSectionInfo(insertAt.AddRow(), "WorkItem Fields", project.WorkItemTypes.Count.ToString(),
                        orientation: XlOrientation.xlUpward);
                    insertAt.IncrementColumns();
                }

                insertAt.MarkStart(XlHlp.MarkType.GroupTable);

                Header_WorkItemStore.Add_TP_WorkItemFields(insertAt);

                Body_WorkItemStore.Add_TP_WorkItemFields(insertAt, options, project);

                insertAt.MarkEnd(XlHlp.MarkType.GroupTable, string.Format("tblWIF_{0}", project.Name));

                insertAt.Group(insertAt.OrientVertical);

                insertAt.EndSectionAndSetNextLocation(insertAt.OrientVertical);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            XlHlp.DisplayInWatchWindow(insertAt, startTicks, "End");
            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);

            return insertAt;
        }

        internal static XlHlp.XlLocation Add_TP_WorkItemTypes(
            XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options,
            WorkItemStore workItemStore,
            Project project)
        {
            Int64 startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);
            XlHlp.DisplayInWatchWindow(insertAt);

            try
            {
                // Save the location of the title so we can update later after have traversed all items.

                Range rngTitle = insertAt.GetCurrentRange();

                if (insertAt.OrientVertical)
                {
                    XlHlp.AddSectionInfo(insertAt.AddRow(), "WorkItem Types", project.WorkItemTypes.Count.ToString());
                }
                else
                {
                    XlHlp.AddSectionInfo(insertAt.AddRow(), "WorkItem Types", project.WorkItemTypes.Count.ToString(),
                        orientation: XlOrientation.xlUpward);
                    insertAt.IncrementColumns();
                }

                insertAt.MarkStart(XlHlp.MarkType.GroupTable);

                Header_WorkItemStore.Add_TP_WorkItemTypes(insertAt);

                DateTime maxLastCreatedDate = DateTime.MinValue;
                DateTime maxLastChangedDate = DateTime.MinValue;
                DateTime maxLastRevisedDate = DateTime.MinValue;

                Body_WorkItemStore.Add_TP_WorkItemTypes(insertAt, options, workItemStore, project, out maxLastCreatedDate, out maxLastChangedDate, out maxLastRevisedDate);

                insertAt.MarkEnd(XlHlp.MarkType.GroupTable, string.Format("tblWIT_{0}", project.Name));

                insertAt.Group(insertAt.OrientVertical);

                // Add the date information
                rngTitle.Offset[0, 2].Value = maxLastCreatedDate.ToString();
                rngTitle.Offset[0, 3].Value = maxLastChangedDate.ToString();
                rngTitle.Offset[0, 4].Value = maxLastRevisedDate.ToString();

                insertAt.EndSectionAndSetNextLocation(insertAt.OrientVertical);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            XlHlp.DisplayInWatchWindow(insertAt, startTicks, "End");
            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);

            return insertAt;
        }

        internal static XlHlp.XlLocation Add_TP_WorkItemActivity(
            XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options,
            WorkItemStore workItemStore,
            Project project)
        {
            Int64 startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);
            XlHlp.DisplayInWatchWindow(insertAt);

            try
            {
                // Save the location of the title so we can update later after have traversed all items.

                Range rngTitle = insertAt.GetCurrentRange();

                if (insertAt.OrientVertical)
                {
                    XlHlp.AddSectionInfo(insertAt.AddRow(), "WorkItem Activity", project.WorkItemTypes.Count.ToString());
                }
                else
                {
                    XlHlp.AddSectionInfo(insertAt.AddRow(), "WorkItem Activity", project.WorkItemTypes.Count.ToString(),
                        orientation: XlOrientation.xlUpward);
                    insertAt.IncrementColumns();
                }

                insertAt.MarkStart(XlHlp.MarkType.GroupTable);

                Header_WorkItemStore.Add_TP_WorkItemActivity(insertAt);

                DateTime maxLastCreatedDate = DateTime.MinValue;
                DateTime maxLastChangedDate = DateTime.MinValue;
                DateTime maxLastRevisedDate = DateTime.MinValue;

                Body_WorkItemStore.Add_TP_WorkItemActivity(insertAt, options, workItemStore, project, out maxLastCreatedDate, out maxLastChangedDate, out maxLastRevisedDate);

                insertAt.MarkEnd(XlHlp.MarkType.GroupTable, string.Format("tblWIA_{0}", project.Name));

                insertAt.Group(insertAt.OrientVertical);

                // Add the date information
                rngTitle.Offset[0, 2].Value = maxLastCreatedDate.ToString();
                rngTitle.Offset[0, 3].Value = maxLastChangedDate.ToString();
                rngTitle.Offset[0, 4].Value = maxLastRevisedDate.ToString();

                insertAt.EndSectionAndSetNextLocation(insertAt.OrientVertical);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            XlHlp.DisplayInWatchWindow(insertAt, startTicks, "End");
            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);

            return insertAt;
        }

        //private static XlHlp.XlLocation Add_TP_WorkItemTypesXML(
        //    XlHlp.XlLocation insertAt, 
        //    Options_AZDO_TFS options, 
        //    WorkItemStore workItemStore, 
        //    Project project)
        //{
        //    long startTicks = Log.Trace("Enter", Common.LOG_CATEGORY);

        //    try
        //    {
        //        //// Save the location of the title so we can update later after have traversed all items.

        //        //Range rngTitle = insertAt.GetCurrentRange();

        //        //if (insertAt.OrientVertical)
        //        //{
        //        //    XlHlp.AddSectionInfo(insertAt.AddRow(), "WorkItem Types", project.WorkItemTypes.Count.ToString());
        //        //}
        //        //else
        //        //{
        //        //    XlHlp.AddSectionInfo(insertAt.AddRow(), "WorkItem Types", project.WorkItemTypes.Count.ToString(),
        //        //        orientation: XlOrientation.xlUpward);
        //        //    insertAt.IncrementColumns();
        //        //}

        //        //insertAt.MarkStart(XlHlp.MarkType.GroupTable);

        //        //Header_WorkItemStore.Add_TP_WorkItemTypes(insertAt);

        //        //DateTime maxLastCreatedDate = DateTime.MinValue;
        //        //DateTime maxLastChangedDate = DateTime.MinValue;
        //        //DateTime maxLastRevisedDate = DateTime.MinValue;

        //        Body_WorkItemStore.Get_TP_WorkItemTypesXML(insertAt, options, workItemStore, project);

        //        //insertAt.MarkEnd(XlHlp.MarkType.GroupTable, string.Format("tblWIT_{0}", project.Name));

        //        //insertAt.Group(insertAt.OrientVertical);

        //        //// Add the date information
        //        //rngTitle.Offset[0, 2].Value = maxLastCreatedDate.ToString();
        //        //rngTitle.Offset[0, 3].Value = maxLastChangedDate.ToString();
        //        //rngTitle.Offset[0, 4].Value = maxLastRevisedDate.ToString();

        //        //insertAt.EndSectionAndSetNextLocation(insertAt.OrientVertical);
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.ToString());
        //    }

        //    XlHlp.DisplayInWatchWindow(insertAt, startTicks, "End");
        //    Log.Trace("Exit", Common.LOG_CATEGORY, startTicks);

        //    return insertAt;
        //}

        internal static XlHlp.XlLocation Test_WorkItem_Fields(
            XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options,
            WorkItem workItem,
            WorkItemActionRequest request)
        {
            Int64 startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);
            XlHlp.DisplayInWatchWindow(insertAt);

            StringBuilder sb = new StringBuilder();

            sb.AppendLine("This is > 1 & < '2' but > \"1.5\" <div> <br> \"<div>\" '<br>' Oh My");
            sb.AppendLine("&lt; (<) &gt; (>) &amp; (&) &apos &quot; (\") &#39; (')");
            sb.AppendLine("This is > 2 & < '3' but > \"2.5\" <div> <br> \"<div>\" '<br>' Oh My");
            sb.AppendLine("&lt; (<) &gt; (>) &amp; (&) &apos &quot; (\") &#39; (')");
            sb.AppendLine("This is > 3 & < '4' but > \"3.5\" <div> <br> \"<div>\" '<br>' Oh My");

            try
            {
                if (insertAt.OrientVertical)
                {
                    XlHlp.AddSectionInfo(insertAt.AddRow(), "WorkItem Fields", workItem.Id.ToString(), columnWidth: 25);
                }
                else
                {
                    XlHlp.AddSectionInfo(insertAt.AddRow(), "WorkItem Fields", workItem.Id.ToString(),
                        orientation: XlOrientation.xlUpward);
                    insertAt.IncrementColumns();
                }

                insertAt.IncrementRows();

                insertAt.MarkStart(XlHlp.MarkType.GroupTable);

                if (request.RetrieveAllWorkItemFieldData)
                {
                    Header_WorkItemStore.Add_TP_WorkItemFieldValues(insertAt);

                    workItem.Open();

                    // HACK(crhodes)
                    // Put better test data in original fields.

                    workItem.Fields["Custom.TestCase.TestData"].Value = sb.ToString();
                    workItem.Fields["Custom.TestCase.ExecutionSteps"].Value = sb.ToString();
                    workItem.Fields["Custom.TestCase.SetupSteps"].Value = sb.ToString();

                    foreach (Field field in workItem.Fields)
                    {
                        insertAt.ClearOffsets();

                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ field.Id }");
                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ field.Name }");
                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ field.ReferenceName }");
                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $">{ field.OriginalValue }<");
                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $">{ field.Value }<");
                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $">{ field.FieldDefinition.FieldType }<");
                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $">{ field.FieldDefinition.SystemType }<");

                        // HACK(crhodes)
                        // 

                        if (field.Name.Equals("STS_Data"))
                        {

                            string htmlEncodeded = HttpUtility.HtmlEncode(workItem.Fields["Custom.TestCase.TestData"].Value);

                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $">{ htmlEncodeded }<");

                            field.Value = htmlEncodeded;
                        }

                        if (field.Name.Equals("STS_Execution"))
                        {
                            string htmlEncodeded = HttpUtility.HtmlEncode(workItem.Fields["Custom.TestCase.ExecutionSteps"].Value);

                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $">{ htmlEncodeded }<");

                            field.Value = htmlEncodeded;
                        }

                        if (field.Name.Equals("STS_Setup"))
                        {
                            string htmlEncodeded = HttpUtility.HtmlEncode(workItem.Fields["Custom.TestCase.SetupSteps"].Value);

                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $">{ htmlEncodeded }<");

                            field.Value = htmlEncodeded;
                        }

                        insertAt.IncrementRows();
                    }
                }
                else
                {
                    Header_WorkItemStore.Add_TP_WorkItemFieldValues2(insertAt);

                    foreach (Field field in workItem.Fields)
                    {
                        if (request.WorkItemFields.Contains(field.ReferenceName))
                        {
                            insertAt.ClearOffsets();

                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ workItem.Project.Name }");
                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ workItem.Id}");
                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ workItem.Type.Name}");


                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ field.Id }");
                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ field.Name }");
                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ field.ReferenceName }");
                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $">{ field.OriginalValue }<");
                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $">{ field.Value }<");
                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $">{ field.FieldDefinition.FieldType }<");
                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $">{ field.FieldDefinition.SystemType }<");

                            insertAt.IncrementRows();
                        }
                    }
                }

                workItem.Save();

                insertAt.MarkEnd(XlHlp.MarkType.GroupTable);

                insertAt.Group(insertAt.OrientVertical);

                insertAt.EndSectionAndSetNextLocation(insertAt.OrientVertical);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            XlHlp.DisplayInWatchWindow(insertAt, startTicks, "End");
            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);

            return insertAt;
        }

        internal static XlHlp.XlLocation Add_WorkItem_Fields(
            XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options,
            WorkItem workItem,
            WorkItemActionRequest request)
        {
            Int64 startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);
            XlHlp.DisplayInWatchWindow(insertAt);

            try
            {
                if (insertAt.OrientVertical)
                {
                    XlHlp.AddSectionInfo(insertAt.AddRow(), "WorkItem Fields", workItem.Id.ToString(), columnWidth: 25);
                }
                else
                {
                    XlHlp.AddSectionInfo(insertAt.AddRow(), "WorkItem Fields", workItem.Id.ToString(),
                        orientation: XlOrientation.xlUpward);
                    insertAt.IncrementColumns();
                }

                insertAt.IncrementRows();

                insertAt.MarkStart(XlHlp.MarkType.GroupTable);

                if (request.RetrieveAllWorkItemFieldData)
                {
                    Header_WorkItemStore.Add_TP_WorkItemFieldValues(insertAt);

                    foreach (Field field in workItem.Fields)
                    {
                        insertAt.ClearOffsets();

                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ field.Id }");
                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ field.Name }");
                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ field.ReferenceName }");
                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $">{ field.OriginalValue }<");
                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $">{ field.Value }<");
                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $">{ field.FieldDefinition.FieldType }<");
                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $">{ field.FieldDefinition.SystemType }<");

                        insertAt.IncrementRows();
                    }
                }
                else
                {
                    Header_WorkItemStore.Add_TP_WorkItemFieldValues2(insertAt);

                    foreach (Field field in workItem.Fields)
                    {
                        if (request.WorkItemFields.Contains(field.ReferenceName))
                        {
                            insertAt.ClearOffsets();

                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ workItem.Project.Name }");
                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ workItem.Id}");
                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ workItem.Type.Name}");


                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ field.Id }");
                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ field.Name }");
                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ field.ReferenceName }");
                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $">{ field.OriginalValue }<");
                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $">{ field.Value }<");
                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $">{ field.FieldDefinition.FieldType }<");
                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $">{ field.FieldDefinition.SystemType }<");

                            insertAt.IncrementRows();
                        }
                    }
                }

                insertAt.MarkEnd(XlHlp.MarkType.GroupTable);

                insertAt.Group(insertAt.OrientVertical);

                insertAt.EndSectionAndSetNextLocation(insertAt.OrientVertical);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            XlHlp.DisplayInWatchWindow(insertAt, startTicks, "End");
            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);

            return insertAt;
        }

        internal static XlHlp.XlLocation Add_WorkItem_Info(
            XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options,
            WorkItem workItem)
        {
            Int64 startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);
            XlHlp.DisplayInWatchWindow(insertAt);

            try
            {
                if (insertAt.OrientVertical)
                {
                    XlHlp.AddSectionInfo(insertAt.AddRow(), "WorkItem Info", workItem.Id.ToString(), columnWidth: 25);
                }
                else
                {
                    XlHlp.AddSectionInfo(insertAt.AddRow(), "WorkItem Info", workItem.Id.ToString(),
                        orientation: XlOrientation.xlUpward);
                    insertAt.IncrementColumns();
                }

                insertAt.MarkStart(XlHlp.MarkType.Group);

                CellFormatSpecification redContent = options.FormatSpecs.RedContent;
                CellFormatSpecification dateLabel = options.FormatSpecs.DateLabel;
                CellFormatSpecification dateContent = options.FormatSpecs.DateContent;

                XlHlp.AddLabeledInfo(insertAt.AddRow(), "Team Project", $"{ workItem.Project.Name }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "Id", $"{ workItem.Id }", contentFormat: redContent);
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "Title", $"{ workItem.Title }");

                XlHlp.AddLabeledInfo(insertAt.AddRow(), "Type", $"{ workItem.Type.Name }");

                XlHlp.AddLabeledInfo(insertAt.AddRow(), "AreaID", $"{ workItem.AreaId }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "AreaPath", $"{ workItem.AreaPath }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "AttachedFileCount", $"{ workItem.AttachedFileCount }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "Attachments", $"{ workItem.Attachments.Count }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "AuthorizedDate", $"{ workItem.AuthorizedDate }", labelFormat: dateLabel, contentFormat: dateContent);

                XlHlp.AddLabeledInfo(insertAt.AddRow(), "Changed By", $"{ workItem.ChangedBy }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "Changed Date", $"{ workItem.ChangedDate }", labelFormat: dateLabel, contentFormat: dateContent);
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "Created By", $"{ workItem.CreatedBy }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "Created Date", $"{ workItem.CreatedDate }", labelFormat: dateLabel, contentFormat: dateContent);

                XlHlp.AddLabeledInfo(insertAt.AddRow(), "Description", $"{ workItem.Description }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "ExternalLinkCount", $"{ workItem.ExternalLinkCount }");

                XlHlp.AddLabeledInfo(insertAt.AddRow(), "Fields", $"{ workItem.Fields.Count }");

                XlHlp.AddLabeledInfo(insertAt.AddRow(), "History", $"{ workItem.History }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "HyperLinkCount", $"{ workItem.HyperLinkCount }");

                XlHlp.AddLabeledInfo(insertAt.AddRow(), "IsDirty", $"{ workItem.IsDirty }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "IsNew", $"{ workItem.IsNew }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "IsOpen", $"{ workItem.IsOpen }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "IsPartialOpen", $"{ workItem.IsPartialOpen }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "IsReadOnly", $"{ workItem.IsReadOnly }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "IsReadOnlyOpen", $"{ workItem.IsReadOnlyOpen }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "IterationId", $"{ workItem.IterationId }"); ;
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "IterationPath", $"{ workItem.IterationPath }"); ;

                XlHlp.AddLabeledInfo(insertAt.AddRow(), "Links", $"{ workItem.Links.Count }");

                XlHlp.AddLabeledInfo(insertAt.AddRow(), "NodeName", $"{ workItem.NodeName }");

                XlHlp.AddLabeledInfo(insertAt.AddRow(), "Reason", $"{ workItem.Reason }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "RelatedLinkCount", $"{ workItem.RelatedLinkCount }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "RevisedDate", $"{ workItem.RevisedDate }", labelFormat: dateLabel, contentFormat: dateContent);
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "Revision", $"{ workItem.Revision }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "Revisions", $"{ workItem.Revisions.Count }");

                XlHlp.AddLabeledInfo(insertAt.AddRow(), "State", $"{ workItem.State }");

                XlHlp.AddLabeledInfo(insertAt.AddRow(), "Tags", $"{ workItem.Tags.Count() }");
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "TemporaryId", $"{ workItem.TemporaryId }");

                XlHlp.AddLabeledInfo(insertAt.AddRow(), "Uri", $"{ workItem.Uri }");

                XlHlp.AddLabeledInfo(insertAt.AddRow(), "Watermark", $"{ workItem.Watermark }");

                XlHlp.AddLabeledInfo(insertAt.AddRow(), "WorkItemLinkHistory", $"{ workItem.WorkItemLinkHistory.Count }");

                XlHlp.AddLabeledInfo(insertAt.AddRow(), "WorkItemLinks", $"{ workItem.WorkItemLinks.Count }");

                insertAt.MarkEnd(XlHlp.MarkType.Group);

                insertAt.Group(insertAt.OrientVertical);

                //if (options.ShowWorkItemFieldData)
                //{
                //    insertAt.IncrementRows();

                //    if (insertAt.OrientVertical)
                //    {
                //        XlHlp.AddLabeledInfo(insertAt.AddRow(), "WorkItem Fields", workItem.Id.ToString());
                //    }
                //    else
                //    {
                //        XlHlp.AddLabeledInfo(insertAt.AddRow(), "WorkItem Fields", workItem.Id.ToString(), orientation: XlOrientation.xlUpward);
                //        insertAt.IncrementColumns();
                //    }

                //    insertAt.IncrementColumns();

                //    insertAt.MarkStart(XlHlp.MarkType.GroupTable);

                //    Header_WorkItemStore.Add_TP_WorkItemFieldValues(insertAt);

                //    foreach (Field field in workItem.Fields)
                //    {
                //        insertAt.ClearOffsets();

                //        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ field.Id }");
                //        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ field.Name }");
                //        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ field.ReferenceName }");
                //        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $">{ field.OriginalValue }<");
                //        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $">{ field.Value }<");

                //        insertAt.IncrementRows();
                //    }

                //    insertAt.MarkEnd(XlHlp.MarkType.GroupTable);

                //    insertAt.Group(insertAt.OrientVertical);

                //    insertAt.DecrementColumns();
                //}

                ////insertAt.MarkEnd(XlHlp.MarkType.Group);

                ////insertAt.Group(insertAt.OrientVertical);

                insertAt.EndSectionAndSetNextLocation(insertAt.OrientVertical);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            XlHlp.DisplayInWatchWindow(insertAt, startTicks, "End");
            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);

            return insertAt;
        }

        internal static XlHlp.XlLocation Add_WorkItem_Links(
            XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options,
            WorkItemStore workItemStore,
            int workItemID)
        {
            Int64 startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);
            XlHlp.DisplayInWatchWindow(insertAt);

            try
            {

                // Save the location of the title so we can update later after have traversed all items.

                Range rngTitle = insertAt.GetCurrentRange();

                if (insertAt.OrientVertical)
                {
                    XlHlp.AddLabeledInfo(insertAt.AddRow(), "WorkItem Links", workItemID.ToString());
                }
                else
                {
                    XlHlp.AddLabeledInfo(insertAt.AddRow(), "WorkItem Links", workItemID.ToString(), orientation: XlOrientation.xlUpward);
                    insertAt.IncrementColumns();
                }

                insertAt.MarkStart(XlHlp.MarkType.GroupTable);

                Header_WorkItemStore.Add_TP_WorkItem_Links(insertAt);

                Body_WorkItemStore.Add_TP_WorkItem_Links(insertAt, options, workItemStore, workItemID);

                insertAt.MarkEnd(XlHlp.MarkType.GroupTable, string.Format("tblWIL_{0}", workItemID.ToString()));

                insertAt.Group(insertAt.OrientVertical);

                insertAt.EndSectionAndSetNextLocation(insertAt.OrientVertical);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            XlHlp.DisplayInWatchWindow(insertAt, startTicks, "End");
            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);

            return insertAt;
        }

        internal static XlHlp.XlLocation Add_WorkItem_WorkItemLinks(
            XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options,
            WorkItemStore workItemStore,
            WorkItem workItem)
        {
            Int64 startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);
            XlHlp.DisplayInWatchWindow(insertAt);

            try
            {

                // Save the location of the title so we can update later after have traversed all items.

                Range rngTitle = insertAt.GetCurrentRange();

                if (insertAt.OrientVertical)
                {
                    XlHlp.AddSectionInfo(insertAt.AddRow(), "WorkItem WorkItemLinks", workItem.Id.ToString());
                }
                else
                {
                    XlHlp.AddSectionInfo(insertAt.AddRow(), "WorkItem WorkItemLinks", workItem.Id.ToString(),
                        orientation: XlOrientation.xlUpward);
                    insertAt.IncrementColumns();
                }

                insertAt.MarkStart(XlHlp.MarkType.GroupTable);

                Header_WorkItemStore.Add_TP_WorkItem_WorkItemLinks(insertAt);

                Body_WorkItemStore.Add_TP_WorkItem_WorkItemLinks(insertAt, options, workItemStore, workItem);

                insertAt.MarkEnd(XlHlp.MarkType.GroupTable, string.Format("tblWIWIL_{0}", workItem.Id.ToString()));

                insertAt.Group(insertAt.OrientVertical);

                insertAt.EndSectionAndSetNextLocation(insertAt.OrientVertical);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            XlHlp.DisplayInWatchWindow(insertAt, startTicks, "End");
            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);

            return insertAt;
        }

        internal static XlHlp.XlLocation AddQueryNodes(
            XlHlp.XlLocation insertAt,
            QueryFolder queryFolder)
        {
            Int64 startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);
            XlHlp.DisplayInWatchWindow(insertAt);

            insertAt.ClearOffsets();

            XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 20, "Name");
            XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 20, "Node Type");

            insertAt.IncrementRows();

            foreach (var item in queryFolder)
            {
                insertAt.ClearOffsets();

                if (item is QueryDefinition)
                {
                    insertAt.ColumnOffset = 0;

                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), item.Name);
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), "QueryDefinition");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), ((QueryDefinition)item).QueryType.ToString());
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), ((QueryDefinition)item).QueryText);
                }

                if (item is QueryFolder)
                {
                    insertAt.ColumnOffset = 0;

                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), item.Name);
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), "QueryFolder");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), item.Id.ToString());
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), item.IsPersonal.ToString());
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), item.Path);

                    insertAt = AddQueryNodes(insertAt, (QueryFolder)item);
                }
            }

            XlHlp.DisplayInWatchWindow("End", startTicks);
            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);

            return insertAt;
        }


    }
}
