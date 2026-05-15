using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Xml;
using System.Xml.Linq;

using Microsoft.Office.Interop.Excel;
using Microsoft.TeamFoundation.Server;
using Microsoft.TeamFoundation.WorkItemTracking.Client;

using SupportTools_Excel.AzureDevOpsExplorer.Domain;

using VNC;
using VNC.AddinHelper;

using XlHlp = VNC.AddinHelper.Excel;

namespace SupportTools_Excel.AzureDevOpsExplorer.Application
{
    class Body_WorkItemStore
    {
        #region WorkItem Store (WIS)

        internal static XlHlp.XlLocation Add_TP_Areas(XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options,
            ICommonStructureService commonStructureService,
            Project project)
        {
            Int64 startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            if (project.AreaRootNodes.Count == 0)
            {
                XlHlp.AddContentToCell(insertAt.AddRowX(), "None");
            }
            else
            {
                insertAt = AddCategoryNodes(insertAt, options, commonStructureService, project.AreaRootNodes, project.Name);
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);

            return insertAt;
        }

        internal static void Add_TP_FieldMapping(
            XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options,
            Microsoft.TeamFoundation.WorkItemTracking.Client.Project project)
        {
            Int64 startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            XlHlp.DisplayInWatchWindow("Begin");

            Dictionary<WorkItemType, List<ControlFieldMap>> allMappings = new Dictionary<WorkItemType, List<ControlFieldMap>>();

            foreach (WorkItemType wit in project.WorkItemTypes.Cast<WorkItemType>().OrderBy(nnn => nnn.Name))
            {
                long loopTicks = XlHlp.DisplayInWatchWindow("WorkItemType Loop");

                try
                {
                    var mappings = GetFieldMappings(allMappings, wit);

                    foreach (var controlFieldMap in mappings)
                    {
                        insertAt.ClearOffsets();

                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{project.Name}");
                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{wit.Name}");

                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{controlFieldMap.FieldMap.Name}");
                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{controlFieldMap.FieldMap.RefName}");
                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{controlFieldMap.FieldMap.Type}");
                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{controlFieldMap.FieldMap.Required}");

                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{controlFieldMap.MapType}");

                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{controlFieldMap.ControlMap.Label}");
                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{controlFieldMap.ControlMap.Name}");
                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{controlFieldMap.ControlMap.FieldName}");
                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{controlFieldMap.ControlMap.Type}");

                        insertAt.IncrementRows();
                    }

                    AZDOHelper.ProcessItemDelay(options);

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }

                XlHlp.DisplayInWatchWindow("Loop", startTicks);
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);
        }

        internal static XlHlp.XlLocation Add_TP_Iterations(XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options,
            ICommonStructureService commonStructureService,
            Project project)
        {
            Int64 startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            if (project.IterationRootNodes.Count == 0)
            {
                XlHlp.AddContentToCell(insertAt.AddRowX(), "None");
            }
            else
            {
                insertAt = AddCategoryNodes(insertAt, options, commonStructureService, project.IterationRootNodes, project.Name);
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);

            return insertAt;
        }

         internal static void Add_TP_WorkItem_Info(
            XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options,
            WorkItemStore workItemStore,
            int workItemID)
        {
            Int64 startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            insertAt.ClearOffsets();
            int count = 0;

            try
            {
                WorkItem wi = VNC.TFS.Helper.RetrieveWorkItem(workItemID, workItemStore);
                insertAt.ClearOffsets();

                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ wi.Project.Name }");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ wi.Id }");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ wi.Type.Name }");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ wi.Title }");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ wi.CreatedBy }");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ wi.CreatedDate }");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ wi.ChangedBy }");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ wi.ChangedDate }");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ wi.Reason }");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ wi.State }");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ wi.AreaPath }");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ wi.IterationPath }");

                insertAt.IncrementRows();
                //}
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);
        }

        internal static void Add_TP_WorkItem_Links(
            XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options,
            WorkItemStore workItemStore,
            int workItemID)
        {
            Int64 startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            insertAt.ClearOffsets();
            int count = 0;

            try
            {
                string query = String.Format(
                    "Select [Id], [Created Date], [Changed Date], [Revised Date]"
                    + " From WorkItems"
                    + " Where [System.Id] = '{0}'",
                    workItemID);

                string query2 = String.Format(
                    "Select [Id], [System.Title]"
                    + " From WorkItemLinks"
                    + " Where Source.[System.Id] = '{0}'",
                    workItemID);


                Query wiQuery = new Query(workItemStore, query2);
                WorkItemLinkInfo[] wiTrees = wiQuery.RunLinkQuery();

                //PrintTrees(wiTrees, workItemID);

                WorkItemCollection queryResults = workItemStore.Query(query);

                if (queryResults.Count > 0)
                {
                    WorkItem wi = queryResults[0];

                    // TODO(crhodes)
                    // Figure out how wi.Links and wi.WorkItemLinks Differ
                    // Ok.  Look at Class Model.  Link is base type.
                    //  There are four derived types: ExternalLink, HyperLink, RelatedLink, WorkItemLink

                    foreach (Link link in wi.Links)
                    {
                        insertAt.ClearOffsets();

                        if (link.ArtifactLinkType != null)
                        {
                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ link.ArtifactLinkType.Name }");
                        }
                        else
                        {
                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), "<null>");
                        }

                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ link.BaseType }");

                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ link.Comment }");
                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ link.IsLocked }");
                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ link.IsNew }");
                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ ((RelatedLink)link).LinkTypeEnd.Id }");
                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ ((RelatedLink)link).LinkTypeEnd.Id }");

                        insertAt.IncrementRows();
                    }
                }
                else
                {
                    XlHlp.AddLabeledInfoX(insertAt.AddRowX(), "ID Not Found", workItemID.ToString()); ;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);
        }

        internal static void Add_TP_WorkItem_WorkItemLinks(
            XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options,
            WorkItemStore workItemStore,
            WorkItem workItem,
            int recursionLevel = 0)
        {
            Int64 startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            long startTicks2 = 0;

            if (options.ShowIndividualItems)
            {
                startTicks2 = XlHlp.DisplayInWatchWindow($"workItem:{workItem.Id} {workItem.Type.Name}");
            }

            insertAt.ClearOffsets();
            int count = 0;

            try
            {
                string queryWIL = String.Format(
                    "Select [Id], [System.Title]"
                    + " From WorkItemLinks"
                    + " Where Source.[System.Id] = '{0}'",
                    workItem.Id);

                Query wilQuery = new Query(workItemStore, queryWIL);
                WorkItemLinkInfo[] wiLinks = wilQuery.RunLinkQuery();

                // Get list of IDs for linked work teams (our targets)

                int[] linkedIDs = wiLinks.Select(i => i.TargetId).ToArray();
                int[] linkedIDsUnique = wiLinks.Select(i => i.TargetId).Distinct().ToArray();

                string queryWILdetails = String.Format(
                    "Select [Id], [System.Title], [System.WorkItemType]"
                    + " From WorkItems");

                // Not happy if duplicates

                //Query wilDetailsQuery = new Query(Server.WorkItemStore, queryWILdetails, linkedIDs);
                Query wilDetailsQuery = new Query(workItemStore, queryWILdetails, linkedIDsUnique);

                WorkItemCollection queryResultsWIL = wilDetailsQuery.RunQuery();

                List<WorkItem> bugWI = new List<WorkItem>();
                List<WorkItem> changeRequestWI = new List<WorkItem>();
                List<WorkItem> featureWI = new List<WorkItem>();
                List<WorkItem> milestoneWI = new List<WorkItem>();
                List<WorkItem> productionIssueWI = new List<WorkItem>();
                List<WorkItem> projectRiskWI = new List<WorkItem>();
                List<WorkItem> releaseWI = new List<WorkItem>();
                List<WorkItem> requirementWI = new List<WorkItem>();
                List<WorkItem> sharedStepsWI = new List<WorkItem>();
                List<WorkItem> specificationWI = new List<WorkItem>();
                List<WorkItem> taskWI = new List<WorkItem>();
                List<WorkItem> testCaseWI = new List<WorkItem>();
                List<WorkItem> testPlanWI = new List<WorkItem>();
                List<WorkItem> testSuiteWI = new List<WorkItem>();
                List<WorkItem> userNeedsWI = new List<WorkItem>();
                List<WorkItem> userStoryWI = new List<WorkItem>();
                List<WorkItem> voiceOfCustomerWI = new List<WorkItem>();

                // This catches what we do not cover specifically yet

                List<WorkItem> otherWI = new List<WorkItem>();

                CellFormatSpecification redContent = options.FormatSpecs.RedContent;
                CellFormatSpecification dateLabel = options.FormatSpecs.DateLabel;
                CellFormatSpecification dateContent = options.FormatSpecs.DateContent;
                CellFormatSpecification witContent = options.FormatSpecs.WITContent;

                foreach (WorkItemLink workItemLink in workItem.WorkItemLinks)
                {
                    insertAt.ClearOffsets();

                    // Doing this inside foreach is SUPER SLOW

                    //WorkItem target = Server.WorkItemStore.GetWorkItem(workItemLink.TargetId);
                    //XlHlp.AddContentToCell(insertAt.AddOffsetColumn(), $"{ target.Type.Name}");

                    // Use the dictionary instead to get the Type

                    //XlHlp.AddContentToCell(insertAt.AddOffsetColumn(), linkTypes[workItemLink.TargetId]);

                    // Use the collection of workitems to get details

                    int wiIndex = queryResultsWIL.IndexOf(workItemLink.TargetId);
                    WorkItem linkedWorkItem = queryResultsWIL[wiIndex];

                    switch (linkedWorkItem.Type.Name)
                    {
                        case "Bug":
                            bugWI.Add(linkedWorkItem);
                            break;

                        case "Change Request":
                            changeRequestWI.Add(linkedWorkItem);
                            break;

                        case "Feature":
                            milestoneWI.Add(linkedWorkItem);
                            break;

                        case "Milestone":
                            milestoneWI.Add(linkedWorkItem);
                            break;

                        case "Production Issue":
                            productionIssueWI.Add(linkedWorkItem);
                            break;

                        case "Project Risk":
                            projectRiskWI.Add(linkedWorkItem);
                            break;

                        case "Release":
                            releaseWI.Add(linkedWorkItem);
                            break;

                        case "Requirement":
                            requirementWI.Add(linkedWorkItem);
                            break;

                        case "Shared Steps":
                            sharedStepsWI.Add(linkedWorkItem);
                            break;

                        case "Specification":
                            specificationWI.Add(linkedWorkItem);
                            break;

                        case "Task":
                            taskWI.Add(linkedWorkItem);
                            break;

                        case "Test Case":
                            testCaseWI.Add(linkedWorkItem);
                            break;

                        case "Test Plan":
                            testPlanWI.Add(linkedWorkItem);
                            break;

                        case "Test Suite":
                            testSuiteWI.Add(linkedWorkItem);
                            break;

                        case "User Needs":
                            userNeedsWI.Add(linkedWorkItem);
                            break;

                        case "User Story":
                            userStoryWI.Add(linkedWorkItem);
                            break;

                        case "Voice Of Customer":
                            voiceOfCustomerWI.Add(linkedWorkItem);
                            break;

                        default:
                            otherWI.Add(linkedWorkItem);
                            break;
                    }

                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ workItem.Type.Name}", cellFormat: witContent);
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ workItem.Id}", cellFormat: redContent);
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ workItem.State}");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ workItem.Title}");

                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ linkedWorkItem.Type.Name}", cellFormat: witContent);
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ linkedWorkItem.Id}", cellFormat: redContent);
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ linkedWorkItem.State}");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ linkedWorkItem.Title}");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ linkedWorkItem.CreatedDate}");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ linkedWorkItem.CreatedBy}");

                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ workItemLink.SourceId}", cellFormat: redContent);
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ workItemLink.TargetId}", cellFormat: redContent);

                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ workItemLink.AddedBy}");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ workItemLink.AddedDate}", cellFormat: dateContent);

                    if (workItemLink.ArtifactLinkType != null)
                    {
                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ workItemLink.ArtifactLinkType.Name}");
                    }
                    else
                    {
                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), "<null>");
                    }

                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ workItemLink.BaseType }");

                    if (workItemLink.ChangedDate != null)
                    {
                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ workItemLink.ChangedDate}", cellFormat: dateContent);
                    }
                    else
                    {
                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), "<null>", cellFormat: dateContent);
                    }

                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ workItemLink.Comment}");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ workItemLink.IsLocked}");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ workItemLink.IsNew}");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ workItemLink.LinkTypeEnd.Id}");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ GetLastPartOfDelimitedName(workItemLink.LinkTypeEnd.ImmutableName, '.')}");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ workItemLink.LinkTypeEnd.IsForwardLink}");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ GetLastPartOfDelimitedName(workItemLink.LinkTypeEnd.LinkType.ToString(), '.')}");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ workItemLink.LinkTypeEnd.Name}");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ workItemLink.LinkTypeEnd.OppositeEnd.Id}");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ GetLastPartOfDelimitedName(workItemLink.LinkTypeEnd.OppositeEnd.ImmutableName, '.')}");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ workItemLink.LinkTypeEnd.OppositeEnd.IsForwardLink}");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ GetLastPartOfDelimitedName(workItemLink.LinkTypeEnd.OppositeEnd.LinkType.ToString(), '.')}");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ workItemLink.LinkTypeEnd.OppositeEnd.Name}");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ workItemLink.RemovedBy}");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ workItemLink.RemovedDate}", cellFormat: dateContent);

                    insertAt.IncrementRows();

                    count++;    // Helpful for debugging to see how far we have gotten
                }

                // Drill down (one level) on the WorkItems and get their links
                // This gets time consuming so only go one level down.

                if (recursionLevel < options.RecursionLevel)
                {
                    recursionLevel++;

                    // Why do we increment above but don't seem to call?
                    ProcessOneLevelDeeper(insertAt, options, workItemStore, recursionLevel, otherWI, "otherWI");

                    ProcessOneLevelDeeper(insertAt, options, workItemStore, recursionLevel, bugWI, "bugWI");
                    ProcessOneLevelDeeper(insertAt, options, workItemStore, recursionLevel, changeRequestWI, "changeRequestWI");
                    ProcessOneLevelDeeper(insertAt, options, workItemStore, recursionLevel, featureWI, "featureWI");
                    ProcessOneLevelDeeper(insertAt, options, workItemStore, recursionLevel, milestoneWI, "milestoneWI");
                    ProcessOneLevelDeeper(insertAt, options, workItemStore, recursionLevel, productionIssueWI, "productionIssueWI");
                    ProcessOneLevelDeeper(insertAt, options, workItemStore, recursionLevel, projectRiskWI, "projectRiskWI");
                    ProcessOneLevelDeeper(insertAt, options, workItemStore, recursionLevel, releaseWI, "releaseWI");
                    ProcessOneLevelDeeper(insertAt, options, workItemStore, recursionLevel, requirementWI, "requirementWI");
                    ProcessOneLevelDeeper(insertAt, options, workItemStore, recursionLevel, sharedStepsWI, "sharedStepsWI");
                    ProcessOneLevelDeeper(insertAt, options, workItemStore, recursionLevel, specificationWI, "specificationWI");
                    ProcessOneLevelDeeper(insertAt, options, workItemStore, recursionLevel, taskWI, "taskWI");
                    ProcessOneLevelDeeper(insertAt, options, workItemStore, recursionLevel, testCaseWI, "testCaseWI");
                    ProcessOneLevelDeeper(insertAt, options, workItemStore, recursionLevel, testPlanWI, "testPlanWI");
                    ProcessOneLevelDeeper(insertAt, options, workItemStore, recursionLevel, testSuiteWI, "testSuiteWI");
                    ProcessOneLevelDeeper(insertAt, options, workItemStore, recursionLevel, userNeedsWI, "userNeedsWI");
                    ProcessOneLevelDeeper(insertAt, options, workItemStore, recursionLevel, userStoryWI, "userStoryWI");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            if (options.ShowIndividualItems)
            {
                XlHlp.DisplayInWatchWindow(insertAt, startTicks2, "End");
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);
        }

        internal static void Add_TP_WorkItemDetails(
            XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options,
            WorkItemCollection queryResults)
        {
            Int64 startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            int itemCount = 0;
            int totalItems = queryResults.Count;

            CellFormatSpecification redContent = options.FormatSpecs.RedContent;

            try
            {
                foreach (WorkItem workItem in queryResults)
                {
                    insertAt.ClearOffsets();

                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ workItem.Project.Name }");

                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ workItem.Id }");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ workItem.Type.Name }");

                    try
                    {
                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ workItem.Title }");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }

                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ workItem.CreatedBy }");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ workItem.CreatedDate }");

                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ workItem.ChangedBy }");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ workItem.ChangedDate }");

                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ workItem.AuthorizedDate }");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ workItem.RevisedDate }");

                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ workItem.Rev }");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ workItem.Revision }");

                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ workItem.State }");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ workItem.Reason }");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ workItem.Tags }");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ workItem.AreaPath }");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ workItem.IterationPath }");

                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ workItem.RelatedLinkCount }");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ workItem.ExternalLinkCount }");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ workItem.HyperLinkCount }");

                    // TODO(crhodes)
                    // Things to consider displaying
                    // AttachedFileCount
                    // Attachments.Count
                    // Description
                    // IsAccessDenied
                    // IsAccessDenied
                    // IsDirty
                    // IsNew
                    // IsOpen
                    // IsPartialOpen
                    // IsReadOnly
                    // IsReadOnlyOpen
                    // AreaId
                    // IterationId
                    // Uri

                    // NOTE(crhodes)
                    // The Query can specify additional fields to display
                    // They are added to the query to avoid a round trip if accessed 
                    // after the result set is returned.

                    // Display them.  Some may not exist, catch and display N/A

                    if ((options.WorkItemQuerySpec.Fields?.Count ?? 0) > 0)
                    {
                        foreach (string field in options.WorkItemQuerySpec.Fields)
                        {
                            try
                            {
                                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ workItem.Fields[field].Value }");
                            }
                            catch (Exception ex)
                            {
                                // NOTE(crhodes)
                                // Exception is thrown trying to access field which occurs
                                // before column is incremented.
                                //insertAt.DecrementColumns();
                                insertAt.ColumnOffset--;
                                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), "...");
                            }
                        }
                    }

                    insertAt.IncrementRows();

                    if (options.RetrieveRevisions is true)
                    {
                        int currentOffset = insertAt.ColumnOffset;

                        Boolean isClosed = false;
                        int closedCount = 0;

                        // TODO(crhodes)
                        // Figure out what we can get from Revisions.
                        foreach (Revision revision in workItem.Revisions)
                        {

                            try
                            {
                                // HACK(crhodes)
                                // Let's see what is in revision.Fields

                                // Sigh, it is all fields.

                                if (options.RetrieveFieldChanges is true)
                                {
                                    string oldValue;
                                    string currentValue;

                                    foreach (Field revisionField in revision.Fields)
                                    {

                                        //insertAt.ClearOffsets();
                                        insertAt.SetColumnOffset(currentOffset);

                                        oldValue = (revision.Fields[revisionField.Name].Value ?? "").ToString();
                                        currentValue = (revision.Fields[revisionField.Name].OriginalValue ?? "").ToString();


                                        //var originalValue = revisionField.OriginalValue;
                                        //var currentValue = revisionField.Value;
                                        //var hashOriginalValue = originalValue.GetHashCode();
                                        //var hashCurrentValue = currentValue.GetHashCode();
                                        //var stringOriginalValue = originalValue.ToString();
                                        //var stringCurrentValue = currentValue.ToString();

                                        //var same = originalValue.Equals(currentValue);
                                        //var hashSame = hashOriginalValue.Equals(hashCurrentValue);
                                        //var stringSame = stringOriginalValue.Equals(stringCurrentValue);

                                        //if (! same)
                                        //{

                                        if (!oldValue.Equals(currentValue))
                                        {
                                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ revisionField.Name }");
                                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ revisionField.ReferenceName }");

                                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ revisionField.OriginalValue }");
                                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ revisionField.Value }");
                                            insertAt.IncrementRows();
                                        }
                                        //XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ same }");
                                        //XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ hashSame }");
                                        //XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ stringSame }");


                                        //}
                                    }
                                }

                                insertAt.ClearOffsets();

                                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ revision.WorkItem.Project.Name }");

                                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ revision.WorkItem.Id }");
                                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ revision.WorkItem.Type.Name }");
                                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ revision.Fields["Title"].Value }");

                                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ revision.Fields["Created By"].Value }");
                                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ revision.Fields["Created Date"].Value }");

                                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ revision.Fields["Changed By"].Value }");
                                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ revision.Fields["Changed Date"].Value }");

                                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ revision.Fields["Authorized Date"].Value }");
                                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ revision.Fields["Revised Date"].Value }");
                                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ revision.Fields["Rev"].Value }");
                                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), "");   // Placeholder for Revision which is not a Field on Revision

                                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ revision.Fields["State"].Value }");
                                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ revision.Fields["Reason"].Value }");
                                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ revision.Fields["Tags"].Value }");
                                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ revision.Fields["Area Path"].Value }");
                                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ revision.Fields["Iteration Path"].Value }");

                                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ revision.Fields["Related Link Count"].Value }");
                                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ revision.Fields["External Link Count"].Value }");
                                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ revision.Fields["HyperLink Count"].Value }");

                                if ((options.WorkItemQuerySpec.Fields?.Count ?? 0) > 0)
                                {
                                    foreach (string field in options.WorkItemQuerySpec.Fields)
                                    {
                                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ revision.Fields[field].Value }");
                                    }
                                }

                                if (isClosed)
                                {
                                    if (revision.Fields["State"].Value.ToString() != "Closed")
                                    {
                                        isClosed = false;
                                    }

                                    continue;
                                }

                                if (revision.Fields["State"].Value.ToString() == "Closed")
                                {
                                    isClosed = true;
                                    closedCount++;
                                }

                                if (closedCount > 1)
                                {
                                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{ closedCount }", redContent);
                                }
                                
                            }
                            catch (Exception ex)
                            {
                                
                            }

                            insertAt.IncrementRows();
                        }
                    }

                    itemCount++;

                    if (itemCount % options.LoopUpdateInterval == 0)
                    {
                        XlHlp.DisplayInWatchWindow($"Added {itemCount} out of {totalItems}");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);
        }

        internal static void Add_TP_WorkItemFields(
            XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options,
            Project project)
        {
            Int64 startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            foreach (WorkItemType wit in project.WorkItemTypes.Cast<WorkItemType>().OrderBy(nnn => nnn.Name))
            {
                insertAt.ClearOffsets();

                foreach (FieldDefinition fieldDef in wit.FieldDefinitions)
                {
                    var fieldName = fieldDef.Name;
                    var fieldType = fieldDef.SystemType;

                    switch (fieldDef.SystemType.FullName)
                    {
                        case "System.DateTime":
                            break;

                        case "System.Double":
                            break;

                        case "System.String":
                            break;

                        default:
                            break;
                    }


                    //sb.AppendFormat("{0}[{1}],", fieldName, fieldType);

                    insertAt.ClearOffsets();

                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), string.Format("{0}", project.Name));
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), string.Format("{0}", wit.Name));
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), string.Format("{0}", fieldDef.Name));
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), string.Format("{0}", fieldDef.FieldType));
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), string.Format("{0}", fieldDef.SystemType));

                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), string.Format("{0}", fieldDef.Id));
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), string.Format("{0}", fieldDef.IsComputed));
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), string.Format("{0}", fieldDef.IsCoreField));
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), string.Format("{0}", fieldDef.IsEditable));
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), string.Format("{0}", fieldDef.IsIdentity));
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), string.Format("{0}", fieldDef.IsIndexed));
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), string.Format("{0}", fieldDef.IsQueryable));
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), string.Format("{0}", fieldDef.ReferenceName));

                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), string.Format("{0}", fieldDef.ReportingAttributes.Name));
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), string.Format("{0}", fieldDef.ReportingAttributes.ReferenceName));

                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), string.Format("{0}", fieldDef.Usage));

                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), string.Format("{0}", fieldName + "[" + fieldType + "]"));

                    if (fieldDef.AllowedValues.Count > 0)
                    {
                        StringBuilder sb = new StringBuilder();

                        foreach (var value in fieldDef.AllowedValues)
                        {
                            if (sb.Length > 0)
                            {
                                sb.Append($";{value}");
                            }
                            else
                            {
                                sb.Append(value);
                            }
                        }

                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{sb}");
                    }

                    if (fieldDef.ProhibitedValues.Count > 0)
                    {
                        StringBuilder sb = new StringBuilder();

                        foreach (var value in fieldDef.ProhibitedValues)
                        {
                            if (sb.Length > 0)
                            {
                                sb.Append($";{value}");
                            }
                            else
                            {
                                sb.Append(value);
                            }
                        }

                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{sb}");
                        //XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), string.Format("{0}", sb.ToString()));
                    }

                    insertAt.IncrementRows();
                }
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);
        }

        internal static void Add_TP_WorkItemTypes(
            XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options,
            WorkItemStore workItemStore,
            Project project,
            out DateTime maxLastCreatedDate,
            out DateTime maxLastChangedDate,
            out DateTime maxLastRevisedDate)
        {
            Int64 startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            XlHlp.DisplayInWatchWindow("Begin");

            maxLastCreatedDate = DateTime.MinValue;
            maxLastChangedDate = DateTime.MinValue;
            maxLastRevisedDate = DateTime.MinValue;

            //DateTime startingDate = (DateTime.Now - TimeSpan.FromDays(options.GoBackDays));
            //string startDate = "1/1/1900";

            //if (options.GoBackDays > 0)
            //{
            //    startDate = startingDate.ToShortDateString();
            //}

            string startDate = options.StartDate.ToShortDateString();

            Dictionary<WorkItemType, List<Transition>> allTransitions = new Dictionary<WorkItemType, List<Transition>>();

            foreach (WorkItemType wit in project.WorkItemTypes.Cast<WorkItemType>().OrderBy(nnn => nnn.Name))
            {
                long loopTicks = XlHlp.DisplayInWatchWindow("WorkItemType Loop");

                string exportXMLFilePath = "";

                if (options.ExportXMLTemplate)
                {
                    exportXMLFilePath = $@"{options.XMLTemplateFilePath}\{project.Name}";

                    Directory.CreateDirectory(exportXMLFilePath);

                    XmlDocument exportXml = wit.Export(includeGlobalListsFlag: false);
                    exportXml.Save($@"{exportXMLFilePath}\{wit.Name}.txt");

                    if (options.IncludeGlobalLists)
                    {
                        XmlDocument exportXmlGlobalLists = wit.Export(includeGlobalListsFlag: true);
                        exportXmlGlobalLists.Save($@"{exportXMLFilePath}\{wit.Name}.gl.txt");
                    }
                }

                try
                {
                    var transitions = GetTransitions(allTransitions, wit);

                    string transitionsDisplay = PrintTransitions(transitions);

                    insertAt.ClearOffsets();
                    int count = 0;

                    string lastCreateDate = "???";
                    string lastChangedDate = "???";
                    string lastRevisedDate = "???";

                    if (options.GetLastActivityDates)
                    {
                        try
                        {
                            string query = String.Format(
                                "Select [Id], [Created Date], [Changed Date], [Revised Date]"
                                + " From WorkItems"
                                + " Where [Work Item Type] = '{0}'"
                                + " and [System.TeamProject] = '{1}'"
                                + " and ([Created Date] >= '{2}' or [Changed Date] >= '{2}')",
                                wit.Name, project.Name, startDate);

                            WorkItemCollection queryResults = workItemStore.Query(query);

                            if ((count = queryResults.Count) > 0)
                            {
                                WorkItem lastCreatedItem = queryResults.Cast<WorkItem>().OrderByDescending(iii => iii.CreatedDate).First();
                                lastCreateDate = lastCreatedItem.CreatedDate.ToString();

                                if (lastCreatedItem.CreatedDate > maxLastCreatedDate)
                                {
                                    maxLastCreatedDate = lastCreatedItem.CreatedDate;
                                }

                                WorkItem lastChangedItem = queryResults.Cast<WorkItem>().OrderByDescending(iii => iii.ChangedDate).First();
                                lastChangedDate = lastChangedItem.ChangedDate.ToString();

                                if (lastChangedItem.ChangedDate > maxLastChangedDate)
                                {
                                    maxLastChangedDate = lastChangedItem.ChangedDate;
                                }

                                WorkItem lastRevisedItem = queryResults.Cast<WorkItem>().OrderByDescending(iii => iii.RevisedDate).First();
                                lastRevisedDate = lastRevisedItem.RevisedDate.ToString();

                                if (lastRevisedItem.RevisedDate > maxLastRevisedDate)
                                {
                                    maxLastRevisedDate = lastRevisedItem.RevisedDate;
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.ToString());
                        }

                        if (options.SkipIfNoActivity && lastCreateDate == "???")
                        {
                            continue;
                        }
                    }

                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{project.Name}");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{wit.Name}");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{count}");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{wit.FieldDefinitions.Count}");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{lastCreateDate}");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{lastChangedDate}");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{lastRevisedDate}");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{transitionsDisplay}");

                    insertAt.IncrementRows();

                    AZDOHelper.ProcessItemDelay(options);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }

                XlHlp.DisplayInWatchWindow("Loop", startTicks);
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);
        }

        internal static void Add_TP_WorkItemActivity(
            XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options,
            WorkItemStore workItemStore,
            Project project,
            out DateTime maxLastCreatedDate,
            out DateTime maxLastChangedDate,
            out DateTime maxLastRevisedDate)
        {
            Int64 startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            XlHlp.DisplayInWatchWindow("Begin");

            maxLastCreatedDate = DateTime.MinValue;
            maxLastChangedDate = DateTime.MinValue;
            maxLastRevisedDate = DateTime.MinValue;

            //DateTime startingDate = (DateTime.Now - TimeSpan.FromDays(options.GoBackDays));
            //string startDate = "1/1/1900";

            //if (options.GoBackDays > 0)
            //{
            //    startDate = startingDate.ToShortDateString();
            //}

            string startDate = options.StartDate.ToShortDateString();

            //Dictionary<WorkItemType, List<Transition>> allTransitions = new Dictionary<WorkItemType, List<Transition>>();

            foreach (WorkItemType wit in project.WorkItemTypes.Cast<WorkItemType>().OrderBy(nnn => nnn.Name))
            {
                long loopTicks = XlHlp.DisplayInWatchWindow("WorkItemType Loop");

                string exportXMLFilePath = "";

                //if (options.ExportXMLTemplate)
                //{
                //    exportXMLFilePath = $@"{options.XMLTemplateFilePath}\{project.Name}";

                //    Directory.CreateDirectory(exportXMLFilePath);

                //    XmlDocument exportXml = wit.Export(includeGlobalListsFlag: false);
                //    exportXml.Save($@"{exportXMLFilePath}\{wit.Name}.txt");

                //    if (options.IncludeGlobalLists)
                //    {
                //        XmlDocument exportXmlGlobalLists = wit.Export(includeGlobalListsFlag: true);
                //        exportXmlGlobalLists.Save($@"{exportXMLFilePath}\{wit.Name}.gl.txt");
                //    }
                //}

                try
                {
                    //var transitions = GetTransitions(allTransitions, wit);

                    //string transitionsDisplay = PrintTransitions(transitions);

                    insertAt.ClearOffsets();
                    int count = 0;

                    string lastCreateDate = "???";
                    string lastChangedDate = "???";
                    string lastRevisedDate = "???";

                    //if (options.GetLastActivityDates)
                    //{
                        try
                        {
                            string query = String.Format(
                                "Select [Id], [Created Date], [Changed Date], [Revised Date]"
                                + " From WorkItems"
                                + " Where [Work Item Type] = '{0}'"
                                + " and [System.TeamProject] = '{1}'"
                                + " and ([Created Date] >= '{2}' or [Changed Date] >= '{2}')",
                                wit.Name, project.Name, startDate);

                            WorkItemCollection queryResults = workItemStore.Query(query);

                            if ((count = queryResults.Count) > 0)
                            {
                                WorkItem lastCreatedItem = queryResults.Cast<WorkItem>().OrderByDescending(iii => iii.CreatedDate).First();
                                lastCreateDate = lastCreatedItem.CreatedDate.ToString();

                                if (lastCreatedItem.CreatedDate > maxLastCreatedDate)
                                {
                                    maxLastCreatedDate = lastCreatedItem.CreatedDate;
                                }

                                WorkItem lastChangedItem = queryResults.Cast<WorkItem>().OrderByDescending(iii => iii.ChangedDate).First();
                                lastChangedDate = lastChangedItem.ChangedDate.ToString();

                                if (lastChangedItem.ChangedDate > maxLastChangedDate)
                                {
                                    maxLastChangedDate = lastChangedItem.ChangedDate;
                                }

                                WorkItem lastRevisedItem = queryResults.Cast<WorkItem>().OrderByDescending(iii => iii.RevisedDate).First();
                                lastRevisedDate = lastRevisedItem.RevisedDate.ToString();

                                if (lastRevisedItem.RevisedDate > maxLastRevisedDate)
                                {
                                    maxLastRevisedDate = lastRevisedItem.RevisedDate;
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.ToString());
                        }

                        if (options.SkipIfNoActivity && lastCreateDate == "???")
                        {
                            continue;
                        }
                    //}

                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{project.Name}");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{wit.Name}");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{count}");
                    //XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{wit.FieldDefinitions.Count}");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{lastCreateDate}");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{lastChangedDate}");
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{lastRevisedDate}");
                    //XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), $"{transitionsDisplay}");

                    insertAt.IncrementRows();

                    AZDOHelper.ProcessItemDelay(options);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }

                XlHlp.DisplayInWatchWindow("Loop", startTicks);
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);
        }

        internal static void Get_TP_WorkItemTypesXML(
            Options_AZDO_TFS options,
            Project project)
        {
            Int64 startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            foreach (WorkItemType wit in project.WorkItemTypes.Cast<WorkItemType>().OrderBy(nnn => nnn.Name))
            {
                long loopTicks = XlHlp.DisplayInWatchWindow("WorkItemType Loop");

                string exportXMLFilePath = $@"{options.XMLTemplateFilePath}\{project.Name}";

                Directory.CreateDirectory(exportXMLFilePath);

                XmlDocument exportXml = wit.Export(includeGlobalListsFlag: false);
                exportXml.Save($@"{exportXMLFilePath}\{wit.Name}.txt");

                if (options.IncludeGlobalLists)
                {
                    XmlDocument exportXmlGlobalLists = wit.Export(includeGlobalListsFlag: true);
                    exportXmlGlobalLists.Save($@"{exportXMLFilePath}\{wit.Name}.gl.txt");
                }

                XlHlp.DisplayInWatchWindow("Loop", startTicks);
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);
        }

        //internal static void Add_WorkItemDetails(
        //    XlHlp.XlLocation insertAt,
        //    Options_AZDO_TFS options)
        //{
        //    // TODO(crhodes)
        //    // Loop across Team Projects and get last change or maybe go back days
        //}

        internal static XlHlp.XlLocation AddCategoryNodes(
            XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options,
            ICommonStructureService commonStructureService,
            NodeCollection childNodes,
            string projectName)
        {
            Int64 startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            insertAt.UpdateOffsets();

            foreach (Node item in childNodes)
            {
                var nodeInfo = commonStructureService.GetNode(item.Uri.ToString());

                Range startofRowRange = (Range)insertAt.workSheet.Cells[insertAt.RowCurrent, 1];

                XlHlp.AddContentToCell(startofRowRange, $"{ projectName }");

                insertAt.IncrementColumns();

                if (item.IsAreaNode)
                {
                    // HACK(crhodes)
                    // Somehow this needs to use the offsetLevel to get back to the first column or just hard code it.

                    XlHlp.AddContentToCell(insertAt.AddRowX(), item.Name);

                    if (options.ShowAllNodeLevels && item.HasChildNodes)
                    {
                        insertAt = AddCategoryNodes(insertAt, options, commonStructureService, item.ChildNodes, projectName);
                    }
                }

                if (item.IsIterationNode)
                {
                    string startdate = nodeInfo.StartDate.HasValue ? ((DateTime)nodeInfo.StartDate).ToShortDateString() : "<null>";
                    string finishdate = nodeInfo.FinishDate.HasValue ? ((DateTime)nodeInfo.FinishDate).ToShortDateString() : "<null>";

                    string days = "??";

                    if (nodeInfo.StartDate.HasValue)
                    {
                        days = ((DateTime)nodeInfo.FinishDate).Subtract((DateTime)nodeInfo.StartDate).TotalDays.ToString();
                    }

                    string iterationinfo = $"{item.Name, -40} (id: {item.Id}) - {days,3} days ({startdate} to {finishdate})";

                    XlHlp.AddContentToCell(insertAt.AddRowX(), iterationinfo);

                    if (options.ShowAllNodeLevels && item.HasChildNodes)
                    {
                        insertAt = AddCategoryNodes(insertAt, options, commonStructureService, item.ChildNodes, projectName);
                    }
                }

                insertAt.DecrementColumns();
            }

            // NOTE(crhodes)
            // This "fixes" the off by one error on the tables that get produced.
            // It is a HACK but does work.
            // Someday go figure out what is not being done above.  Maybe AddCategoryNodes should call Add OffsetColumn
            insertAt.UpdateOffsets();

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);

            return insertAt;
        }

        private static List<ControlFieldMap> GetFieldMappings(
            Dictionary<WorkItemType, List<ControlFieldMap>> allMappings,
            WorkItemType workItemType)
        {
            Int64 startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            List<ControlFieldMap> currentMappings;

            allMappings.TryGetValue(workItemType, out currentMappings);

            if (currentMappings != null)
            {
                return currentMappings;
            }

            var newMappings = new List<ControlFieldMap>();

            try
            {
                XmlDocument workItemTypeXml = workItemType.Export(false);
                XDocument xDocument = XDocument.Parse(workItemTypeXml.OuterXml);
                XElement xElement = XElement.Parse(workItemTypeXml.OuterXml).Element("WORKITEMTYPE");

                // N.B. FIELDS and FIELD exist in WORKITEMTYPE and WORKITEMTYPE\WORKFLOW\STATES
                // be careful to only get the ones from WORKITEMTYPE\FIELDS

                var fields = xElement.Element("FIELDS").Elements("FIELD");
                //var fields2 = xDocument.Descendants("WORKITEMTYPE").      .Element("FIELDS").Descendants("FIELD");

                // This seems to be a clean way of getting to what we want.
                var layoutControls = xDocument.Descendants("Layout").Descendants("Control");
                var webLayoutControls = xDocument.Descendants("WebLayout").Descendants("Control");

                List<ControlMap> layoutControlList = new List<ControlMap>();
                List<ControlMap> webLayoutControlList = new List<ControlMap>();

                // By default Dictionary is Case sENSITIVE
                //Dictionary<string, FieldMap> fieldDictionary = new Dictionary<string, FieldMap>();

                // Tell Dictionary to ignore case and avoid the ToLower() junk when interacting with Keys.

                Dictionary<string, FieldMap> fieldDictionary = new Dictionary<string, FieldMap>(StringComparer.OrdinalIgnoreCase);

                Dictionary<string, ControlMap> layoutControlDictionary = new Dictionary<string, ControlMap>(StringComparer.OrdinalIgnoreCase);
                Dictionary<string, ControlMap> webLayoutControlDictionary = new Dictionary<string, ControlMap>(StringComparer.OrdinalIgnoreCase);

                Hashtable controlHashtable = new Hashtable();

                // Get all the Fields

                var countFieldNodes = fields.Count();

                foreach (XElement field in fields)
                {
                    // Some fields are inconsistent between the Fields Definition and Layout Sections
                    // e.g. System.Id and System.ID.  Force to lower so we can find them later.
                    // But if we do that they show up in lower case :( system.id
                    // Better to tell dictionary to ignore case, infra :)

                    //string refName = field.Attributes["refname"].Value.ToLower();
                    string refName = field.Attribute("refname").Value;
                    string name = "";
                    string type = "";
                    bool required = false;

                    name = field.Attribute("name")?.Value ?? "";
                    type = field.Attribute("type")?.Value ?? "";

                    if (field.Descendants("REQUIRED").Any())
                    {
                        required = true;
                    }

                    // TODO(crhodes)
                    // Name and Type may not exist.  Figure out null check.

                    if (fieldDictionary.ContainsKey(refName))
                    {
                        MessageBox.Show($"refName: {refName} already exists in fieldDictionary.  ");
                    }
                    else
                    {
                        fieldDictionary.Add(refName, new FieldMap
                        {
                            Name = name,
                            RefName = refName,
                            Type = type,
                            Required = required
                        });
                    }
                }

                foreach (XElement control in layoutControls)
                {
                    ControlMap controLMap = new ControlMap
                    {
                        Name = control.Attribute("Name")?.Value ?? "",
                        FieldName = control.Attribute("FieldName")?.Value ?? "",
                        Type = control.Attribute("Type")?.Value ?? "",
                        Label = control.Attribute("Label")?.Value ?? ""
                    };

                    layoutControlList.Add(controLMap);

                    if (control.Attribute("FieldName") != null)
                    {
                        if (layoutControlDictionary.ContainsKey(control.Attribute("FieldName").Value))
                        {
                            MessageBox.Show($"WIT: {workItemType.Name}  Already found {control.Attribute("FieldName").Value} in layoutControls Label: {controLMap.Label} Name: {controLMap.Name} ");
                        }
                        else
                        {
                            layoutControlDictionary.Add(control.Attribute("FieldName").Value, controLMap);
                        }
                    }
                    else
                    {
                        var type = control.Attribute("Type").Value;

                        if ((type != "LinksControl") && (type != "AttachmentsControl") && (type != "AssociatedAutomationControl"))
                        {
                            MessageBox.Show($"No FieldName and unrecognized type: {type}");
                        }
                    }
                }

                foreach (XElement control in webLayoutControls)
                {

                    ControlMap controLMap = new ControlMap
                    {
                        Name = control.Attribute("Name")?.Value ?? "",
                        FieldName = control.Attribute("FieldName")?.Value ?? "",
                        Type = control.Attribute("Type")?.Value ?? "",
                        Label = control.Attribute("Label")?.Value ?? ""
                    };

                    webLayoutControlList.Add(controLMap);

                    if (control.Attribute("FieldName") != null)
                    {
                        if (webLayoutControlDictionary.ContainsKey(control.Attribute("FieldName").Value))
                        {
                            MessageBox.Show($"WIT: {workItemType.Name}  Already found {control.Attribute("FieldName").Value} in webLayoutControls Label: {controLMap.Label} Name: {controLMap.Name} ");
                        }
                        else
                        {
                            webLayoutControlDictionary.Add(control.Attribute("FieldName").Value, controLMap);
                        }
                    }
                    else
                    {
                        var type = control.Attribute("Type").Value;

                        if ((type != "LinksControl") && (type != "AttachmentsControl") && (type != "AssociatedAutomationControl"))
                        {
                            MessageBox.Show($"No FieldName and unrecognized type: {type}");
                        }
                    }
                }

                var countLayoutControlList = layoutControlList.Count;
                var countWebLayoutControlList = webLayoutControlList.Count;

                // TODO(crhodes)
                // Maybe we should go the other way and loop the fields and then see if any
                // Layout or WebLayout controls display the field.

                // Iterate all the Layout Controls and get the appropriate FieldMap

                foreach (var item in fieldDictionary)
                {
                    try
                    {
                        ControlFieldMap controlFieldMap = new ControlFieldMap();

                        controlFieldMap.FieldMap = fieldDictionary[item.Key];
                        string refName = controlFieldMap.FieldMap.RefName;

                        // Have to loop as field could be use in many places

                        foreach (var control in layoutControlDictionary.Values.Where(c => c.FieldName == refName))
                        {
                            controlFieldMap.MapType = "Layout";

                            controlFieldMap.ControlMap = layoutControlDictionary[refName];
                            newMappings.Add(controlFieldMap);
                        }

                        foreach (var control in webLayoutControlDictionary.Values.Where(c => c.FieldName == refName))
                        {
                            controlFieldMap.MapType = "WebLayout";

                            controlFieldMap.ControlMap = webLayoutControlDictionary[refName];
                            newMappings.Add(controlFieldMap);
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }
                }

                allMappings.Add(workItemType, newMappings);

                XlHlp.DisplayInWatchWindow($"WorkItem: {workItemType.Name} FieldNodes: {countFieldNodes}  LayoutControlList: {countLayoutControlList}  WebLayoutControlList: {countWebLayoutControlList}  newMappings: {newMappings.Count}");

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);

            return newMappings;
        }

        static string GetLastPartOfDelimitedName(string inputString, char delimiter)
        {
            var lio = inputString.LastIndexOf(delimiter);
            return inputString.Substring(lio + 1);
        }

        /// <summary>
        /// Get the transitions for this <see cref="WorkItemType"/>
        /// </summary>
        /// <param name="workItemType"></param>
        /// <returns></returns>
        private static List<Transition> GetTransitions(Dictionary<WorkItemType, List<Transition>> allTransitions,
            WorkItemType workItemType)
        {
            Int64 startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            List<Transition> currentTransitions;
            Boolean supportsExport = true;

            // See if this WorkItemType has already had it's transitions figured out.
            allTransitions.TryGetValue(workItemType, out currentTransitions);

            if (currentTransitions != null)
            {
                return currentTransitions;
            }

            // Create a dictionary to allow us to look up the "to" state using a "from" state.
            var newTransitions = new List<Transition>();

            // Get this worktype type as xml
            try
            {
                var foo = workItemType.Name;
                XmlDocument workItemTypeXml = null;
                try
                {
                    workItemTypeXml = workItemType.Export(false);
                }
                catch (Exception ex)
                {
                    // Inherited Process Templates cannot be exported.
                    newTransitions.Add(new Transition
                    {
                        From = "Not Supported by Process Template",
                        To = "",
                        Fields = "",
                        Reasons = "",
                        For = ""
                    });

                    supportsExport = false;
                }

                if (supportsExport)
                {
                    // get the transitions node.
                    XmlNodeList transitionsList = workItemTypeXml.GetElementsByTagName("TRANSITIONS");

                    // As there is only one transitions item we can just get the first
                    XmlNode transitions = transitionsList[0];

                    // Iterate all the transitions
                    foreach (XmlNode transition in transitions)
                    {
                        StringBuilder reasons = new StringBuilder();
                        StringBuilder fields = new StringBuilder();

                        XmlNode reasonsNode = transition.SelectSingleNode("REASONS");
                        XmlNode fieldsNode = transition.SelectSingleNode("FIELDS");

                        foreach (XmlNode reason in reasonsNode)
                        {
                            if (reasons.Length != 0)
                            {
                                reasons.Append($", {reason.Attributes["value"].Value}");
                            }
                            else
                            {
                                reasons.Append(reason.Attributes["value"].Value);
                            }

                            if (reason.Name == "DEFAULTREASON")
                            {
                                reasons.Append("*");
                            }
                        }

                        // Not all REASONS have required FIELDS

                        if (fieldsNode != null)
                        {
                            foreach (XmlNode field in fieldsNode)
                            {
                                try
                                {
                                    string trimedField = field.Attributes["refname"].Value.Replace("Microsoft.", "M.");

                                    if (fields.Length != 0)
                                    {
                                        fields.Append($", {trimedField}");
                                    }
                                    else
                                    {
                                        fields.Append(trimedField);
                                    }
                                }
                                catch (Exception ex)
                                {

                                }

                                // TODO(crhodes)
                                // Maybe show <EMPTY />
                                //if (field.Name == "DEFAULTREASON")
                                //{
                                //    reasons.Append("*");
                                //}
                            }
                        }

                        // save off the transition 
                        newTransitions.Add(new Transition
                        {
                            From = transition.Attributes["from"].Value,
                            To = transition.Attributes["to"].Value,
                            For = transition.Attributes["for"] != null ? $"for {transition.Attributes["for"].Value}" : "",
                            Reasons = reasons.ToString(),
                            Fields = fields.ToString()
                        });
                    }

                    // Add transition so we don't do it again if it is needed.
                    allTransitions.Add(workItemType, newTransitions);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);

            return newTransitions;
        }

        private static string PrintTransitions(List<Transition> transitions)
        {
            StringBuilder sb = new StringBuilder();
            string pad = new string(' ', 40);

            foreach (var transition in transitions.OrderBy(n => n.From))
            {
                if (sb.Length == 0)
                {
                    sb.Append($"{transition.From,17} -> {transition.To,-17} ({transition.Reasons})");
                    if (transition.For.Length > 0) sb.Append($" > {pad}{transition.For}");

                    if (transition.Fields.Length > 0) sb.Append($"\n{pad}Fields: {transition.Fields}");
                }
                else
                {
                    sb.Append($"\n{transition.From,17} -> {transition.To,-17} ({transition.Reasons})");
                    if (transition.For.Length > 0) sb.Append($"> {transition.For}");

                    if (transition.Fields.Length > 0) sb.Append($"\n{pad}Fields: {transition.Fields}");
                }
            }

            return sb.ToString();
        }

        private static void ProcessOneLevelDeeper(
            XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options,
            WorkItemStore workItemStore,
            int recursionLevel,
            List<WorkItem> typeofWI, string workItemType)
        {
            Int64 startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            int totalItems = typeofWI.Count;

            XlHlp.DisplayInWatchWindow($"WorkItem Type: {workItemType} Count:{totalItems} RecursionLevel:{recursionLevel}");

            if (typeofWI.Count > 0)
            {
                int itemCount = 0;
                insertAt.IncrementRows();

                foreach (WorkItem wi in typeofWI)
                {
                    options.ShowIndividualItems = false;

                    Add_TP_WorkItem_WorkItemLinks(insertAt, options, workItemStore, wi, recursionLevel);
                    itemCount++;    // Useful if debugging to see how far we have progressed

                    AZDOHelper.DisplayLoopUpdates(startTicks, options, totalItems, itemCount);
                }
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);
        }

        //private string PrintMappings(List<FieldMap> mappings)
        //{
        //    StringBuilder sb = new StringBuilder();
        //    string pad = new string(' ', 40);

        //    //foreach (var transition in mappings.OrderBy(n => n.From))
        //    //{
        //    //    if (sb.Length == 0)
        //    //    {
        //    //        sb.Append($"{transition.From,17} -> {transition.To,-17} ({transition.Reasons})");
        //    //        if (transition.For.Length > 0) sb.Append($" > {pad}{transition.For}");

        //    //        if (transition.Fields.Length > 0) sb.Append($"\n{pad}Fields: {transition.Fields}");
        //    //    }
        //    //    else
        //    //    {
        //    //        sb.Append($"\n{transition.From,17} -> {transition.To,-17} ({transition.Reasons})");
        //    //        if (transition.For.Length > 0) sb.Append($"> {transition.For}");

        //    //        if (transition.Fields.Length > 0) sb.Append($"\n{pad}Fields: {transition.Fields}");
        //    //    }
        //    //}

        //    return sb.ToString();
        //}

        internal struct ControlFieldMap
        {
            public ControlMap ControlMap { get; set; }
            public FieldMap FieldMap { get; set; }
            public string MapType { get; set; }
        }

        internal struct ControlMap
        {
            public string FieldName { get; set; }
            public string Label { get; set; }
            public string Name { get; set; }
            public string Type { get; set; }
        }

        internal struct FieldMap
        {
            public string Name { get; set; }
            public string RefName { get; set; }
            public string Type { get; set; }
            public bool Required { get; set; }
        }

        internal struct Transition
        {
            public string Fields { get; set; }
            public string For { get; set; }
            public string From { get; set; }
            public string Reasons { get; set; }
            public string To { get; set; }
        }

        #endregion
    }
}
