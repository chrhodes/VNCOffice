using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows;

using Microsoft.Office.Interop.Excel;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.Server;
using Microsoft.TeamFoundation.VersionControl.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;

using SupportTools_Excel.AzureDevOpsExplorer.Domain;

using XlHlp = VNC.AddinHelper.Excel;

namespace SupportTools_Excel.AzureDevOpsExplorer.Application
{
    class Section_VersionControlServer
    {
        #region Version Control Server (VCS)

        internal static XlHlp.XlLocation AddSections(XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options,
            VersionControlServer versionControlServer,
            TeamProject teamProject,
            ProjectInfo projectInfo,
            List<string> sectionsToDisplay)
        {
            long startTicks = XlHlp.DisplayInWatchWindow(insertAt);

            if (sectionsToDisplay.Count != 0)
            {
                if (insertAt.OrientVertical)
                {
                    //XlHlp.AddContentToCell(insertAt.AddRowX(), "Version Control Server (VCS) Information");
                    XlHlp.AddSectionInfo(insertAt.AddRow(), "Version Control Server (VCS) Information", "");
                    insertAt.IncrementRows();
                }
                else
                {
                    //XlHlp.AddContentToCell(insertAt.AddRowX(), "Version Control Server (VCS) Information");
                    XlHlp.AddSectionInfo(insertAt.AddRow(), "Version Control Server (VCS) Information", "",
                        orientation: XlOrientation.xlUpward);
                    insertAt.DecrementRows();   // AddRow bumped it.
                    insertAt.IncrementColumns();
                }

                if (sectionsToDisplay.Contains("Info"))
                {
                    insertAt = Add_Info(insertAt, options, versionControlServer).IncrementPosition(insertAt.OrientVertical);
                }

                if (sectionsToDisplay.Contains("Affected Projects"))
                {
                    insertAt = Add_TP_AffectedTeamProjects(insertAt, options, teamProject).IncrementPosition(insertAt.OrientVertical);
                }

                if (sectionsToDisplay.Contains("Branches"))
                {
                    insertAt = Add_TP_Branches(insertAt, options, versionControlServer, teamProject).IncrementPosition(insertAt.OrientVertical);
                }

                if (sectionsToDisplay.Contains("ChangeSets"))
                {
                    //insertAt = Add_TP_Changesets(insertAt, options,
                    //    teamProject, (bool)ceListChangesetChanges.IsChecked, (bool)ceListChangesetWorkItems.IsChecked);
                }

                if (sectionsToDisplay.Contains("Developers"))
                {
                    insertAt = Add_TP_Developers(insertAt, options, versionControlServer, teamProject, false, null).IncrementPosition(insertAt.OrientVertical);
                }

                if (sectionsToDisplay.Contains("ItemSets"))
                {
                    insertAt = Add_TP_ItemSets(insertAt, options, teamProject).IncrementPosition(insertAt.OrientVertical);
                }

                if (sectionsToDisplay.Contains("PendingSets"))
                {
                    insertAt = Add_TP_PendingSets(insertAt, options, teamProject).IncrementPosition(insertAt.OrientVertical);
                }

                if (sectionsToDisplay.Contains("ShelveSets"))
                {
                    insertAt = Add_TP_ShelveSets(insertAt, options, teamProject).IncrementPosition(insertAt.OrientVertical);
                }

                if (sectionsToDisplay.Contains("Teams"))
                {
                    insertAt = Add_TP_Teams(insertAt, options, versionControlServer, teamProject, projectInfo, false);
                }

                insertAt.IncrementPosition(insertAt.OrientVertical);
            }

            XlHlp.DisplayInWatchWindow(insertAt, startTicks, "End");

            return insertAt;
        }

        /// <summary>
        /// Display branches and recursively descend.
        /// </summary>
        /// <param name="insertAt"></param>
        /// <param name="branch"></param>
        /// <returns></returns>
        internal static XlHlp.XlLocation Display_VCS_Branch(XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options,
            VersionControlServer versionControlServer,
            BranchObject branch)
        {
            long startTicks = XlHlp.DisplayInWatchWindow(insertAt);

            return insertAt;

            try
            {
                insertAt.ClearOffsets();

                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), branch.Properties.RootItem.Item);
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), branch.DateCreated.ToString());

                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), branch.Properties.RootItem.ChangeType.ToString());
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), branch.Properties.RootItem.Version.DisplayString);
                //ExcelHlp.AddContentToCell(insertAt.AddOffsetColumn(), branch.Properties.RootItem.ToString());

                insertAt.IncrementRows();

                var childBranches = versionControlServer.QueryBranchObjects(branch.Properties.RootItem, RecursionType.OneLevel);

                //currentColumn++;

                foreach (BranchObject childBranch in childBranches)
                {
                    insertAt.ClearOffsets();

                    if (childBranch.Properties.RootItem.Item == branch.Properties.RootItem.Item)
                    {
                        continue;
                    }

                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), childBranch.Properties.RootItem.Item);
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), childBranch.DateCreated.ToString());

                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), childBranch.Properties.RootItem.ChangeType.ToString());
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), childBranch.Properties.RootItem.Version.DisplayString);

                    insertAt.IncrementRows();

                    if (childBranch.ChildBranches.Count() > 0)
                    {
                        insertAt = Display_VCS_Branch(insertAt, options, versionControlServer, childBranch);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            XlHlp.DisplayInWatchWindow(insertAt, startTicks, "End");

            return insertAt;
        }

        internal static XlHlp.XlLocation Add_Info(XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options, 
            VersionControlServer versionControlServer)
        {
            long startTicks = XlHlp.DisplayInWatchWindow(insertAt);

            try
            {
                insertAt.MarkStart(XlHlp.MarkType.Group);

                XlHlp.AddLabeledInfoX(insertAt.AddRowX(2), "ServerGuid:", versionControlServer.ServerGuid.ToString());
                XlHlp.AddLabeledInfoX(insertAt.AddRowX(2), "TeamProjectCollection:", versionControlServer.TeamProjectCollection.Name);
                XlHlp.AddLabeledInfoX(insertAt.AddRowX(2), "WebServiceLevel:", versionControlServer.WebServiceLevel.ToString());

                int changesetId = versionControlServer.GetLatestChangesetId();
                Changeset changeset = versionControlServer.GetChangeset(changesetId);
                string creationDate = "";

                if (changeset != null)
                {
                    creationDate = changeset.CreationDate.ToString();
                }

                //XlHlp.AddTitledInfo(insertAt.AddRow(2), "LatestChangeSetId:", changesetId.ToString());
                //insertAt.DecrementRows();   // Backup so can add more info
                //XlHlp.AddContentToCell(insertAt.AddOffsetColumn(), creationDate);
                //insertAt.ClearOffsets();
                //insertAt.IncrementRows();

                XlHlp.AddLabeledInfoX(insertAt.AddRowX(2), "LatestChangeSetId:", changesetId.ToString());
                //insertAt.DecrementRows();   // Backup so can add more info
                XlHlp.AddContentToCell(insertAt.GetCurrentRange().Offset[-1, 2], creationDate);

                insertAt.MarkEnd(XlHlp.MarkType.Group);

                insertAt.Group(insertAt.OrientVertical, hide: true);

                insertAt.EndSectionAndSetNextLocation(insertAt.OrientVertical);

                XlHlp.DisplayInWatchWindow(insertAt, startTicks, "End");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            return insertAt;
        }

        internal static XlHlp.XlLocation Add_RootBranches(XlHlp.XlLocation insertAt, 
            Options_AZDO_TFS options,
            VersionControlServer versionControlServer)
        {
            long startTicks = XlHlp.DisplayInWatchWindow(insertAt);

            try
            {
                int count = 0;

                var rootBranches = versionControlServer.QueryRootBranchObjects(RecursionType.None);
                count = rootBranches.Count();
                XlHlp.AddLabeledInfoX(insertAt.AddRowX(), "QueryRootBranchObjects(None)", count.ToString(), 40);
                insertAt = DisplayListOf_Branches(insertAt, rootBranches, false, "None");

                insertAt.IncrementRows();

                XlHlp.DisplayInWatchWindow("After QueryRootBranchObjects(none)");

                rootBranches = versionControlServer.QueryRootBranchObjects(RecursionType.OneLevel);
                count = rootBranches.Count();
                XlHlp.AddLabeledInfoX(insertAt.AddRowX(), "QueryRootBranchObjects(One)", count.ToString(), 40);
                insertAt = DisplayListOf_Branches(insertAt, rootBranches, false, "OneLevel");

                insertAt.IncrementRows();

                XlHlp.DisplayInWatchWindow("After QueryRootBranchObjects(one)");

                rootBranches = versionControlServer.QueryRootBranchObjects(RecursionType.Full);
                count = rootBranches.Count();
                XlHlp.AddLabeledInfoX(insertAt.AddRowX(), "QueryRootBranchObjects(Full)", count.ToString(), 40);
                insertAt = DisplayListOf_Branches(insertAt, rootBranches, false, "Full");

                XlHlp.DisplayInWatchWindow("After QueryRootBranchObjects(Full)");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            XlHlp.DisplayInWatchWindow(insertAt, "End");

            return insertAt;
        }

        internal static XlHlp.XlLocation Add_Shelvesets(XlHlp.XlLocation insertAt, 
            Options_AZDO_TFS options,
            VersionControlServer versionControlServer)
        {
            long startTicks = XlHlp.DisplayInWatchWindow(insertAt);

            try
            {
                Worksheet ws = insertAt.workSheet;

                // QueryShelvesets(shelvesetName, shelvesetOwner)
                // QueryShelvesets(null, null) returns all shelveets for all owners.

                Shelveset[] shelvesets = versionControlServer.QueryShelvesets(null, null);

                int count = shelvesets.Count();

                if (insertAt.OrientVertical)
                {
                    XlHlp.AddLabeledInfoX(insertAt.AddRowX(), "ShelveSets", count.ToString());
                }
                else
                {
                    XlHlp.AddLabeledInfoX(insertAt.AddRowX(), "ShelveSets", count.ToString(), orientation: XlOrientation.xlUpward);
                    insertAt.IncrementColumns();
                }

                insertAt.MarkStart(XlHlp.MarkType.GroupTable);

                Header_VersionControlServer.Add_Shelvesets(insertAt);

                Body_VersionControlServer.Add_TP_Shelvesets(insertAt, options, shelvesets);

                insertAt.MarkEnd(XlHlp.MarkType.GroupTable, string.Format("tblShelveSets_{0}", ws.Name));

                insertAt.Group(insertAt.OrientVertical, hide: true);

                insertAt.EndSectionAndSetNextLocation(insertAt.OrientVertical);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            XlHlp.DisplayInWatchWindow(insertAt, startTicks, "End");

            return insertAt;
        }

        internal static XlHlp.XlLocation Add_TP_AffectedTeamProjects(XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options,
            TeamProject teamProject)
        {
            long startTicks = XlHlp.DisplayInWatchWindow(insertAt);

            try
            {
                int count = 0;

                if (insertAt.OrientVertical)
                {
                    //XlHlp.AddLabeledInfoX(insertAt.AddRowX(), "AffectedTeamProjects", count.ToString());
                    XlHlp.AddSectionInfo(insertAt.AddRow(), "AffectedTeamProjects", $"{ count }");
                }
                else
                {
                    //XlHlp.AddLabeledInfoX(insertAt.AddRowX(), "AffectedTeamProjects", count.ToString(), orientation: XlOrientation.xlUpward);
                    XlHlp.AddSectionInfo(insertAt.AddRow(), "AffectedTeamProjects", $"{ count }",
                        orientation: XlOrientation.xlUpward);
                    insertAt.IncrementColumns();
                }

                insertAt.MarkStart(XlHlp.MarkType.GroupTable);

                XlHlp.AddContentToCell(insertAt.AddRowX(), "Not Implemented Yet");

                insertAt.MarkEnd(XlHlp.MarkType.GroupTable, string.Format("tblAffectedTP_{0}", teamProject.Name));

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

        internal static XlHlp.XlLocation Add_TP_Branches(XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options,
            VersionControlServer versionControlServer,
            TeamProject teamProject)
        {
            long startTicks = XlHlp.DisplayInWatchWindow(insertAt);

            try
            {
                int count = 0;

                if (insertAt.OrientVertical)
                {
                    //XlHlp.AddLabeledInfoX(insertAt.AddRowX(), "Branches", count.ToString());
                    XlHlp.AddSectionInfo(insertAt.AddRow(), "Branches", $"{ count }");
                }
                else
                {
                    //XlHlp.AddLabeledInfoX(insertAt.AddRowX(), "Branches", count.ToString(), orientation: XlOrientation.xlUpward);
                    XlHlp.AddSectionInfo(insertAt.AddRow(), "Branches", $"{ count }",
                        orientation: XlOrientation.xlUpward);
                    insertAt.IncrementColumns();
                }

                insertAt.MarkStart(XlHlp.MarkType.GroupTable);

                // The number of levels to traverse. 
                // None will only return the branch object.
                // OneLevel will return children.
                // Full will return all descendants.

                var rootBranchesOneLevel = versionControlServer.QueryRootBranchObjects(RecursionType.OneLevel);
                var rootBranchesNone = versionControlServer.QueryRootBranchObjects(RecursionType.None);
                var rootBranchesFull = versionControlServer.QueryRootBranchObjects(RecursionType.Full);

                var branchesOneLevel = versionControlServer.QueryBranchObjects(new ItemIdentifier(teamProject.ServerItem), RecursionType.OneLevel);
                var branchesNone = versionControlServer.QueryBranchObjects(new ItemIdentifier(teamProject.ServerItem), RecursionType.None);
                var branchesFull = versionControlServer.QueryBranchObjects(new ItemIdentifier(teamProject.ServerItem), RecursionType.Full);

                //Array.ForEach(rootBranches, (bo) => DisplayAllBranches(bo, vcs, currentColumn));

                if (rootBranchesOneLevel.Count() > 0)
                {
                    foreach (BranchObject branch in rootBranchesOneLevel)
                    {
                        insertAt = Display_VCS_Branch(insertAt, options, versionControlServer, branch);
                    }
                    //insertAt = Add_Branches(insertAt, rootBranchesOneLevel);
                }
                else
                {
                    XlHlp.AddContentToCell(insertAt.AddRowX(), "<None>");
                }

                insertAt.MarkEnd(XlHlp.MarkType.GroupTable, string.Format("tblBranches_{0}", teamProject.Name));

                insertAt.Group(insertAt.OrientVertical, hide: true);

                insertAt.EndSectionAndSetNextLocation(insertAt.OrientVertical);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            XlHlp.DisplayInWatchWindow(insertAt, startTicks, "End");

            return insertAt;
        }

        internal static XlHlp.XlLocation Add_TP_Changesets(XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options,
            VersionControlServer versionControlServer,
            ICommonStructureService commonStructureService,
            TeamProject teamProject,
            bool listChanges,
            bool listWorkItems)
        {
            long startTicks = XlHlp.DisplayInWatchWindow(insertAt);

            try
            {
                Worksheet ws = insertAt.workSheet;

                System.Collections.IEnumerable history =
                    versionControlServer.QueryHistory(
                        teamProject.ServerItem,
                        LatestVersionSpec.Instance,
                        0,
                        RecursionType.Full,
                        null,                       // Any User
                        new DateVersionSpec(DateTime.Now - TimeSpan.FromDays(options.GoBackDays)),
                        LatestVersionSpec.Instance,
                        Int32.MaxValue,
                        true,                       // includeChanges
                        false);

                if (insertAt.OrientVertical)
                {
                    XlHlp.AddLabeledInfoX(insertAt.AddRowX(), "Changesets", history.Cast<object>().ToList().Count().ToString());
                }
                else
                {
                    XlHlp.AddLabeledInfoX(insertAt.AddRowX(), "Changesets", history.Cast<object>().ToList().Count().ToString(), orientation: XlOrientation.xlUpward);
                    insertAt.IncrementColumns();
                }

                insertAt.MarkStart(XlHlp.MarkType.GroupTable);

                Header_VersionControlServer.Add_TP_Changesets(insertAt);

                Body_VersionControlServer.Add_TP_Changesets(insertAt, options, commonStructureService, listChanges, listWorkItems, history);

                insertAt.MarkEnd(XlHlp.MarkType.GroupTable, string.Format("tblChangesets_{0}", ws.Name));

                insertAt.Group(insertAt.OrientVertical, hide: true);

                insertAt.EndSectionAndSetNextLocation(insertAt.OrientVertical);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            XlHlp.DisplayInWatchWindow(insertAt, startTicks, "End");

            return insertAt;
        }

        internal static XlHlp.XlLocation Add_TP_Developers(XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options,
            VersionControlServer versionControlServer,
            TeamProject teamProject,
            bool displayDataOnly,
            string teamProjectName)
        {
            long startTicks = XlHlp.DisplayInWatchWindow(insertAt);

            try
            {
                SortedDictionary<string, int> developers = new SortedDictionary<string, int>();
                SortedDictionary<string, DateTime> developersLatestDate = new SortedDictionary<string, DateTime>();
                SortedDictionary<string, DateTime> developersEarliestDate = new SortedDictionary<string, DateTime>();

                //Following will get all Changesets in last goBackDays. 

                GetDevelopersWithChangesets(versionControlServer, teamProject, options.GoBackDays, developers, developersLatestDate, developersEarliestDate);

                int count = developers.Count;

                if (!displayDataOnly)
                {

                    if (insertAt.OrientVertical)
                    {
                        //XlHlp.AddLabeledInfoX(insertAt.AddRowX(),
                        //    string.Format("Developers (Changesets(Last {0} days)", options.GoBackDays), count.ToString());
                        XlHlp.AddSectionInfo(insertAt.AddRow(), $"Developers(Changesets(Last { options.GoBackDays } days)", $"{ count }");
                    }
                    else
                    {
                        //XlHlp.AddLabeledInfoX(insertAt.AddRowX(),
                            //string.Format("Developers (Changesets(Last {0} days)", options.GoBackDays), count.ToString(), orientation: XlOrientation.xlUpward);

                        XlHlp.AddSectionInfo(insertAt.AddRow(), $"Developers(Changesets(Last { options.GoBackDays } days)", $"{ count }",
                            orientation: XlOrientation.xlUpward);
                        insertAt.IncrementColumns();
                    }
                };

                insertAt.ClearOffsets();

                if (!displayDataOnly)
                {
                    insertAt.MarkStart(XlHlp.MarkType.GroupTable);

                    Header_VersionControlServer.Add_TP_Developers(insertAt);
                }

                Body_VersionControlServer.Add_TP_Developers(insertAt, options, teamProjectName, developers, developersLatestDate, developersEarliestDate);

                if (!displayDataOnly)
                {
                    insertAt.MarkEnd(XlHlp.MarkType.GroupTable, string.Format("tblDevelopers_{0}", teamProject.Name));

                    insertAt.Group(insertAt.OrientVertical, hide: true);
                }

                insertAt.EndSectionAndSetNextLocation(insertAt.OrientVertical);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            XlHlp.DisplayInWatchWindow(insertAt, startTicks, "End");

            return insertAt;
        }

        internal static XlHlp.XlLocation Add_TP_ItemSets(XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options,
            TeamProject teamProject)
        {
            long startTicks = XlHlp.DisplayInWatchWindow(insertAt);

            try
            {
                int count = 0;

                if (insertAt.OrientVertical)
                {
                    XlHlp.AddLabeledInfoX(insertAt.AddRowX(), "ItemSets", count.ToString());
                }
                else
                {
                    XlHlp.AddLabeledInfoX(insertAt.AddRowX(), "ItemSets", count.ToString(), orientation: XlOrientation.xlUpward);
                    insertAt.IncrementColumns();
                }

                insertAt.MarkStart(XlHlp.MarkType.GroupTable);

                XlHlp.AddContentToCell(insertAt.AddRowX(), "Not Implemented Yet");

                insertAt.MarkEnd(XlHlp.MarkType.GroupTable);

                insertAt.Group(insertAt.OrientVertical);

                if (!insertAt.OrientVertical)
                {
                    // Skip past the info just added.
                    insertAt.SetLocation(insertAt.RowStart, insertAt.TableEndColumn + 1);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            XlHlp.DisplayInWatchWindow(insertAt, startTicks, "End");

            return insertAt;
        }

        internal static XlHlp.XlLocation Add_TP_PendingSets(XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options,
            TeamProject teamProject)
        {
            long startTicks = XlHlp.DisplayInWatchWindow(insertAt);

            try
            {
                //PendingSet[] pendingChanges = versionControlServer.GetPendingSets(new String[] { teamProject.ServerItem }, RecursionType.Full);
                //PendingSet[] pendingChanges = versionControlServer.GetPendingSets(null, RecursionType.Full);

                // TODO: Pass in parameters
                //PendingSet[] pendingSets = versionControlServer.QueryPendingSets(new[] { "$/" }, RecursionType.Full, "CHRDEV1I", "Christopher Rhodes");

                //PendingChange[] pendingChanges = pendingSets.First().PendingChanges;

                //count = pendingChanges.Count();

                //ExcelHlp.AddTitledInfo(rngOutput.Offset[XlLocation.Rows++, 0], "PendingSets", count.ToString());

                //foreach (PendingChange pendingChange in pendingChanges)
                //{
                //    int col = 1;

                //    ExcelHlp.AddContentToCell(rngOutput.Offset[XlLocation.Rows, col++], pendingChange.ChangeType.ToString());
                //    ExcelHlp.AddContentToCell(rngOutput.Offset[XlLocation.Rows, col++], pendingChange.FileName);
                //    ExcelHlp.AddContentToCell(rngOutput.Offset[XlLocation.Rows, col++], pendingChange.ServerItem);
                //    ExcelHlp.AddContentToCell(rngOutput.Offset[XlLocation.Rows, col++], pendingChange.CreationDate.ToString());
                //    //ExcelHlp.AddContentToCell(rngOutput.Offset[XlLocation.Rows, col++], pendingChange.Type.ToString());

                //    XlLocation.Rows++;
                //}

                int count = 0;

                if (insertAt.OrientVertical)
                {
                    XlHlp.AddLabeledInfoX(insertAt.AddRowX(), "PendingSets", count.ToString());
                }
                else
                {
                    XlHlp.AddLabeledInfoX(insertAt.AddRowX(), "PendingSets", count.ToString(), orientation: XlOrientation.xlUpward);
                    insertAt.IncrementColumns();
                }

                insertAt.MarkStart(XlHlp.MarkType.GroupTable);

                XlHlp.AddContentToCell(insertAt.AddRowX(), "Not Implemented Yet");

                insertAt.MarkEnd(XlHlp.MarkType.GroupTable);
                //insertAt.MarkEnd(XlHlp.MarkType.GroupTable, string.Format("tblPendingSetsTP_{0}", teamProject.Name));

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

        internal static XlHlp.XlLocation Add_TP_ShelveSets(XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options,
            TeamProject teamProject)
        {
            long startTicks = XlHlp.DisplayInWatchWindow(insertAt);

            try
            {
                //PendingSet[] pendingChanges = versionControlServer.GetPendingSets(new String[] { teamProject.ServerItem }, RecursionType.Full);
                //PendingSet[] pendingChanges = versionControlServer.GetPendingSets(null, RecursionType.Full);

                // TODO: Pass in parameters
                //PendingSet[] pendingSets = versionControlServer.QueryPendingSets(new[] { "$/" }, RecursionType.Full, "CHRDEV1I", "Christopher Rhodes");

                //PendingChange[] pendingChanges = pendingSets.First().PendingChanges;

                //count = pendingChanges.Count();

                //ExcelHlp.AddTitledInfo(rngOutput.Offset[XlLocation.Rows++, 0], "PendingSets", count.ToString());

                //foreach (PendingChange pendingChange in pendingChanges)
                //{
                //    int col = 1;

                //    ExcelHlp.AddContentToCell(rngOutput.Offset[XlLocation.Rows, col++], pendingChange.ChangeType.ToString());
                //    ExcelHlp.AddContentToCell(rngOutput.Offset[XlLocation.Rows, col++], pendingChange.FileName);
                //    ExcelHlp.AddContentToCell(rngOutput.Offset[XlLocation.Rows, col++], pendingChange.ServerItem);
                //    ExcelHlp.AddContentToCell(rngOutput.Offset[XlLocation.Rows, col++], pendingChange.CreationDate.ToString());
                //    //ExcelHlp.AddContentToCell(rngOutput.Offset[XlLocation.Rows, col++], pendingChange.Type.ToString());

                //    XlLocation.Rows++;
                //}
                int count = 0;

                if (insertAt.OrientVertical)
                {
                    //XlHlp.AddLabeledInfoX(insertAt.AddRowX(), "ShelveSets", count.ToString());
                    XlHlp.AddSectionInfo(insertAt.AddRow(), "ShelveSets", $"{ count }");
                }
                else
                {
                    //XlHlp.AddLabeledInfoX(insertAt.AddRowX(), "ShelveSets", count.ToString(), orientation: XlOrientation.xlUpward);
                    XlHlp.AddSectionInfo(insertAt.AddRow(), "ShelveSets", $"{ count }",
                        orientation: XlOrientation.xlUpward);
                    insertAt.IncrementColumns();
                }

                insertAt.MarkStart(XlHlp.MarkType.GroupTable);

                XlHlp.AddContentToCell(insertAt.AddRowX(), "Not Implemented Yet");

                insertAt.MarkEnd(XlHlp.MarkType.GroupTable);
                //insertAt.MarkEnd(XlHlp.MarkType.GroupTable, string.Format("tblShelveSetsTP_{0}", teamProject.Name));

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

        internal static XlHlp.XlLocation Add_TP_Teams(XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options,
            VersionControlServer versionControlServer,
            TeamProject teamProject,
            ProjectInfo projectInfo,
            bool displayDataOnly)
        {
            long startTicks = XlHlp.DisplayInWatchWindow(insertAt);
            
            try
            {
                //ProjectInfo teamProjectInfo = Server.CommonStructureService.GetProjectFromName(teamProject.Name);
                var tpUri = projectInfo.Uri;

                TfsTeamService teamService = new TfsTeamService();

                teamService.Initialize(teamProject.TeamProjectCollection);

                var defaultTeam = teamService.GetDefaultTeam(tpUri, new List<String>());

                IEnumerable<TeamFoundationTeam> allTeams = teamService.QueryTeams(tpUri);

                if (insertAt.OrientVertical)
                {
                    //XlHlp.AddLabeledInfoX(insertAt.AddRowX(), "Teams", allTeams.Count().ToString());
                    XlHlp.AddSectionInfo(insertAt.AddRow(), "Teams", $"{ allTeams.Count() }");
                }
                else
                {
                    //XlHlp.AddLabeledInfoX(insertAt.AddRowX(), "Teams", allTeams.Count().ToString(), orientation: XlOrientation.xlUpward);
                    XlHlp.AddSectionInfo(insertAt.AddRow(), "Teams", $"{ allTeams.Count() }",
                        orientation: XlOrientation.xlUpward);
                    insertAt.IncrementColumns();
                }

                if (!displayDataOnly)
                {
                    insertAt.MarkStart(XlHlp.MarkType.GroupTable);

                    Header_VersionControlServer.Add_TP_Teams(insertAt);
                }

                Body_VersionControlServer.Add_TP_Teams(insertAt, options, versionControlServer, teamProject, allTeams, defaultTeam);

                if (!displayDataOnly)
                {
                    insertAt.MarkEnd(XlHlp.MarkType.GroupTable, string.Format("tblTeams_{0}", teamProject.Name));

                    insertAt.Group(insertAt.OrientVertical, hide: true);
                }

                if (!insertAt.OrientVertical)
                {
                    // Skip past the info just added.
                    insertAt.SetLocation(insertAt.RowStart, insertAt.TableEndColumn + 1);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            XlHlp.DisplayInWatchWindow(insertAt, startTicks, "End");

            return insertAt;
        }

        internal static XlHlp.XlLocation Add_Workspaces(XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options,
            VersionControlServer versionControlServer)
        {
            long startTicks = XlHlp.DisplayInWatchWindow(insertAt);

            try
            {
                Worksheet ws = insertAt.workSheet;

                // QueryWorkspaces(workspaceName, workspaceOwner, computer)
                // QueryWorkspaces(null, null, null) returns all workspaces for all owners for all computers.

                Workspace[] workSpaces = versionControlServer.QueryWorkspaces(null, null, null);

                int count = workSpaces.Count();

                if (insertAt.OrientVertical)
                {
                    //XlHlp.AddLabeledInfoX(insertAt.AddRowX(), "Workspaces", count.ToString());
                    XlHlp.AddSectionInfo(insertAt.AddRow(), "Workspaces", $"{ count }");
                    //XlHlp.AddContentToCell(insertAt.AddRow(), defaultTeam.Name);
                }
                else
                {
                    XlHlp.AddLabeledInfoX(insertAt.AddRowX(), "Workspaces", count.ToString(), orientation: XlOrientation.xlUpward);
                    XlHlp.AddSectionInfo(insertAt.AddRow(), "Workspaces", $"{ count }",
                        orientation: XlOrientation.xlUpward);
                    insertAt.IncrementColumns();
                }

                insertAt.MarkStart(XlHlp.MarkType.GroupTable);

                // Keep in same order as fields, infra.

                Header_VersionControlServer.Add_TP_Workspaces(insertAt);

                Body_VersionControlServer.Add_TP_Workspaces(insertAt, options, workSpaces);

                insertAt.MarkEnd(XlHlp.MarkType.GroupTable, string.Format("tblWorkSpaces_{0}", ws.Name));

                insertAt.Group(insertAt.OrientVertical, hide: true);

                insertAt.EndSectionAndSetNextLocation(insertAt.OrientVertical);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            XlHlp.DisplayInWatchWindow(insertAt, startTicks, "End");

            return insertAt;
        }

        internal static void Display_VCS_ChangeSet_AssociatedWorkItems(Changeset changeSet)
        {
            long startTicks = XlHlp.DisplayInWatchWindow(string.Format("  AssociatedWorkItems.Count: {0}", changeSet.AssociatedWorkItems.Count()));

            foreach (AssociatedWorkItemInfo item in changeSet.AssociatedWorkItems)
            {
                XlHlp.DisplayInWatchWindow(string.Format("    ID: {0}  Title: {1}  AssignedTo: {2}  WorkItemType: {3}  State: {4}",
                    item.Id,
                    item.Title,
                    item.AssignedTo,
                    item.WorkItemType,
                    item.State));
            }
        }

        internal static void Display_VCS_ChangeSet_Changes(Changeset changeSet)
        {
            long startTicks = XlHlp.DisplayInWatchWindow(string.Format("  Changes.Count: {0}", changeSet.Changes.Count()));

            foreach (Change change in changeSet.Changes)
            {
                XlHlp.DisplayInWatchWindow(string.Format("    ChangeType: {0}  CheckinDate: {1}  IsBranch: {2}  ItemType: {3}",
                    change.ChangeType,
                    change.Item.CheckinDate,
                    change.Item.IsBranch,
                    change.Item.ItemType));

                XlHlp.DisplayInWatchWindow(string.Format("      ServerItem: {0}", change.Item.ServerItem));

                if (change.MergeSources != null)
                {
                    XlHlp.DisplayInWatchWindow(string.Format("      MergeSources.Count: {0}", change.MergeSources.Count()));
                }
            }
        }

        internal static void Display_VCS_Changeset_Info(Changeset changeSet)
        {
            XlHlp.DisplayInWatchWindow(string.Format(" Committer: {0}", changeSet.Committer));
            XlHlp.DisplayInWatchWindow(string.Format(" CreationDate: {0}", changeSet.CreationDate));
            XlHlp.DisplayInWatchWindow(string.Format(" Owner: {0}", changeSet.Owner));
            XlHlp.DisplayInWatchWindow(string.Format(" CheckinNote: {0}", changeSet.CheckinNote));
            XlHlp.DisplayInWatchWindow(string.Format(" Comment: {0}", changeSet.Comment));
        }

        internal static void Display_VCS_Changeset_WorkItems(Changeset changeSet, WorkItemStore workItemStore, ICommonStructureService commonStructureService)
        {
            long startTicks = XlHlp.DisplayInWatchWindow(string.Format("  WorkItems.Count: {0}", changeSet.WorkItems.Count()));

            foreach (WorkItem item in changeSet.WorkItems)
            {
                XlHlp.DisplayInWatchWindow(string.Format("    ID: {0}   Title: {1}", item.Id, item.Title));
                XlHlp.DisplayInWatchWindow(string.Format("    AuthorizedDate: {0}", item.AuthorizedDate));
                XlHlp.DisplayInWatchWindow(string.Format("    AreadId: {0}  AreaPath: {1}", item.AreaId, item.AreaPath));

                string iterationInfo = GetIterationInfo(item, commonStructureService);

                XlHlp.DisplayInWatchWindow(string.Format("    {0}", iterationInfo));

                XlHlp.DisplayInWatchWindow(string.Format("    AttachedFileCount: {0}", item.AttachedFileCount));

                foreach (Attachment attachment in item.Attachments)
                {
                    XlHlp.DisplayInWatchWindow(string.Format("       ID: {0}  Name: {1}", attachment.Id, attachment.Name));
                    XlHlp.DisplayInWatchWindow(string.Format("         CreationTime: {0}  AttachedTime: {0}", attachment.CreationTime, attachment.AttachedTime));
                    XlHlp.DisplayInWatchWindow(string.Format("         Comment: {0}", attachment.Comment));
                }

                XlHlp.DisplayInWatchWindow(string.Format("    CreatedBy: {0}  CreatedDate: {1}", item.CreatedBy, item.CreatedDate));
                XlHlp.DisplayInWatchWindow(string.Format("    ChangedBy: {0}  ChangedDate: {1}", item.ChangedBy, item.ChangedDate));
                XlHlp.DisplayInWatchWindow(string.Format("    Description: {0}", item.Description));


                XlHlp.DisplayInWatchWindow(string.Format("    Reason: {0}", item.Reason));

                XlHlp.DisplayInWatchWindow(string.Format("    Rev: {0}", item.Rev.ToString()));
                XlHlp.DisplayInWatchWindow(string.Format("    RevisedDate: {0}", item.RevisedDate));
                XlHlp.DisplayInWatchWindow(string.Format("    Revsion: {0}", item.Revision));

                //ExcelHlp.DisplayInWatchWindow(string.Format("    Revisons.Count: {0}", item.Revisions.Count));

                //foreach (Revision revsion in item.Revisions)
                //{
                //    ExcelHlp.DisplayInWatchWindow(string.Format("        Fields.Count: {0}", revsion.Fields.Count));
                //}

                XlHlp.DisplayInWatchWindow(string.Format("    State: {0}", item.State));
                XlHlp.DisplayInWatchWindow(string.Format("    Type: {0}", item.Type.Name));

                XlHlp.DisplayInWatchWindow(string.Format("    ExternalLinkCount: {0}  HyperLinkCount: {1}  RelatedLinkCount: {2}",
                    item.ExternalLinkCount, item.HyperLinkCount, item.RelatedLinkCount));

                XlHlp.DisplayInWatchWindow(string.Format("    Links.Count: {0}", item.Links.Count));

                foreach (Link link in item.Links)
                {
                    XlHlp.DisplayInWatchWindow(string.Format("        ArtifactLinkType.Name: {0}", link.ArtifactLinkType.Name));
                    XlHlp.DisplayInWatchWindow(string.Format("        Comment: {0}", link.Comment));
                }

                XlHlp.DisplayInWatchWindow(string.Format("    WorkItemLinks.Count: {0}", item.WorkItemLinks.Count));

                foreach (WorkItemLink link in item.WorkItemLinks)
                {
                    XlHlp.DisplayInWatchWindow(string.Format("        AddedBy: {0}  AddedDate: {0}", link.AddedBy, link.AddedDate));
                    XlHlp.DisplayInWatchWindow(string.Format("        BaseType: {0}", link.BaseType));
                    XlHlp.DisplayInWatchWindow(string.Format("        ChangedDate: {0}", link.ChangedDate));
                    XlHlp.DisplayInWatchWindow(string.Format("        Comment: {0}", link.Comment));
                    XlHlp.DisplayInWatchWindow(string.Format("        LinkTypeEnd.Id: {0}", link.LinkTypeEnd.Id));
                    XlHlp.DisplayInWatchWindow(string.Format("        LinkTypeEnd.ImmutableName: {0}", link.LinkTypeEnd.ImmutableName));
                    XlHlp.DisplayInWatchWindow(string.Format("        LinkTypeEnd.IsForwardLink: {0}", link.LinkTypeEnd.IsForwardLink));
                    XlHlp.DisplayInWatchWindow(string.Format("        LinkTypeEnd.LinkType: {0}", link.LinkTypeEnd.LinkType));
                    XlHlp.DisplayInWatchWindow(string.Format("        SourceId: {0}", link.SourceId));
                    XlHlp.DisplayInWatchWindow(string.Format("        TargetId: {0}", link.TargetId));

                    WorkItem targetWorkItem = workItemStore.GetWorkItem(link.TargetId);

                    XlHlp.DisplayInWatchWindow(string.Format("            ID: {0}  Title: {1}", targetWorkItem.Id, targetWorkItem.Title));
                    XlHlp.DisplayInWatchWindow(string.Format("            CreatedBy: {0}  CreatedDate: {1}", targetWorkItem.CreatedBy, targetWorkItem.CreatedDate));
                    XlHlp.DisplayInWatchWindow(string.Format("            Description: {0}", targetWorkItem.Description));
                    XlHlp.DisplayInWatchWindow(string.Format("            Reason: {0}", targetWorkItem.Reason));
                    XlHlp.DisplayInWatchWindow(string.Format("            Rev: {0}", targetWorkItem.Rev));
                    XlHlp.DisplayInWatchWindow(string.Format("            RevisedDate: {0}", targetWorkItem.RevisedDate));
                    XlHlp.DisplayInWatchWindow(string.Format("            State: {0}", targetWorkItem.State));
                    XlHlp.DisplayInWatchWindow(string.Format("            Type: {0}", targetWorkItem.Type));
                    XlHlp.DisplayInWatchWindow(string.Format("            AreaPath: {0}", targetWorkItem.AreaPath));
                    XlHlp.DisplayInWatchWindow(string.Format("            IterationPath: {0}", targetWorkItem.IterationPath));
                }
            }
        }

        internal static string GetChangeInfo(Microsoft.TeamFoundation.VersionControl.Client.Change change)
        {
            long startTicks = XlHlp.DisplayInWatchWindow("Begin");

            StringBuilder changeInfo = new StringBuilder();

            changeInfo.AppendFormat("{0,-10} > {1}  CheckinDate: {2}  IsBranch: {3}  ItemType: {4}",
                    "Change",
                    change.ChangeType,
                    change.Item.CheckinDate,
                    change.Item.IsBranch,
                    change.Item.ItemType);

            changeInfo.AppendFormat("  ServerItem: {0}", change.Item.ServerItem);

            if (change.MergeSources != null)
            {
                changeInfo.AppendFormat("  MergeSources.Count: {0}", change.MergeSources.Count());
            }

            XlHlp.DisplayInWatchWindow("End", startTicks);

            return changeInfo.ToString();
        }

        internal static string GetIterationInfo(WorkItem workItem, ICommonStructureService commonStructureService)
        {
            long startTicks = XlHlp.DisplayInWatchWindow("Begin");
            // TODO(crhodes):
            //	On some planet this must make sense :)

            var projectNameIndex = workItem.IterationPath.IndexOf("\\", 2);

            var iterationPath = "";

            if (projectNameIndex > 0)
            {
                iterationPath = workItem.IterationPath.Insert(projectNameIndex, "\\Iteration");
            }
            else
            {
                // Directly under Team Project
                iterationPath = string.Format("{0}\\Iteration", workItem.IterationPath);
            }

            // ProjectName\\Iteration\\<Iteration Path>>

            Uri itemUri = workItem.Uri;

            NodeInfo iteration = null;

            try
            {
                iteration = commonStructureService.GetNodeFromPath(iterationPath);
            }
            catch (Exception ex)
            {
                string result = ex.ToString();
            }

            string startDate = iteration.StartDate.HasValue ? ((DateTime)iteration.StartDate).ToShortDateString() : "<null>";
            string finishDate = iteration.FinishDate.HasValue ? ((DateTime)iteration.FinishDate).ToShortDateString() : "<null>";

            string iterationInfo = string.Format("{0,-10} > {2} to {3}  IterationId:{1}  IterationPath:{4}",
                "Iteration",
                workItem.IterationId,
                startDate, finishDate,
                workItem.IterationPath);

            XlHlp.DisplayInWatchWindow("End", startTicks);

            return iterationInfo;
        }

        internal static string GetWorkItemInfo(WorkItem workItem)
        {
            long startTicks = XlHlp.DisplayInWatchWindow("Begin");
            StringBuilder workItemInfo = new StringBuilder();

            workItemInfo.AppendFormat("{0,-10} > ID: {1}  Title: {2}  AreaPath: {3}  IterationPath: {4}",
                "WorkItem",
                workItem.Id,
                workItem.Title,
                workItem.AreaPath, workItem.IterationPath);

            XlHlp.DisplayInWatchWindow("End", startTicks);

            return workItemInfo.ToString();
        }

        private static XlHlp.XlLocation DisplayListOf_Branches(XlHlp.XlLocation insertAt, BranchObject[] branches, bool displayDataOnly, string tableSuffix)
        {
            long startTicks = XlHlp.DisplayInWatchWindow(insertAt);

            Worksheet ws = insertAt.workSheet;

            if (!displayDataOnly)
            {
                insertAt.MarkStart(XlHlp.MarkType.GroupTable);

                XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 40, "RootItem.Item");
                XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 30, "ParentBranch.Item");
                XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 15, "DateCreated");
                XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 20, "Owner");
                XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 20, "Owner DisplayName");
                XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 15, "ChangeType");
                XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 10, "Version");
                XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 20, "Related Branches");

                insertAt.IncrementRows();
            }

            foreach (BranchObject branch in branches)
            {
                insertAt.ClearOffsets();

                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), branch.Properties.RootItem.Item);

                string parentBranch = branch.Properties.ParentBranch != null ? branch.Properties.ParentBranch.Item : "<none>";

                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), parentBranch);
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), branch.DateCreated.ToString());
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), branch.Properties.Owner);
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), branch.Properties.OwnerDisplayName);
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), branch.Properties.RootItem.ChangeType.ToString());
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), branch.Properties.RootItem.Version.DisplayString);
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), branch.RelatedBranches.Count().ToString());

                insertAt.IncrementRows();
            }

            if (!displayDataOnly)
            {
                insertAt.MarkEnd(XlHlp.MarkType.GroupTable, string.Format("tblRootBranches_{0}", tableSuffix));

                insertAt.Group(insertAt.OrientVertical, hide: true);
            }

            XlHlp.DisplayInWatchWindow(insertAt, startTicks, "End");

            return insertAt;
        }

        private static void GetDevelopersWithChangesets(
            VersionControlServer versionControlServer,
            TeamProject teamProject,
            int goBackDays,
            SortedDictionary<string, int> developers,
            SortedDictionary<string, DateTime> developersLatestDate,
            SortedDictionary<string, DateTime> developersEarliestDate)
        {
            long startTicks = XlHlp.DisplayInWatchWindow("Begin");

            System.Collections.IEnumerable history =
                            versionControlServer.QueryHistory(
                                teamProject.ServerItem,
                                LatestVersionSpec.Instance,
                                0,
                                RecursionType.Full,
                                null,
                                new DateVersionSpec(DateTime.Now - TimeSpan.FromDays(goBackDays)),
                                LatestVersionSpec.Instance,
                                Int32.MaxValue,
                                false,
                                false);

            // Go find everyone that has issued changes.

            foreach (Changeset changeset in history)
            {
                if (developers.ContainsKey(changeset.Owner))
                {
                    developers[changeset.Owner] += 1;

                    if (developersEarliestDate[changeset.Owner] > changeset.CreationDate)
                    {
                        developersEarliestDate[changeset.Owner] = changeset.CreationDate;
                    }

                    if (developersLatestDate[changeset.Owner] < changeset.CreationDate)
                    {
                        developersLatestDate[changeset.Owner] = changeset.CreationDate;
                    }
                }
                else
                {
                    developers.Add(changeset.Owner, 1);
                    developersEarliestDate.Add(changeset.Owner, changeset.CreationDate);
                    developersLatestDate.Add(changeset.Owner, changeset.CreationDate);
                }
            }

            XlHlp.DisplayInWatchWindow("End", startTicks);
        }

        #endregion
    }
}
