using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Windows;

using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.Framework.Client;
using Microsoft.TeamFoundation.Framework.Common;
using Microsoft.TeamFoundation.Server;
using Microsoft.TeamFoundation.VersionControl.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;

using SupportTools_Excel.AzureDevOpsExplorer.Domain;

using SupportTools_Excel.Domain;

using VNC;

using XlHlp = VNC.AddinHelper.Excel;

namespace SupportTools_Excel.AzureDevOpsExplorer.Application
{
    class Body_VersionControlServer
    {
        #region Version Control Server (VCS)

        internal static void Add_TP_Changesets(
            XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options,
            VersionControlServer versionControlServer,
            TeamProject teamProject)
        {
            Int64 startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            //TeamProject teamProject = VNC.TFS.Helper.Get_TeamProject(versionControlServer, teamProjectName);

            var path = teamProject.ServerItem;

            //var queryHistory = Server.VersionControlServer.QueryHistory(
            //    teamProject.ServerItem,
            //    VersionSpec.Latest,
            //    0,
            //    RecursionType.Full,
            //    null,
            //    VersionSpec.Latest,
            //    VersionSpec.Latest,
            //    Int32.MaxValue,
            //    true,
            //    true,
            //    false,
            //    false);

            var queryHistory = versionControlServer.QueryHistory(
                teamProject.ServerItem,
                VersionSpec.Latest,
                0,
                RecursionType.Full,
                null,
                null,
                null,
                Int32.MaxValue,
                true,
                true,
                false,
                false);

            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), teamProject.Name);

            try
            {
                Changeset lastestChangeset = queryHistory.Cast<Changeset>().First();

                string lastChangesetId = lastestChangeset.ChangesetId.ToString();
                string lastChangeSetCreationDate = lastestChangeset.CreationDate.ToString();

                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), lastChangesetId);
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), lastChangeSetCreationDate);
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), teamProject.VersionControlServer.SupportedFeatures.ToString());
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), teamProject.VersionControlServer.WebServiceLevel.ToString());
            }
            catch (InvalidOperationException ioe)
            {
                if (ioe.Message.Equals("Sequence contains no elements"))
                {
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), "No Changesets");
                }
                else
                {
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), ioe.ToString());
                }

            }
            catch (Exception ex)
            {
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), ex.ToString());
            }

            insertAt.IncrementRows();

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);
        }

        internal static void Add_TP_Changesets(
            XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options,
            ICommonStructureService commonStructureService,
            bool listChanges, bool listWorkItems, IEnumerable history)
        {
            Int64 startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            foreach (Changeset changeset in history)
            {
                insertAt.ClearOffsets();

                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), changeset.ChangesetId.ToString());
                //XlHlp.AddContentToCell(insertAt.AddOffsetColumn(), changeset.CheckinNote.ToString());
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), changeset.Committer);
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), changeset.CommitterDisplayName);
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), changeset.Owner);
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), changeset.OwnerDisplayName);
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), changeset.CreationDate.ToString());
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), changeset.CheckinNote.ToString());
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), changeset.Comment);
                //XlHlp.AddContentToCell(insertAt.AddOffsetColumn(), changeset.AssociatedWorkItems.Count().ToString());
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), changeset.Changes.Count().ToString());
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), changeset.WorkItems.Count().ToString());
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), changeset.AssociatedWorkItems.Count().ToString());

                insertAt.IncrementRows();

                if (listChanges)
                {
                    insertAt.IncrementColumns();

                    foreach (Change change in changeset.Changes)
                    {
                        try
                        {
                            XlHlp.AddContentToCell(insertAt.AddRowX(1), Section_VersionControlServer.GetChangeInfo(change));
                            //XlHlp.AddContentToCell(insertAt.AddRow(), GetIterationInfo(workItem));
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.ToString());
                        }
                    }

                    insertAt.DecrementColumns();
                }

                if (listWorkItems)
                {
                    insertAt.IncrementColumns();

                    foreach (WorkItem workItem in changeset.WorkItems)
                    {
                        try
                        {
                            XlHlp.AddContentToCell(insertAt.AddRowX(1), Section_VersionControlServer.GetWorkItemInfo(workItem));
                            XlHlp.AddContentToCell(insertAt.AddRowX(1), Section_VersionControlServer.GetIterationInfo(workItem, commonStructureService));
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.ToString());
                        }
                    }

                    insertAt.DecrementColumns();
                }
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);
        }

        internal static void Add_TP_Developers(
            XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options,
            string teamProjectName,
            SortedDictionary<string, int> developers,
            SortedDictionary<string, DateTime> developersLatestDate,
            SortedDictionary<string, DateTime> developersEarliestDate)
        {
            Int64 startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            foreach (string developer in developers.Keys)
            {
                insertAt.ClearOffsets();

                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), teamProjectName);
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), developer);
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), developers[developer].ToString());
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), developersEarliestDate[developer].ToString());
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), developersLatestDate[developer].ToString());

                insertAt.IncrementRows();
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);
        }

        internal static void Add_TP_Shelvesets(
            XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options,
            Shelveset[] shelvesets)
        {
            Int64 startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            try
            {
                foreach (Shelveset item in shelvesets)
                {
                    insertAt.ClearOffsets();

                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), item.OwnerDisplayName);
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), item.OwnerName);
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), item.Name);
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), item.CreationDate.ToString());
                    //ExcelHlp.AddContentToCell(insertAt.AddOffsetColumn(), item.DisplayName);
                    //ExcelHlp.AddContentToCell(insertAt.AddOffsetColumn(), item.QualifiedName);
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), item.CheckinNote.ToString());
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), item.Comment);

                    insertAt.IncrementRows();
                }
            }
            catch (Exception ex)
            {
                string msg = string.Format("{0} - {1}", "TP", ex.ToString());

                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), msg);
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);
        }

        internal static void Add_TP_Teams(
            XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options,
            VersionControlServer versionControlServer,
            TeamProject teamProject,
            IEnumerable<TeamFoundationTeam> allTeams,
            TeamFoundationTeam defaultTeam)
        {
            Int64 startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            foreach (var team in allTeams.OrderBy(team => team.Name))
            {
                insertAt.ClearOffsets();

                TeamFoundationIdentity[] teamMembers = team.GetMembers(versionControlServer.TeamProjectCollection, MembershipQuery.Expanded);

                foreach (var member in teamMembers.OrderBy(m => m.UniqueName))
                {
                    insertAt.ClearOffsets();

                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), teamProject.Name);

                    // Team 

                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), team.Name);
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), team.Description);
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(),
                        defaultTeam.Name.Equals(team.Name) ? "*" : "");

                    // Members

                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), member.DisplayName);
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), member.UniqueName);

                    insertAt.IncrementRows();
                }
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);
        }

        internal static void Add_TP_Workspaces(
            XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options,
            Workspace[] workSpaces)
        {
            Int64 startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            foreach (Workspace workspace in workSpaces)
            {
                insertAt.ClearOffsets();

                // Keep in same order with headers, supra.

                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), workspace.Computer);
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), workspace.Name);
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), workspace.OwnerDisplayName);
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), workspace.OwnerName);
                //ExcelHlp.AddContentToCell(rngOutput.Offset[currentRow, col++], workspace.DisambiguatedDisplayName);
                //ExcelHlp.AddContentToCell(rngOutput.Offset[currentRow, col++], workspace.DisplayName);
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), workspace.LastAccessDate.ToString());
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), workspace.Comment);
                //ExcelHlp.AddContentToCell(rngOutput.Offset[XlLocation.Rows, col++], workspace.QualifiedName);

                insertAt.IncrementRows();
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion
    }
}
