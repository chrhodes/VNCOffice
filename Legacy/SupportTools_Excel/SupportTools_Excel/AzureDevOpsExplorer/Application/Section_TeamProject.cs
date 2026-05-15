using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;

using Microsoft.Office.Interop.Excel;
using Microsoft.TeamFoundation.Framework.Client;
using Microsoft.TeamFoundation.Framework.Common;
using Microsoft.TeamFoundation.VersionControl.Client;

using VNC;

using XlHlp = VNC.AddinHelper.Excel;

namespace SupportTools_Excel.AzureDevOpsExplorer.Application
{
    class Section_TeamProject
    {
        #region Team Project (TP)

        private static List<TeamFoundationIdentity> _TeamProject_Groups =
            new List<TeamFoundationIdentity>();

        // For Team Project
        private static Dictionary<IdentityDescriptor, TeamFoundationIdentity> _TeamProject_Identities =
            new Dictionary<IdentityDescriptor, TeamFoundationIdentity>(IdentityDescriptorComparer.Instance);

        internal static XlHlp.XlLocation AddSections(
            XlHlp.XlLocation insertAt,
            TeamProject teamProject,
            List<string> sectionsToDisplay)
        {
            long startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            if (sectionsToDisplay.Count != 0)
            {

                if (insertAt.OrientVertical)
                {
                    //XlHlp.AddContentToCell(insertAt.AddRowX(), "TeamProject (TP) Information");
                    XlHlp.AddSectionInfo(insertAt.AddRow(), "TeamProject (TP) Information", "");
                    insertAt.IncrementRows();
                }
                else
                {
                    //XlHlp.AddContentToCell(insertAt.AddRowX(), "TeamProject (TP) Information");
                    XlHlp.AddSectionInfo(insertAt.AddRow(), "TeamProject(TP) Information", "",
                        orientation: XlOrientation.xlUpward);
                    insertAt.DecrementRows();   // AddRow bumped it.
                    insertAt.IncrementColumns();
                }

                if (sectionsToDisplay.Contains("Info"))
                {
                    insertAt = Add_Info(insertAt, teamProject).IncrementPosition(insertAt.OrientVertical);
                }

                if (sectionsToDisplay.Contains("Members"))
                {
                    insertAt = Add_Members(insertAt, teamProject).IncrementPosition(insertAt.OrientVertical);
                }
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);

            return insertAt;
        }

        internal static XlHlp.XlLocation Add_Info(
            XlHlp.XlLocation insertAt,
            TeamProject teamProject)
        {
            long startTicks = XlHlp.DisplayInWatchWindow(insertAt);

            insertAt.MarkStart();

            XlHlp.AddLabeledInfo(insertAt.AddRow(2), "TP Name", teamProject.Name);
            XlHlp.AddLabeledInfo(insertAt.AddRow(2), "AbsoluteUri", teamProject.ArtifactUri.AbsoluteUri);
            XlHlp.AddLabeledInfo(insertAt.AddRow(2), "ServerItem", teamProject.ServerItem);
            XlHlp.AddLabeledInfo(insertAt.AddRow(2), "VCS ServerGuid", teamProject.VersionControlServer.ServerGuid.ToString());

            // TODO(crhodes)
            // What else can we get here?
            // Capabilities
            // Template?
            // Creation Date?

            insertAt.MarkEnd();

            if (!insertAt.OrientVertical)
            {
                // Skip past the info just added.
                insertAt.SetLocation(insertAt.RowStart, insertAt.MarkEndColumn + 1);
            }

            XlHlp.DisplayInWatchWindow(insertAt, startTicks, "End");

            return insertAt;
        }

        internal static XlHlp.XlLocation Add_Members(
            XlHlp.XlLocation insertAt,
            TeamProject teamProject)
        {
            long startTicks = XlHlp.DisplayInWatchWindow(insertAt);

            try
            {
                int currentRows = insertAt.RowsAdded;

                // Save the location of the count so we can update later after have traversed all items.

                Range rngTitle = insertAt.GetCurrentRange();

                if (insertAt.OrientVertical)
                {
                    XlHlp.AddSectionInfo(insertAt.AddRow(),
                        "Members Group", "");
                }
                else
                {
                    XlHlp.AddSectionInfo(insertAt.AddRow(),
                        "Members Group", "",
                        orientation: XlOrientation.xlUpward);
                    insertAt.IncrementColumns();
                }

                TeamFoundationIdentity[] projectGroups = AzureDevOpsExplorer.Presentation.Views.Server.IdentityManagementService.ListApplicationGroups(
                    null, ReadIdentityOptions.None);

                //TeamFoundationIdentity[] projectGroups = AzureDevOpsExplorer.Presentation.Views.Server.IdentityManagementService.ListApplicationGroups(
                //    teamProject.ArtifactUri.AbsoluteUri, ReadIdentityOptions.None);

                Dictionary<IdentityDescriptor, object> descriptorSet = new Dictionary<IdentityDescriptor, object>(IdentityDescriptorComparer.Instance);

                foreach (TeamFoundationIdentity projectGroup in projectGroups)
                {
                    descriptorSet[projectGroup.Descriptor] = projectGroup.Descriptor;
                }

                // Expanded membership of project groups
                projectGroups = AzureDevOpsExplorer.Presentation.Views.Server.IdentityManagementService.ReadIdentities(descriptorSet.Keys.ToArray(), MembershipQuery.Expanded, ReadIdentityOptions.None);

                // Collect all descriptors
                foreach (TeamFoundationIdentity projectGroup in projectGroups)
                {
                    foreach (IdentityDescriptor mem in projectGroup.Members)
                    {
                        descriptorSet[mem] = mem;
                    }
                }

                // NOTE(crhodes)
                // Might need to ensure that _Global_Groups and _Global_Identities already populated.

                if (Section_TeamProjectCollection._Global_Identities.Count == 0)
                {
                    TeamFoundationIdentity everyoneExpanded = AzureDevOpsExplorer.Presentation.Views.Server.IdentityManagementService.ReadIdentity(
                        GroupWellKnownDescriptors.EveryoneGroup, MembershipQuery.Expanded, ReadIdentityOptions.None);
                    AZDOHelper.FetchIdentities(everyoneExpanded.Members, Section_TeamProjectCollection._Global_Groups, Section_TeamProjectCollection._Global_Identities);
                }

                _TeamProject_Groups.Clear();
                _TeamProject_Identities.Clear();

                AZDOHelper.FetchIdentities(descriptorSet.Keys.ToArray(), _TeamProject_Groups, _TeamProject_Identities);

                insertAt.MarkStart(XlHlp.MarkType.GroupTable);
                // Keep in same order as fields, infra.

                // Group

                XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 50, "Identifier");
                XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 50, "Identity");

                // Members

                XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 15, "IsContainer");
                XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 50, "DisplayName");
                XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 80, "UniqueName");
                XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 40, "IdentityType");
                XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "UniqueUserId");
                XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 10, "IsActive");

                insertAt.IncrementRows();

                foreach (TeamFoundationIdentity identity in _TeamProject_Groups)
                {
                    foreach (IdentityDescriptor member in identity.Members)
                    {
                        insertAt.ClearOffsets();

                        try
                        {
                            // Group

                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), identity.Descriptor.Identifier);
                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), identity.DisplayName);

                            // Members

                            // NOTE(crhodes)
                            // This line is throwing exception.  Why?

                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), Section_TeamProjectCollection._Global_Identities[member].IsContainer.ToString());
                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), Section_TeamProjectCollection._Global_Identities[member].DisplayName);
                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), Section_TeamProjectCollection._Global_Identities[member].UniqueName);
                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), Section_TeamProjectCollection._Global_Identities[member].Descriptor.IdentityType);
                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), Section_TeamProjectCollection._Global_Identities[member].UniqueUserId.ToString());
                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), Section_TeamProjectCollection._Global_Identities[member].IsActive.ToString());

                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.ToString());
                        }

                        insertAt.IncrementRows();
                    }
                }

                insertAt.MarkEnd(XlHlp.MarkType.GroupTable, string.Format("tblTPMembers_{0}", insertAt.workSheet.Name));

                insertAt.Group(insertAt.OrientVertical);

                // Update counts.  -2 covers Header and Table Column Header

                if (insertAt.OrientVertical)
                {
                    XlHlp.AddLabeledInfoX(rngTitle, "Members Group", (insertAt.RowsAdded - currentRows - 2).ToString());
                }
                else
                {
                    XlHlp.AddLabeledInfoX(rngTitle, "Members Group", (insertAt.RowsAdded - currentRows - 2).ToString(), orientation: XlOrientation.xlUpward);
                }

                insertAt.EndSectionAndSetNextLocation(insertAt.OrientVertical);

                //insertAt.AddRow();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            XlHlp.DisplayInWatchWindow(insertAt, startTicks, "End");

            return insertAt;
        }

        #endregion

    }
}
