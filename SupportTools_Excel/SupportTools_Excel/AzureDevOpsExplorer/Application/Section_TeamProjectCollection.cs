using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.Framework.Client;
using Microsoft.TeamFoundation.Framework.Common;

using SupportTools_Excel.AzureDevOpsExplorer.Domain;

using XlHlp = VNC.AddinHelper.Excel;

namespace SupportTools_Excel.AzureDevOpsExplorer.Application
{
    class Section_TeamProjectCollection
    {
        internal static List<TeamFoundationIdentity> _Global_Groups =
            new List<TeamFoundationIdentity>();

        // Global
        internal static Dictionary<IdentityDescriptor, TeamFoundationIdentity> _Global_Identities =
            new Dictionary<IdentityDescriptor, TeamFoundationIdentity>(IdentityDescriptorComparer.Instance);

        internal static XlHlp.XlLocation Add_Info(
            XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options,
            TfsTeamProjectCollection tpc, bool showDetails)
        {
            long startTicks = XlHlp.DisplayInWatchWindow(insertAt);

            XlHlp.AddLabeledInfo(insertAt.AddRow(), "Name:", tpc.Name);

            // HACK(crhodes)
            // Not sure why this is being called.  Passing null is throwing exception.
            //insertAt = Section_WorkItemStore.Add_Info(insertAt, options, null);

            XlHlp.DisplayInWatchWindow(insertAt, startTicks, "End");

            return insertAt;
        }
        
        internal static XlHlp.XlLocation Add_Members(
            XlHlp.XlLocation insertAt,
            Options_AZDO_TFS options)
        {
            long startTicks = XlHlp.DisplayInWatchWindow(insertAt);

            TeamFoundationIdentity everyone = AzureDevOpsExplorer.Presentation.Views.Server.IdentityManagementService.ReadIdentity(
                GroupWellKnownDescriptors.EveryoneGroup, MembershipQuery.Direct, ReadIdentityOptions.None);

            TeamFoundationIdentity licensees = AzureDevOpsExplorer.Presentation.Views.Server.IdentityManagementService.ReadIdentity(
                GroupWellKnownDescriptors.LicenseesGroup, MembershipQuery.Direct, ReadIdentityOptions.None);

            TeamFoundationIdentity namespaceAdministrators = AzureDevOpsExplorer.Presentation.Views.Server.IdentityManagementService.ReadIdentity(
                GroupWellKnownDescriptors.NamespaceAdministratorsGroup, MembershipQuery.Direct, ReadIdentityOptions.None);

            TeamFoundationIdentity serviceUsers = AzureDevOpsExplorer.Presentation.Views.Server.IdentityManagementService.ReadIdentity(
                GroupWellKnownDescriptors.ServiceUsersGroup, MembershipQuery.Direct, ReadIdentityOptions.None);

            if (everyone != null)
            {
                insertAt.ClearOffsets();

                XlHlp.AddLabeledInfo(insertAt.AddRow(), "Everyone", everyone.Members.Count().ToString());

                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), everyone.DisplayName);
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), everyone.UniqueName);
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), everyone.Descriptor.IdentityType);

                insertAt.IncrementRows();
            }
            else
            {
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "Everyone", "null");
            }

            if (licensees != null)
            {
                insertAt.ClearOffsets();

                XlHlp.AddLabeledInfo(insertAt.AddRow(), "Licensees", licensees.Members.Count().ToString());

                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), licensees.DisplayName);
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), licensees.UniqueName);
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), licensees.Descriptor.IdentityType);

                insertAt.IncrementRows();
            }
            else
            {
                XlHlp.AddLabeledInfoX(insertAt.AddRowX(), "Licensees", "null");
            }

            if (namespaceAdministrators != null)
            {
                insertAt.ClearOffsets();

                XlHlp.AddLabeledInfo(insertAt.AddRow(), "NamespaceAdministrators", namespaceAdministrators.Members.Count().ToString());

                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), namespaceAdministrators.DisplayName);
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), namespaceAdministrators.UniqueName);
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), namespaceAdministrators.Descriptor.IdentityType);

                insertAt.IncrementRows();
            }
            else
            {
                XlHlp.AddLabeledInfo(insertAt.AddRow(), "NamespaceAdministrators", "null");
            }

            if (serviceUsers != null)
            {
                insertAt.ClearOffsets();

                XlHlp.AddLabeledInfo(insertAt.AddRow(), "ServiceUsers", serviceUsers.Members.Count().ToString());

                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), serviceUsers.DisplayName);
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), serviceUsers.UniqueName);
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), serviceUsers.Descriptor.IdentityType);

                insertAt.IncrementRows();
            }
            else
            {
                XlHlp.AddLabeledInfoX(insertAt.AddRowX(), "ServiceUsers", "null");
            }

            TeamFoundationIdentity everyoneExpanded = AzureDevOpsExplorer.Presentation.Views.Server.IdentityManagementService.ReadIdentity(
                GroupWellKnownDescriptors.EveryoneGroup, MembershipQuery.Expanded, ReadIdentityOptions.None);

            TeamFoundationIdentity everyoneExpanded2 = AzureDevOpsExplorer.Presentation.Views.Server.IdentityManagementService.ReadIdentity(
                GroupWellKnownDescriptors.EveryoneGroup, MembershipQuery.Expanded, ReadIdentityOptions.IncludeReadFromSource);

            if (everyoneExpanded != null)
            {
                AZDOHelper.FetchIdentities(everyoneExpanded.Members, _Global_Groups, _Global_Identities);
            }

            TeamFoundationIdentity licenseesExpanded = AzureDevOpsExplorer.Presentation.Views.Server.IdentityManagementService.ReadIdentity(
                GroupWellKnownDescriptors.LicenseesGroup, MembershipQuery.Expanded, ReadIdentityOptions.None);

            if (licenseesExpanded != null)
            {
                AZDOHelper.FetchIdentities(licenseesExpanded.Members, _Global_Groups, _Global_Identities);
            }

            TeamFoundationIdentity serviceUsersExpanded = AzureDevOpsExplorer.Presentation.Views.Server.IdentityManagementService.ReadIdentity(
                GroupWellKnownDescriptors.ServiceUsersGroup, MembershipQuery.Expanded, ReadIdentityOptions.None);

            if (serviceUsersExpanded != null)
            {
                AZDOHelper.FetchIdentities(serviceUsersExpanded.Members, _Global_Groups, _Global_Identities);
            }

            TeamFoundationIdentity namespaceAdministratorsExpanded = AzureDevOpsExplorer.Presentation.Views.Server.IdentityManagementService.ReadIdentity(
                GroupWellKnownDescriptors.NamespaceAdministratorsGroup, MembershipQuery.Expanded, ReadIdentityOptions.None);

            if (namespaceAdministratorsExpanded != null)
            {
                AZDOHelper.FetchIdentities(namespaceAdministratorsExpanded.Members, _Global_Groups, _Global_Identities);
            }

            XlHlp.AddLabeledInfoX(insertAt.AddRowX(), "All Groups and Identities", "Lots");

            insertAt.MarkStart(XlHlp.MarkType.GroupTable);

            // Keep in same order as fields, infra.

            // Group

            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 50, "Top Level");

            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 50, "Group Identifier");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 50, "Group Identity");

            // Members

            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 15, "IsContainer");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 30, "TeamFoundationId");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 50, "DisplayName");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 80, "UniqueName");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 40, "IdentityType");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 40, "Identity");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "UniqueUserId");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 10, "IsActive");


            insertAt.IncrementRows();

            foreach (TeamFoundationIdentity identity in _Global_Groups)
            {
                Globals.ThisAddIn.Application.StatusBar = "Processing " + identity.DisplayName;

                foreach (IdentityDescriptor member in identity.Members)
                {
                    insertAt.ClearOffsets();

                    // Top Level

                    string topLevel = "";

                    MatchCollection matches = Regex.Matches(identity.DisplayName, @"\[.*\]");

                    if (matches.Count == 1)
                    {
                        topLevel = matches[0].Value;
                    }
                    else
                    {
                        topLevel = identity.DisplayName;

                    }

                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), topLevel);

                    // Group

                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), identity.Descriptor.Identifier);
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), identity.DisplayName);

                    // Members

                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), _Global_Identities[member].IsContainer.ToString());
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), _Global_Identities[member].TeamFoundationId.ToString());
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), _Global_Identities[member].DisplayName);
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), _Global_Identities[member].UniqueName);
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), _Global_Identities[member].Descriptor.IdentityType);
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), _Global_Identities[member].Descriptor.Identifier);
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), _Global_Identities[member].UniqueUserId.ToString());
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), _Global_Identities[member].IsActive.ToString());

                    insertAt.IncrementRows();
                }
            }

            insertAt.MarkEnd(XlHlp.MarkType.GroupTable, string.Format("tblMembers_{0}", insertAt.workSheet.Name));

            insertAt.Group(insertAt.OrientVertical);

            if (!insertAt.OrientVertical)
            {
                // Skip past the info just added.
                insertAt.SetLocation(insertAt.RowStart, insertAt.MarkEndColumn + 1);
            }

            XlHlp.DisplayInWatchWindow(insertAt, startTicks, "End");

            return insertAt;
        }

        internal static XlHlp.XlLocation AddSection_TeamProjects(
            XlHlp.XlLocation insertAt,
            ReadOnlyCollection<CatalogNode> teamProjects)
        {
            long startTicks = XlHlp.DisplayInWatchWindow(insertAt);

            XlHlp.AddLabeledInfo(insertAt.AddRow(), "Team Projects", teamProjects.Count().ToString());

            Worksheet ws = insertAt.workSheet;

            insertAt = DisplayListOf_TeamProjects(insertAt, teamProjects, displayDataOnly: false, string.Format("tblTP_{0}", ws.Name));

            if (!insertAt.OrientVertical)
            {
                // Skip past the info just added.
                insertAt.SetLocation(insertAt.RowStart, insertAt.TableEndColumn + 1);
            }

            return insertAt;
        }

        internal static XlHlp.XlLocation DisplayListOf_TeamProjects(XlHlp.XlLocation insertAt,
            ReadOnlyCollection<CatalogNode> projectNodes, bool displayDataOnly, string tableSuffix)
        {
            long startTicks = XlHlp.DisplayInWatchWindow(insertAt);

            if (!displayDataOnly)
            {
                insertAt.MarkStart(XlHlp.MarkType.GroupTable);

                //XlHlp.AddTitledInfo(insertAt.AddRow(), "Name", teamProjects.Count.ToString());
                //XlHlp.AddContentToCell(insertAt.AddOffsetColumn(), Name, 12, XlHlp.MakeBold.Yes);
                XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 25, "DisplayName");
                XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 35, "Description");
                XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 35, "Identifier");
                XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 35, "ProjectId");
                //XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 25, "ProjectName", 12);
                XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 25, "ProjectState");
                XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 62, "ProjectUri");
                XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 10, "Tfvc Enabled");
                XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 25, "SCC");



                //XlHlp.AddTitledInfo(insertAt.AddRow(2), "TP Name", teamProject.Name);
                //XlHlp.AddTitledInfo(insertAt.AddRow(2), "AbsoluteUri", teamProject.ArtifactUri.AbsoluteUri);
                //XlHlp.AddTitledInfo(insertAt.AddRow(2), "ServerItem", teamProject.ServerItem);
                //XlHlp.AddTitledInfo(insertAt.AddRow(2), "VCS ServerQuid", teamProject.VersionControlServer.ServerGuid.ToString());

                insertAt.IncrementRows();
            }
            // The columns in this method need to be kept in sync with CreateTeamProjectsInfo()

            foreach (CatalogNode projectNode in projectNodes.OrderBy(tp => tp.Resource.DisplayName))
            {
                insertAt.ClearOffsets();

                try
                {
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), projectNode.Resource.DisplayName);
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), projectNode.Resource.Description);
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), projectNode.Resource.Identifier.ToString());
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), projectNode.Resource.Properties["ProjectId"]);
                    //XlHlp.AddContentToCell(insertAt.AddOffsetColumn(), projectNode.Resource.Properties["ProjectName"]);
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), projectNode.Resource.Properties["ProjectState"]);
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), projectNode.Resource.Properties["ProjectUri"]);
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), projectNode.Resource.Properties["SourceControlTfvcEnabled"]);

                    string sccType = "??";

                    if (projectNode.Resource.Properties.Keys.Contains("SourceControlCapabilityFlags"))
                    {
                        switch (int.Parse(projectNode.Resource.Properties["SourceControlCapabilityFlags"]))
                        {
                            case 0:
                                sccType = "NONE";
                                break;

                            case 1:
                                sccType = "TFS";
                                break;

                            case 2:
                                sccType = "GIT";
                                break;

                            case 3:
                                sccType = "TFS/GIT";
                                break;

                            default:
                                break;

                        }
                    }

                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), sccType);
                }
                catch (Exception ex)
                {

                }

                //projectNode.FullPath
                //    projectNode.Resource.Description
                //    projectNode.Resource.Identifier


                insertAt.IncrementRows();
            }

            if (!displayDataOnly)
            {
                insertAt.MarkEnd(XlHlp.MarkType.GroupTable, string.Format("tblTP_{0}", tableSuffix));

                insertAt.Group(insertAt.OrientVertical, hide: true);
            }

            XlHlp.DisplayInWatchWindow(insertAt, startTicks, "End");

            return insertAt;
        }

    }
}
