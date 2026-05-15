
using Microsoft.TeamFoundation;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.Framework.Client;

using XlHlp = VNC.AddinHelper.Excel;
using VNC;

namespace SupportTools_Excel.AzureDevOpsExplorer.Application
{
    class Section_ConfigurationServer
    {
        internal static XlHlp.XlLocation AddSection_ConfigurationServer_Info(
            XlHlp.XlLocation insertAt,
            TfsConfigurationServer configurationServer)
        {
            insertAt.MarkStart(XlHlp.MarkType.None);

            XlHlp.AddLabeledInfo(insertAt.AddRow(2), "Name:", configurationServer.Name);
            XlHlp.AddLabeledInfo(insertAt.AddRow(2), "Culture:", configurationServer.Culture.DisplayName);
            XlHlp.AddLabeledInfo(insertAt.AddRow(2), "InstanceId:", configurationServer.InstanceId.ToString());
            XlHlp.AddLabeledInfo(insertAt.AddRow(2), "ServerCapabilities:", configurationServer.ServerCapabilities.ToString());
            XlHlp.AddLabeledInfo(insertAt.AddRow(2), "SessionId:", configurationServer.SessionId.ToString());
            XlHlp.AddLabeledInfo(insertAt.AddRow(2), "TimeZone:", configurationServer.TimeZone.ToString());
            XlHlp.AddLabeledInfo(insertAt.AddRow(2), "UICulture", configurationServer.UICulture.ToString());
            XlHlp.AddLabeledInfo(insertAt.AddRow(2), "Uri", configurationServer.Uri.ToString());

            XlHlp.AddLabeledInfo(insertAt.AddRow(2), "AuthorizedIdentity:", configurationServer.AuthorizedIdentity.DisplayName);
            XlHlp.AddLabeledInfo(insertAt.AddRow(2), "CatalogNode:", configurationServer.CatalogNode.FullPath);
            XlHlp.AddLabeledInfo(insertAt.AddRow(2), "HasAuthenticated:", configurationServer.HasAuthenticated.ToString());
            XlHlp.AddLabeledInfo(insertAt.AddRow(2), "IsHostedServer:", configurationServer.IsHostedServer.ToString());
            XlHlp.AddLabeledInfo(insertAt.AddRow(2), "ClientCacheDirectoryForInstance:", configurationServer.ClientCacheDirectoryForInstance);
            XlHlp.AddLabeledInfo(insertAt.AddRow(2), "ClientCacheDirectoryForUser:", configurationServer.ClientCacheDirectoryForUser);

            insertAt.MarkEnd(XlHlp.MarkType.None);

            if (!insertAt.OrientVertical)
            {
                // Skip past the info just added.
                insertAt.SetLocation(insertAt.RowStart, insertAt.MarkEndColumn + 1);
            }

            return insertAt;
        }

        internal static XlHlp.XlLocation AddSection_OperationalDatabaseNames(XlHlp.XlLocation insertAt)
        {
            long startTicks = Log.Trace("Enter", Common.LOG_CATEGORY);

            XlHlp.AddLabeledInfoX(insertAt.AddRowX(), "Operational Database Names:", "");

            insertAt.MarkStart(XlHlp.MarkType.GroupTable);

            XlHlp.AddLabeledInfoX(insertAt.AddRowX(2), "AnalysisCube:", $"{OperationalDatabaseNames.AnalysisCube}");
            XlHlp.AddLabeledInfoX(insertAt.AddRowX(2), "CoreServices:", $"{OperationalDatabaseNames.CoreServices}");
            XlHlp.AddLabeledInfoX(insertAt.AddRowX(2), "DeploymentRig:", $"{OperationalDatabaseNames.DeploymentRig}");
            XlHlp.AddLabeledInfoX(insertAt.AddRowX(2), "LabExecution:", $"{ OperationalDatabaseNames.LabExecution}");
            XlHlp.AddLabeledInfoX(insertAt.AddRowX(2), "TeamBuild:", $"{OperationalDatabaseNames.TeamBuild}");
            XlHlp.AddLabeledInfoX(insertAt.AddRowX(2), "TestRig:", $"{ OperationalDatabaseNames.TestRig}");
            XlHlp.AddLabeledInfoX(insertAt.AddRowX(2), "VersionControl:", $"{ OperationalDatabaseNames.VersionControl}");
            XlHlp.AddLabeledInfoX(insertAt.AddRowX(2), "Warehouse:", $"{ OperationalDatabaseNames.Warehouse}");
            XlHlp.AddLabeledInfoX(insertAt.AddRowX(2), "WorkItemTracking:", $"{OperationalDatabaseNames.WorkItemTracking}");
            XlHlp.AddLabeledInfoX(insertAt.AddRowX(2), "WorkItemTrackingAttachments:", $"{OperationalDatabaseNames.WorkItemTrackingAttachments}");

            Log.Trace("Exit", Common.LOG_CATEGORY, startTicks);

            return insertAt;
        }
    }
}
