
using XlHlp = VNC.AddinHelper.Excel;

namespace SupportTools_Excel.AzureDevOpsExplorer.Application
{
    class Header_VersionControlServer
    {
        #region Version Control Sever (VCS)

        internal static void Add_Changesets(XlHlp.XlLocation insertAt)
        {
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Team Project");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Changeset ID");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Creation Date");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Supported Features");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "WebService Level");

            insertAt.IncrementRows();
        }


        internal static void Add_Shelvesets(XlHlp.XlLocation insertAt)
        {
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "OwnerDisplayName");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "OwnerName");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Name");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "CreationDate");
            //ExcelHlp.AddColumnHeaderToheet(insertAt.AddOffsetCoumn(), 20, "DisplayName");
            //ExcelHlp.AddColumnHeaderToheet(insertAt.AddOffsetCoumn(), 20, "QualifiedName");

            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "CheckinNote");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Comment");

            insertAt.IncrementRows();
        }

        internal static void Add_TP_Changesets(XlHlp.XlLocation insertAt)
        {
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "ChangesetId");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Committer");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Committer DisplayName");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Owner");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Owner DisplayName");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "CreationDate");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 50, "CheckinNote");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 30, "Comment");
            //ExcelHlp.AddColumnHeaderToheet(insertAt.AddOffsetCoumn(),ciatedWorkItems");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Changes");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "WorkItems");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Associated WorkItems");

            insertAt.IncrementRows();
        }

        internal static void Add_TP_Developers(XlHlp.XlLocation insertAt)
        {
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "TFS Team Project");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Developer");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Changeset Count");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Earliest Date");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Latest Date");

            insertAt.IncrementRows();
        }

        internal static void Add_TP_Teams(XlHlp.XlLocation insertAt)
        {
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "TeamProject");

            // Team

            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Team");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Description");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 5, "DefaultTeam");

            // Members

            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "DisplayName");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "UniqueName");

            insertAt.IncrementRows();
        }

        internal static void Add_TP_Workspaces(XlHlp.XlLocation insertAt)
        {
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Computer");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Name");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "OwnerDisplayName");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "OwnerName");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "LastAccessDate");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Comment");

            insertAt.IncrementRows();
        }

        #endregion
    }
}
