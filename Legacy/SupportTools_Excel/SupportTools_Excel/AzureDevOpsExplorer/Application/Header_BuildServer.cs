using XlHlp = VNC.AddinHelper.Excel;

namespace SupportTools_Excel.AzureDevOpsExplorer.Application
{
    public class Header_BuildServer
    {
        #region Build Server (BS)

        internal static void Add_BuildAgents(XlHlp.XlLocation insertAt)
        {
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Name");
            //XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 30, "Description");
            //XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 10, "QueueStatus");

            insertAt.IncrementRows();
        }

        internal static void Add_BuildControllers(XlHlp.XlLocation insertAt)
        {
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Name");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 30, "Description");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 10, "Enabled");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 10, "AgentsCount");

            insertAt.IncrementRows();
        }

        internal static void  Add_BuildDefinitions(XlHlp.XlLocation insertAt)
        {
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Team Project");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Name");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 30, "Description");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 10, "QueueStatus");

            insertAt.IncrementRows();
        }

        internal static void Add_BuildProcessTemplates(XlHlp.XlLocation insertAt)
        {
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Id");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Description");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "TemplateType");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Version");

            insertAt.IncrementRows();
        }

        internal static void Add_Builds(XlHlp.XlLocation insertAt)
        {
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "BuildController");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "LabelName");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 10, "StartTime");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 10, "FinishTime");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 30, "Finished");

            insertAt.IncrementRows();
        }

        internal static void Add_BuildServiceHosts(XlHlp.XlLocation insertAt)
        {
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Name");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Status");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "StatusChangedOn");

            insertAt.IncrementRows();
        }

        #endregion
    }
}
