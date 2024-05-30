
using SupportTools_Excel.AzureDevOpsExplorer.Domain;

using VNC;
using VNC.AddinHelper;

using XlHlp = VNC.AddinHelper.Excel;

namespace SupportTools_Excel.AzureDevOpsExplorer.Application
{
    class Header_WorkItemStore
    {
        #region WorkItem Store (WIS)

        internal static void Add_TP_Areas(XlHlp.XlLocation insertAt)
        {
            long startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Team Project");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Area");

            // TODO(crhodes)
            // This can have a variable number of columns.  Not sure how to label them.

            insertAt.IncrementRows();

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);
        }


        internal static void Add_TP_Iterations(XlHlp.XlLocation insertAt)
        {
            long startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Team Project");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Iteration");

            // TODO(crhodes)
            // This can have a variable number of columns.  Not sure how to label them.

            insertAt.IncrementRows();

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);
        }

        internal static void Add_TP_FieldMapping(XlHlp.XlLocation insertAt)
        {
            long startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Team Project");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "WIT Name");

            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 55, "Field.Name");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 55, "Field.RefName");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 15, "Field.Type");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 15, "Field.Required");

            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "MapType");

            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Control.Label");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Control.Name");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 55, "Control.FieldName");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 25, "Control.Type");

            insertAt.IncrementRows();

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);
        }

        internal static void Add_TP_WorkItemFields(XlHlp.XlLocation insertAt)
        {
            long startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Team Project");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "WIT Name");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Name");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "FieldType");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "SystemType");

            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 6, "Id");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 12, "IsComputed");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 12, "IsCoreField");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 12, "IsEditable");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 12, "IsIdentity");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 12, "IsIndexed");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 12, "IsQueryable");

            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 12, "ReferenceName");

            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 30, "ReportingAttributes.Name");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 40, "ReportingAttributes.ReferenceName");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Usage");

            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 40, "FieldNameType");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 60, "AllowedValues");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 60, "ProhibitedValues");

            insertAt.IncrementRows();

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);
        }

        internal static void Add_TP_WorkItemActivity(XlHlp.XlLocation insertAt)
        {
            long startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Team Project");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Name");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Count");
            //XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "FieldCount");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "LastCreateDate");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "LastChangedDate");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "LastRevisedDate");

            // TODO(crhodes)
            // Since we now can pass in a CellFormatSpecification, might be able to go back to just using insertAt.AddOffsetColumn
            //insertAt.AddOffsetColumnX();

            //CellFormatSpecification lucidia7 = insertAt.CreateCellFormat("lucidia7", fontSize: 7);
            //lucidia7.Font.Name = "Lucida Sans Typewriter";

            //XlHlp.AddColumnHeaderToSheetX(insertAt.workSheet, insertAt.RowCurrent, insertAt.ColumnOffset,
            //    180, "Transitions", lucidia7);

            insertAt.IncrementRows();

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);
        }

        internal static void Add_TP_WorkItemTypes(XlHlp.XlLocation insertAt)
        {
            long startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Team Project");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Name");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Count");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "FieldCount");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "LastCreateDate");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "LastChangedDate");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "LastRevisedDate");

            // TODO(crhodes)
            // Since we now can pass in a CellFormatSpecification, might be able to go back to just using insertAt.AddOffsetColumn
            insertAt.AddOffsetColumnX();

            CellFormatSpecification lucidia7 = insertAt.CreateCellFormat("lucidia7", fontSize: 7);
            lucidia7.Font.Name = "Lucida Sans Typewriter";

            XlHlp.AddColumnHeaderToSheetX(insertAt.workSheet, insertAt.RowCurrent, insertAt.ColumnOffset,
                180, "Transitions", lucidia7);

            insertAt.IncrementRows();

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);
        }

        internal static void Add_TP_WorkItemDetails(XlHlp.XlLocation insertAt, Options_AZDO_TFS options)
        {
            long startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Project");

            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Id");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Type");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Title");

            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "CreatedBy");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "CreatedDate");

            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "ChangedBy");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "ChangedDate");

            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "AuthorizedDate");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "RevisedDate");

            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Rev");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Revision");

            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "State");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Reason");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Tags");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "AreaPath");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "IterationPath");

            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "RelatedLinkCount");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "ExternalLinkCount");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "HyperLinkCount");

            // NOTE(crhodes)
            // The Query can specify additional fields to display
            // They are added to the query to avoid a round trip if accessed 
            // after the result set is returned.

            // Add Headers for any requested

            if ((options.WorkItemQuerySpec.Fields?.Count ?? 0) > 0)
            {
                foreach (string field in options.WorkItemQuerySpec.Fields)
                {
                    XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, field);
                }
            }

            insertAt.IncrementRows();

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);
        }

        internal static void Add_TP_WorkItemFieldValues(XlHlp.XlLocation insertAt)
        {
            long startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 10, "Id");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 45, "Name");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 45, "ReferenceName");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 30, "OriginalValue");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 30, "Value");

            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 30, "FieldType");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 30, "SystemType");

            insertAt.IncrementRows();

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);
        }

        internal static void Add_TP_WorkItemFieldValues2(XlHlp.XlLocation insertAt)
        {
            long startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "Team Project");

            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 10, "WI Id");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 10, "WI Type");

            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 10, "Field Id");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 45, "Name");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 45, "ReferenceName");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 30, "OriginalValue");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 30, "Value");

            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 30, "FieldType");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 30, "SystemType");

            insertAt.IncrementRows();

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);
        }

        internal static void Add_TP_WorkItem_Links(XlHlp.XlLocation insertAt)
        {
            long startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 10, "ArtifactLinkType");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 10, "BaseType");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 10, "Comment");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 10, "IsLocked");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 10, "IsNew");

            insertAt.IncrementRows();

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);
        }

        internal static void Add_TP_WorkItem_WorkItemLinks(XlHlp.XlLocation insertAt)
        {
            long startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 15, "Source.Type");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 12, "Source.Id");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 15, "Source.State");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 30, "Source.Title");

            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 15, "Target.Type");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 12, "Target.Id");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 15, "Target.State");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 30, "Target.Title");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 15, "Target.Created");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 15, "Target.CreatedBy");

            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 15, "Link Source.Id");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 15, "Link Target.Id");

            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "AddedBy");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 15, "AddedDate");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 18, "ArtifactLinkType");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 12, "BaseType");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 15, "ChangedDate");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 15, "Comment");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 10, "IsLocked");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 10, "IsNew");

            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 10, "LinkTypeEnd.Id");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "LinkTypeEnd.ImmutableName");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "LinkTypeEnd.IsForwardLink");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 60, "LinkTypeEnd.LinkType");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "LinkTypeEnd.Name");

            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 10, "LinkTypeEnd.OppositeEnd.Id");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "LinkTypeEnd.OppostieEnd.ImmutableName");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "LinkTypeEnd.OppostieEnd.IsForwardLink");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 60, "LinkTypeEnd.OppostieEnd.LinkType");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 20, "LinkTypeEnd.OppositeEnd.Name");

            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 10, "RemovedBy");
            XlHlp.AddColumnHeaderToSheet(insertAt.AddOffsetColumn(), 15, "RemovedDate");

            insertAt.IncrementRows();

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion
    }
}
