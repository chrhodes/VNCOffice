using System.Collections.ObjectModel;

using Microsoft.Office.Interop.Visio;

namespace VNC.Visio.VSTOAddIn.Domain
{
    public class ActionRow
    {
        public string Name { get; set; }

        public string Action { get; set; }
        public string Menu { get; set; }
        public string TagName { get; set; }
        public string ButtonFace { get; set; }
        public string SortKey { get; set; }
        public string Checked { get; set; }
        public string Disabled { get; set; }
        public string ReadOnly { get; set; }
        public string Invisible { get; set; }
        public string BeginGroup { get; set; }
        public string FlyoutChild { get; set; }

        public static ActionRow GetRow(Shape shape)
        {
            ActionRow row = new ActionRow();

            Section section = shape.Section[(short)VisSectionIndices.visSectionAction];
            Row sectionRow = section[0];


            row.Action = sectionRow[VisCellIndices.visActionAction].FormulaU;

            return row;
        }

        public static ObservableCollection<ActionRow> GetRows(Shape shape)
        {
            var rows = new ObservableCollection<ActionRow>();

            Section section = shape.Section[(short)VisSectionIndices.visSectionAction];

            var rowCount = section.Count;

            for (short i = 0; i < rowCount; i++)
            {
                ActionRow actionRow = new ActionRow();

                var row = section[i];

                actionRow.Name = row.NameU;

                actionRow.Action = row[(short)VisCellIndices.visActionAction].FormulaU;
                actionRow.Menu = row[(short)VisCellIndices.visActionMenu].FormulaU;
                actionRow.TagName = row[(short)VisCellIndices.visActionTagName].FormulaU;
                actionRow.ButtonFace = row[(short)VisCellIndices.visActionButtonFace].FormulaU;
                actionRow.SortKey = row[(short)VisCellIndices.visActionSortKey].FormulaU;
                actionRow.Checked = row[(short)VisCellIndices.visActionChecked].FormulaU;
                actionRow.Disabled = row[(short)VisCellIndices.visActionDisabled].FormulaU;
                actionRow.ReadOnly = row[(short)VisCellIndices.visActionReadOnly].FormulaU;
                actionRow.Invisible = row[(short)VisCellIndices.visActionInvisible].FormulaU;
                actionRow.BeginGroup = row[(short)VisCellIndices.visActionBeginGroup].FormulaU;
                actionRow.FlyoutChild = row[(short)VisCellIndices.visActionFlyoutChild].FormulaU;

                rows.Add(actionRow);
            }

            return rows;
        }
    }
}
