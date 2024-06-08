using Microsoft.Office.Interop.Visio;
using System.Collections.ObjectModel;

namespace VNC.Visio.VSTOAddIn.Domain
{
    public class ActionTagRow
    {
        public string Name { get; set; }

        public string X { get; set; }
        public string Y { get; set; }
        public string TagName { get; set; }
        public string XJustify { get; set; }
        public string YJustify { get; set; }
        public string DisplayMode { get; set; }
        public string ButtonFace { get; set; }
        public string Description { get; set; }
        public string Disabled { get; set; }

        public static ActionTagRow GetRow(Shape shape)
        {
            ActionTagRow row = new ActionTagRow();

            // TODO(crhodes)
            // Can't find Section Index

            //Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.];
            //Visio.Row sectionRow = section[0];

            //// TODO(crhodes)
            //// Handle multiple rows

            //row.Action = sectionRow[VisCellIndices.visActionAction].FormulaU;

            return row;
        }

        public static ObservableCollection<ActionTagRow> GetRows(Shape shape)
        {
            var rows = new ObservableCollection<ActionTagRow>();

            Section section = shape.Section[(short)VisSectionIndices.visSectionSmartTag];

            var rowCount = section.Count;

            for (short i = 0; i < rowCount; i++)
            {
                ActionTagRow actionTagRow = new ActionTagRow();

                var row = section[i];

                actionTagRow.Name = row.NameU;

                actionTagRow.X = row[(short)VisCellIndices.visSmartTagX].FormulaU;
                actionTagRow.Y = row[(short)VisCellIndices.visSmartTagY].FormulaU;
                actionTagRow.TagName = row[(short)VisCellIndices.visSmartTagName].FormulaU;
                actionTagRow.XJustify = row[(short)VisCellIndices.visSmartTagXJustify].FormulaU;
                actionTagRow.YJustify = row[(short)VisCellIndices.visSmartTagYJustify].FormulaU;
                actionTagRow.DisplayMode = row[(short)VisCellIndices.visSmartTagDisplayMode].FormulaU;
                actionTagRow.ButtonFace = row[(short)VisCellIndices.visSmartTagButtonFace].FormulaU;
                actionTagRow.Description = row[(short)VisCellIndices.visSmartTagDescription].FormulaU;
                actionTagRow.Disabled = row[(short)VisCellIndices.visSmartTagDisabled].FormulaU;

                rows.Add(actionTagRow);
            }

            return rows;
        }
    }
}
