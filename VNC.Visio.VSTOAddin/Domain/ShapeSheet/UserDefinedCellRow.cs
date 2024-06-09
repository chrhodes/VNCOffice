using System.Collections.ObjectModel;

using Microsoft.Office.Interop.Visio;

namespace VNC.Visio.VSTOAddIn.Domain
{
    public class UserDefinedCellRow
    {
        public string Name { get; set; }
        public string Value { get; set; }
        public string Prompt { get; set; }

        public static ObservableCollection<UserDefinedCellRow> GetRows(Shape shape)
        {
            var rows = new ObservableCollection<UserDefinedCellRow>();

            Section section = shape.Section[(short)VisSectionIndices.visSectionUser];

            var rowCount = section.Count;

            for (short i = 0; i < rowCount; i++)
            {
                var userRow = new UserDefinedCellRow();

                var row = section[i];

                userRow.Name = row.NameU;

                userRow.Value = row[(short)VisCellIndices.visUserValue].FormulaU;
                userRow.Prompt = row[(short)VisCellIndices.visUserPrompt].FormulaU;

                rows.Add(userRow);
            }

            return rows;
        }
    }
}
