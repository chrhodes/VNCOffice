using Microsoft.Office.Interop.Visio;

namespace VNC.Visio.VSTOAddIn.Domain
{
    public class ConnectionPointRow
    {
        public string Name { get; set; }
        public string X { get; set; }
        public string Y { get; set; }
        public string DirX { get; set; }
        public string A { get; set; }
        public string DirY { get; set; }
        public string B { get; set; }
        public string Type { get; set; }
        public string C { get; set; }
        public string D { get; set; }

        public static ConnectionPointRow Get_ConnectionPointRow(Shape shape)
        {
            ConnectionPointRow row = new ConnectionPointRow();

            Section section = shape.Section[(short)VisSectionIndices.visSectionConnectionPts];
            Row sectionRow = section[0];

            // TODO(crhodes)
            // Handle multiple rows
            row.Name = sectionRow.Name;

            row.X = sectionRow[VisCellIndices.visCnnctX].FormulaU;

            return row;
        }
    }
}
