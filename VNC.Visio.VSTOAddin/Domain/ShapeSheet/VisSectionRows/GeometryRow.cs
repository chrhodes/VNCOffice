using Microsoft.Office.Interop.Visio;

namespace VNC.Visio.VSTOAddIn.Domain
{
    public class GeometryRow
    {
        // TODO(crhodes)
        // This is going to take work
        public string Name { get; set; }
        public string X { get; set; }
        public string Y { get; set; }
        public string A { get; set; }
        public string B { get; set; }
        public string C { get; set; }
        public string D { get; set; }

        public static GeometryRow GetRow(Shape shape)
        {
            GeometryRow row = new GeometryRow();

            // TODO(crhodes)
            // Can't find Section Index

            //Visio.Section section = shape.Section[(short)Visio.VisSectionIndices.v];
            //Visio.Row sectionRow = section[0];

            //// TODO(crhodes)
            //// Handle multiple rows

            //row.Action = sectionRow[VisCellIndices.visActionAction].FormulaU;

            return row;
        }
    }
}
