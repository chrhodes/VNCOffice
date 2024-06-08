using Microsoft.Office.Interop.Visio;

namespace VNC.Visio.VSTOAddIn.Domain
{
    public class LineGradientStopRow
    {
        public string Name { get; set; }
        public string Color { get; set; }
        public string ColorTrans { get; set; }
        public string Position { get; set; }

        public static LineGradientStopRow Get_LineGradientStopRow(Shape shape)
        {
            LineGradientStopRow row = new LineGradientStopRow();

            Section section = shape.Section[(short)VisSectionIndices.visSectionLineGradientStops];
            Row sectionRow = section[0];

            //// TODO(crhodes)
            //// Handle multiple rows

            //row. = sectionRow[VisCellIndices.visLineGradientEnabled].FormulaU;

            return row;
        }
    }
}
