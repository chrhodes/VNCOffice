using Microsoft.Office.Interop.Visio;

namespace VNC.Visio.VSTOAddIn.Domain
{
    public class FillGradientStopRow
    {
        public string Color { get; set; }
        public string ColorTrans { get; set; }
        public string Position { get; set; }

        public static Domain.FillGradientStopRow GetRow(Shape shape)
        {
            FillGradientStopRow row = new FillGradientStopRow();

            // Shape Transform Section is part of object
            Section section = shape.Section[(short)VisSectionIndices.visSectionObject];
            Row sectionRow = section[(short)VisRowIndices.visRowGradientStop];

            // TODO(crhodes)
            // Handle multiple rows

            //row.NoObjHandles = sectionRow[VisCellIndices.visNoObjHandles].FormulaU;

            return row;
        }
    }
}
