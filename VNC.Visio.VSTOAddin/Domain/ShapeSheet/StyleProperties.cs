using Microsoft.Office.Interop.Visio;

namespace VNC.Visio.VSTOAddIn.Domain
{
    public class StyleProperties
    {
        public string EnableTextProps { get; set; }
        public string EnableLineProps { get; set; }
        public string EnableFillProps { get; set; }
        public string HideForApply { get; set; }

        public static StyleProperties GetRow(Shape shape)
        {
            StyleProperties row = new StyleProperties();

            // Shape Transform Section is part of object
            Section section = shape.Section[(short)VisSectionIndices.visSectionObject];
            Row sectionRow = section[(short)VisRowIndices.visRowStyle];

            //row.LeftMargin = sectionRow[VisCellIndices.visTxtBlkLeftMargin].FormulaU;

            return row;
        }
    }
}
