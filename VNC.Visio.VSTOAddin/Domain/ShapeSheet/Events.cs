using Microsoft.Office.Interop.Visio;

namespace VNC.Visio.VSTOAddIn.Domain
{
    public class Events
    {
        public string TheData { get; set; }
        public string EventDblClick { get; set; }
        public string EventDrop { get; set; }
        public string TheText { get; set; }
        public string EventXFMod { get; set; }
        public string EventMultiDrop { get; set; }

        public static Events Get_Events(Shape shape)
        {
            Events row = new Events();

            Section section = shape.Section[(short)VisSectionIndices.visSectionObject];
            Row sectionRow = section[(short)VisRowIndices.visRowEvent];

            //row.TheData = sectionRow[VisCellIndices.???].FormulaU;
            row.EventDblClick = sectionRow[VisCellIndices.visEvtCellDblClick].FormulaU;
            row.EventDrop = sectionRow[VisCellIndices.visEvtCellDrop].FormulaU;
            row.TheText = sectionRow[VisCellIndices.visEvtCellTheText].FormulaU;
            row.EventXFMod = sectionRow[VisCellIndices.visEvtCellXFMod].FormulaU;
            row.EventMultiDrop = sectionRow[VisCellIndices.visEvtCellMultiDrop].FormulaU;

            return row;
        }
    }
}
