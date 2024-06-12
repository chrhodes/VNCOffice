using System;

using System.Windows;

using Microsoft.Office.Interop.Visio;

namespace VNC.Visio.VSTOAddIn.Domain
{
    public class EventsRow
    {
        public string TheData { get; set; }
        public string EventDblClick { get; set; }
        public string EventDrop { get; set; }
        public string TheText { get; set; }
        public string EventXFMod { get; set; }
        public string EventMultiDrop { get; set; }

        public static EventsRow GetRow(Shape shape)
        {
            EventsRow row = new EventsRow();

            Section section = shape.Section[(short)VisSectionIndices.visSectionObject];

            if (Helpers.RowExists(shape, VisSectionIndices.visSectionObject, VisRowIndices.visRowEvent))
            {
                Row sectionRow = section[(short)VisRowIndices.visRowEvent];

                //row.TheData = sectionRow[VisCellIndices.???].FormulaU;
                row.EventDblClick = sectionRow[VisCellIndices.visEvtCellDblClick].FormulaU;
                row.EventDrop = sectionRow[VisCellIndices.visEvtCellDrop].FormulaU;
                row.TheText = sectionRow[VisCellIndices.visEvtCellTheText].FormulaU;
                row.EventXFMod = sectionRow[VisCellIndices.visEvtCellXFMod].FormulaU;
                row.EventMultiDrop = sectionRow[VisCellIndices.visEvtCellMultiDrop].FormulaU;
            }
            else
            {
                MessageBox.Show("No visRowEvent exists");
            }

            return row;
        }

        public static void SetRow(Shape shape, EventsRow events)
        {
            // TODO(crhodes)
            // Implement

            if (Helpers.RowExists(shape, VisSectionIndices.visSectionObject, VisRowIndices.visRowEvent))
            {
                if (0 == shape.RowExists[(short)VisSectionIndices.visSectionObject,
                                        (short)VisRowIndices.visRowEvent, 0])
                {
                    MessageBox.Show("No visRowEvent exists");
                }
                else
                {

                }
            }
            else
            {
                MessageBox.Show("No visRowEvent exists");
            }
        }
    }
}
