using System;
using System.Windows;
using Microsoft.Office.Interop.Visio;

namespace VNC.Visio.VSTOAddIn.Domain
{
    public class EventsRow
    {
        // TODO(crhodes)
        // Where did TheData come from.  Check ShapeSheet UI
        public string TheData { get; set; }
        public string EventDblClick { get; set; }
        public string Drop { get; set; }
        public string TheText { get; set; }
        public string XFMod { get; set; }
        public string MultiDrop { get; set; }

        public static EventsRow GetRow(Shape shape)
        {
            EventsRow row = new EventsRow();

            Section section = shape.Section[(short)VisSectionIndices.visSectionObject];

            if (Helpers.RowExists(shape, VisSectionIndices.visSectionObject, VisRowIndices.visRowEvent))
            {
                Row sectionRow = section[(short)VisRowIndices.visRowEvent];

                //row.TheData = sectionRow[VisCellIndices.???].FormulaU;
                row.EventDblClick = sectionRow[VisCellIndices.visEvtCellDblClick].FormulaU;
                row.Drop = sectionRow[VisCellIndices.visEvtCellDrop].FormulaU;
                row.TheText = sectionRow[VisCellIndices.visEvtCellTheText].FormulaU;
                row.XFMod = sectionRow[VisCellIndices.visEvtCellXFMod].FormulaU;
                row.MultiDrop = sectionRow[VisCellIndices.visEvtCellMultiDrop].FormulaU;
            }
            else
            {
                MessageBox.Show("No visRowEvent exists");
            }

            return row;
        }

        public static void SetRow(Shape shape, EventsRow events)
        {
            try
            {
                Section section = shape.Section[(short)VisSectionIndices.visSectionObject];

                if (Helpers.RowExists(shape, VisSectionIndices.visSectionObject, VisRowIndices.visRowEvent))
                {
                    Row sectionRow = section[(short)VisRowIndices.visRowEvent];

                    sectionRow[VisCellIndices.visEvtCellDblClick].FormulaU = events.EventDblClick;
                    sectionRow[VisCellIndices.visEvtCellDrop].FormulaU = events.Drop;
                    sectionRow[VisCellIndices.visEvtCellTheText].FormulaU = events.TheText;
                    sectionRow[VisCellIndices.visEvtCellXFMod].FormulaU = events.XFMod;
                    sectionRow[VisCellIndices.visEvtCellMultiDrop].FormulaU = events.MultiDrop;
                }
                else
                {
                    MessageBox.Show("No visRowEvent exists");
                }
            }
            catch (Exception ex)
            {
                Log.ERROR(ex, Common.LOG_CATEGORY);
            }
        }
    }
}
