using System;

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
            Row sectionRow = section[(short)VisRowIndices.visRowEvent];

            //row.TheData = sectionRow[VisCellIndices.???].FormulaU;
            row.EventDblClick = sectionRow[VisCellIndices.visEvtCellDblClick].FormulaU;
            row.EventDrop = sectionRow[VisCellIndices.visEvtCellDrop].FormulaU;
            row.TheText = sectionRow[VisCellIndices.visEvtCellTheText].FormulaU;
            row.EventXFMod = sectionRow[VisCellIndices.visEvtCellXFMod].FormulaU;
            row.EventMultiDrop = sectionRow[VisCellIndices.visEvtCellMultiDrop].FormulaU;

            return row;
        }

        public static void SetRow(Shape shape, EventsRow events)
        {
            // TODO(crhodes)
         // Implement

         //    try
         //    {
         //        Section section = shape.Section[(short)VisSectionIndices.visSectionObject];
         //        Row sectionRow = section[(short)VisRowIndices.visRowMisc];

            //        sectionRow[VisCellIndices.visBegTrigger].FormulaU = glueInfo.BegTrigger;
            //        sectionRow[VisCellIndices.visEndTrigger].FormulaU = glueInfo.EndTrigger;
            //        sectionRow[VisCellIndices.visGlueType].FormulaU = glueInfo.GlueType;
            //        sectionRow[VisCellIndices.visWalkPref].FormulaU = glueInfo.WalkPreference;
            //    }
            //    catch (Exception ex)
            //    {
            //        Log.Error(ex, Common.LOG_CATEGORY);
            //    }
        }
    }
}
