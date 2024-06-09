using System.Collections.ObjectModel;

using Microsoft.Office.Interop.Visio;

namespace VNC.Visio.VSTOAddIn.Domain
{
    public class LineGradientStopRow
    {
        public string Name { get; set; }
        public string Color { get; set; }
        public string ColorTrans { get; set; }
        public string Position { get; set; }

        public static LineGradientStopRow GetRow(Shape shape)
        {
            LineGradientStopRow row = new LineGradientStopRow();

            Section section = shape.Section[(short)VisSectionIndices.visSectionLineGradientStops];
            Row sectionRow = section[0];

            //// TODO(crhodes)
            //// Handle multiple rows

            //row. = sectionRow[VisCellIndices.visLineGradientEnabled].FormulaU;

            return row;
        }

        public static ObservableCollection<LineGradientStopRow> GetRows(Shape shape)
        {
            var rows = new ObservableCollection<LineGradientStopRow>();

            Section section = shape.Section[(short)VisSectionIndices.visSectionLineGradientStops];

            var rowCount = section.Count;

            for (short i = 0; i < rowCount; i++)
            {
                LineGradientStopRow lineGradientStopRow = new LineGradientStopRow();

                var row = section[i];

                // TODO(crhodes)
                // Implement

                //layerRow.Name = row[(short)VisCellIndices.visLayerName].FormulaU;
                //layerRow.Visible = row[(short)VisCellIndices.visLayerVisible].FormulaU;
                //layerRow.Print = row[(short)VisCellIndices.visLayerPrint].FormulaU;
                //layerRow.Active = row[(short)VisCellIndices.visLayerActive].FormulaU;
                //layerRow.Lock = row[(short)VisCellIndices.visLayerLock].FormulaU;
                //layerRow.Snap = row[(short)VisCellIndices.visLayerSnap].FormulaU;
                //layerRow.Glue = row[(short)VisCellIndices.visLayerGlue].FormulaU;
                //layerRow.Color = row[(short)VisCellIndices.visLayerColor].FormulaU;
                //layerRow.Transparency = row[(short)VisCellIndices.visLayerColorTrans].FormulaU;

                //// NOTE(crhodes)
                //// There are a few more VisCellIndices.  See what they do
                ////VisCellIndices.visLayerMember
                ////VisCellIndices.visLayerStatus

                rows.Add(lineGradientStopRow);
            }

            return rows;
        }
    }
}
