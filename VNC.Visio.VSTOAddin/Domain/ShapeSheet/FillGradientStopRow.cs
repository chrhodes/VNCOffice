using System.Collections.ObjectModel;

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

        public static ObservableCollection<FillGradientStopRow> GetRows(Shape shape)
        {
            var rows = new ObservableCollection<FillGradientStopRow>();

            Section section = shape.Section[(short)VisSectionIndices.visSectionFillGradientStops];

            var rowCount = section.Count;

            for (short i = 0; i < rowCount; i++)
            {
                FillGradientStopRow fillGradientStopRow = new FillGradientStopRow();

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

                rows.Add(fillGradientStopRow);
            }

            return rows;
        }
    }
}
