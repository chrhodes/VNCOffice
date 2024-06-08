using System.Collections.ObjectModel;

using Microsoft.Office.Interop.Visio;

namespace VNC.Visio.VSTOAddIn.Domain
{
    public class LayerRow
    {
        public string Name { get; set; }
        public string Visible { get; set; }
        public string Print { get; set; }
        public string Active { get; set; }
        public string Lock { get; set; }
        public string Snap { get; set; }
        public string Glue { get; set; }
        public string Color { get; set; }
        public string Transparency { get; set; }

        public static LayerRow GetRow(Shape shape)
        {
            LayerRow row = new LayerRow();

            Section section = shape.Section[(short)VisSectionIndices.visSectionObject];
            Row sectionRow = section[(short)VisRowIndices.visRowLayer];

            // TODO(crhodes)
            // Handle multiple rows

            row.Name = sectionRow.Name;
            row.Visible = sectionRow[VisCellIndices.visLayerVisible].FormulaU;
            row.Print = sectionRow[VisCellIndices.visLayerPrint].FormulaU;
            row.Active = sectionRow[VisCellIndices.visLayerActive].FormulaU;
            row.Lock = sectionRow[VisCellIndices.visLayerLock].FormulaU;
            row.Snap = sectionRow[VisCellIndices.visLayerLock].FormulaU;
            row.Glue = sectionRow[VisCellIndices.visLayerGlue].FormulaU;
            row.Color = sectionRow[VisCellIndices.visLayerColor].FormulaU;
            row.Transparency = sectionRow[VisCellIndices.visLayerColorTrans].FormulaU;

            return row;
        }

        public static ObservableCollection<LayerRow> GetRows(Shape shape)
        {
            var rows = new ObservableCollection<LayerRow>();

            Section section = shape.Section[(short)VisSectionIndices.visSectionLayer];

            var rowCount = section.Count;

            for (short i = 0; i < rowCount; i++)
            {
                LayerRow layerRow = new LayerRow();

                var row = section[i];

                layerRow.Name = row[(short)VisCellIndices.visLayerName].FormulaU;
                layerRow.Visible = row[(short)VisCellIndices.visLayerVisible].FormulaU;
                layerRow.Print = row[(short)VisCellIndices.visLayerPrint].FormulaU;
                layerRow.Active = row[(short)VisCellIndices.visLayerActive].FormulaU;
                layerRow.Lock = row[(short)VisCellIndices.visLayerLock].FormulaU;
                layerRow.Snap = row[(short)VisCellIndices.visLayerSnap].FormulaU;
                layerRow.Glue = row[(short)VisCellIndices.visLayerGlue].FormulaU;
                layerRow.Color = row[(short)VisCellIndices.visLayerColor].FormulaU;
                layerRow.Transparency = row[(short)VisCellIndices.visLayerColorTrans].FormulaU;

                // NOTE(crhodes)
                // There are a few more VisCellIndices.  See what they do
                //VisCellIndices.visLayerMember
                //VisCellIndices.visLayerStatus

                rows.Add(layerRow);
            }

            return rows;
        }
    }
}
