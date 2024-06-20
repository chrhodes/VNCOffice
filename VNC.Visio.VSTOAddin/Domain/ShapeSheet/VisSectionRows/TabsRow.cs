using System;
using System.Collections.ObjectModel;
using System.Windows;

using Microsoft.Office.Interop.Visio;

namespace VNC.Visio.VSTOAddIn.Domain
{
    public class TabsRow
    {
        // TODO(crhodes)
        // This looks tricky as there are an unknown number of tabs
        public string Name { get; set; }
        public string Position1 { get; set; }
        public string Alignment1 { get; set; }
        public string Position2 { get; set; }
        public string Alignment2 { get; set; }

        public static TabsRow GetRow(Shape shape)
        {
            TabsRow row = new TabsRow();

            Section section = shape.Section[(short)VisSectionIndices.visSectionObject];
            Row sectionRow = section[(short)VisRowIndices.visRowTab];

            // TODO(crhodes)
            // Handle multiple rows

            //row.Name = sectionRow.Name;
            //row.Visible = sectionRow[VisCellIndices.visLayerVisible].FormulaU;
            //row.Print = sectionRow[VisCellIndices.visLayerPrint].FormulaU;
            //row.Active = sectionRow[VisCellIndices.visLayerActive].FormulaU;
            //row.Lock = sectionRow[VisCellIndices.visLayerLock].FormulaU;
            //row.Snap = sectionRow[VisCellIndices.visLayerLock].FormulaU;
            //row.Glue = sectionRow[VisCellIndices.visLayerGlue].FormulaU;
            //row.Color = sectionRow[VisCellIndices.visLayerColor].FormulaU;
            //row.Transparency = sectionRow[VisCellIndices.visLayerColorTrans].FormulaU;

            return row;
        }

        public static ObservableCollection<TabsRow> GetRows(Shape shape)
        {
            var rows = new ObservableCollection<TabsRow>();

            if (0 == shape.SectionExists[(short)VisSectionIndices.visSectionTab, 0])
            {
                MessageBox.Show("No visSectionLayer exists");
            }
            else
            {
                Section section = shape.Section[(short)VisSectionIndices.visSectionTab];

                var rowCount = section.Count;

                for (short i = 0; i < rowCount; i++)
                {
                    TabsRow tabRow = new TabsRow();

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

                    rows.Add(tabRow);
                }
            }

            return rows;
        }
    }
}
