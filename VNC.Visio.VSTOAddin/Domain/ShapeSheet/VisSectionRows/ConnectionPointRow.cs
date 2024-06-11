using System.Collections.ObjectModel;
using System.Windows;

using Microsoft.Office.Interop.Visio;

namespace VNC.Visio.VSTOAddIn.Domain
{
    public class ConnectionPointRow
    {
        public string Name { get; set; }
        public string X { get; set; }
        public string Y { get; set; }
        public string DirX { get; set; }
        public string A { get; set; }
        public string DirY { get; set; }
        public string B { get; set; }
        public string Type { get; set; }
        public string C { get; set; }
        public string D { get; set; }

        public static ConnectionPointRow GetRow(Shape shape)
        {
            ConnectionPointRow row = new ConnectionPointRow();

            Section section = shape.Section[(short)VisSectionIndices.visSectionConnectionPts];
            Row sectionRow = section[0];

            // TODO(crhodes)
            // Handle multiple rows
            row.Name = sectionRow.Name;

            row.X = sectionRow[VisCellIndices.visCnnctX].FormulaU;

            return row;
        }

        public static ObservableCollection<ConnectionPointRow> GetRows(Shape shape)
        {
            var rows = new ObservableCollection<ConnectionPointRow>();

            if (0 == shape.SectionExists[(short)VisSectionIndices.visSectionConnectionPts, 0])
            {
                MessageBox.Show("No visSectionConnectionPts exists");
            }
            else
            {
                Section section = shape.Section[(short)VisSectionIndices.visSectionConnectionPts];

                var rowCount = section.Count;

                for (short i = 0; i < rowCount; i++)
                {
                    ConnectionPointRow connectionPointRow = new ConnectionPointRow();

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

                    // NOTE(crhodes)
                    // There are a few more VisCellIndices.  See what they do
                    //VisCellIndices.visLayerMember
                    //VisCellIndices.visLayerStatus

                    rows.Add(connectionPointRow);
                }
            }

            return rows;
        }
    }
}
