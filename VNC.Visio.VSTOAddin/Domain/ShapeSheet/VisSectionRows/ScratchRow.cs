using System.Collections.ObjectModel;
using System.Windows;

using Microsoft.Office.Interop.Visio;

namespace VNC.Visio.VSTOAddIn.Domain
{
    public class ScratchRow
    {
        public string Row { get; set; }
        public string X { get; set; }
        public string Y { get; set; }
        public string A { get; set; }
        public string B { get; set; }
        public string C { get; set; }
        public string D { get; set; }


        public static ScratchRow GetRow(Shape shape, short rowNumber)
        {
            ScratchRow row = new ScratchRow();

            Section section = shape.Section[(short)VisSectionIndices.visSectionScratch];
            Row sectionRow = section[rowNumber];

            return row;
        }

        public static ObservableCollection<ScratchRow> GetRows(Shape shape)
        {
            var rows = new ObservableCollection<ScratchRow>();
            if (0 == shape.SectionExists[(short)VisSectionIndices.visSectionScratch, 0])
            {
                MessageBox.Show("No visSectionScratch exists");
            }
            else
            {
                Section section = shape.Section[(short)VisSectionIndices.visSectionScratch];

                var rowCount = section.Count;

                for (short i = 0; i < rowCount; i++)
                {
                    ScratchRow scratchRow = new ScratchRow();

                    var row = section[i];

                    scratchRow.Row = $"{i}";

                    scratchRow.X = row[(short)VisCellIndices.visScratchX].FormulaU;
                    scratchRow.Y = row[(short)VisCellIndices.visScratchY].FormulaU;
                    scratchRow.A = row[(short)VisCellIndices.visScratchA].FormulaU;
                    scratchRow.B = row[(short)VisCellIndices.visScratchB].FormulaU;
                    scratchRow.C = row[(short)VisCellIndices.visScratchC].FormulaU;
                    scratchRow.D = row[(short)VisCellIndices.visScratchD].FormulaU;

                    rows.Add(scratchRow);
                }
            }

            return rows;
        }
    }
}
