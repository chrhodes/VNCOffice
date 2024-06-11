using System;
using System.Collections.ObjectModel;
using System.Windows;

using Microsoft.Office.Interop.Visio;

namespace VNC.Visio.VSTOAddIn.Domain
{
    public class HyperlinkRow
    {
        public string Name { get; set; }
        public string Description { get; set; }
        public string Address { get; set; }
        public string SubAddress { get; set; }
        public string ExtraInfo { get; set; }
        public string Frame { get; set; }
        public string SortKey { get; set; }
        public string NewWindow { get; set; }
        public string Default { get; set; }
        public string Invisible { get; set; }

        public static HyperlinkRow GetRow(Shape shape)
        {
            HyperlinkRow row = new HyperlinkRow();

            Section section = shape.Section[(short)VisSectionIndices.visSectionHyperlink];
            Row sectionRow = section[0];

            // TODO(crhodes)
            // Handle multiple rows

            row.Name = sectionRow.Name;

            row.Address = sectionRow[VisCellIndices.visHLinkAddress].FormulaU;

            return row;
        }

        public static ObservableCollection<HyperlinkRow> GetRows(Shape shape)
        {
            var rows = new ObservableCollection<HyperlinkRow>();

            if (0 == shape.SectionExists[(short)VisSectionIndices.visSectionHyperlink, 0])
            {
                MessageBox.Show("No visSectionHyperlink exists");
            }
            else
            {
                Section section = shape.Section[(short)VisSectionIndices.visSectionHyperlink];

                var rowCount = section.Count;

                for (short i = 0; i < rowCount; i++)
                {
                    HyperlinkRow hyperlinkRow = new HyperlinkRow();

                    var row = section[i];

                    hyperlinkRow.Name = row.NameU;

                    hyperlinkRow.Description = row[(short)VisCellIndices.visHLinkDescription].FormulaU;
                    hyperlinkRow.Address = row[(short)VisCellIndices.visHLinkAddress].FormulaU;
                    hyperlinkRow.SubAddress = row[(short)VisCellIndices.visHLinkSubAddress].FormulaU;
                    hyperlinkRow.ExtraInfo = row[(short)VisCellIndices.visHLinkExtraInfo].FormulaU;
                    hyperlinkRow.Frame = row[(short)VisCellIndices.visHLinkFrame].FormulaU;
                    hyperlinkRow.SortKey = row[(short)VisCellIndices.visHLinkSortKey].FormulaU;
                    hyperlinkRow.NewWindow = row[(short)VisCellIndices.visHLinkNewWin].FormulaU;
                    hyperlinkRow.Default = row[(short)VisCellIndices.visHLinkDefault].FormulaU;
                    hyperlinkRow.Invisible = row[(short)VisCellIndices.visHLinkInvisible].FormulaU;

                    rows.Add(hyperlinkRow);
                }
            }

            return rows;
        }
    }
}
