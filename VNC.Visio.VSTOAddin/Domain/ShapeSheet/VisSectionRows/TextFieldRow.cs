using System;
using System.Collections.ObjectModel;
using System.Windows;

using Microsoft.Office.Interop.Visio;

namespace VNC.Visio.VSTOAddIn.Domain
{
    public class TextFieldRow
    {
        public string Name { get; set; }
        public string Format { get; set; }
        public string Value { get; set; }
        public string Calendar { get; set; }
        public string ObjectKind { get; set; }

        public static TextFieldRow GetRow(Shape shape)
        {
            throw new NotImplementedException();
        }

        public static ObservableCollection<TextFieldRow> GetRows(Shape shape)
        {
            var rows = new ObservableCollection<TextFieldRow>();

            if (0 == shape.SectionExists[(short)VisSectionIndices.visSectionTextField, 0])
            {
                MessageBox.Show("No visSectionTextField exists");
            }
            else
            {
                Section section = shape.Section[(short)VisSectionIndices.visSectionTextField];

                var rowCount = section.Count;

                for (short i = 0; i < rowCount; i++)
                {
                    TextFieldRow textFieldRow = new TextFieldRow();

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

                    rows.Add(textFieldRow);
                }
            }

            return rows;
        }
    }
}
