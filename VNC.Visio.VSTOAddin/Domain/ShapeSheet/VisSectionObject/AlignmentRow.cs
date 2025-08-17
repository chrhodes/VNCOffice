using System;
using System.Windows;

using Microsoft.Office.Interop.Visio;

namespace VNC.Visio.VSTOAddIn.Domain
{
    public class AlignmentRow
    {

        public string AlignBottom { get; set; }
        public string AlignCenter { get; set; }
        public string AlignLeft { get; set; }
        public string AlignMiddle { get; set; }
        public string AlignRight { get; set; }
        public string AlignTop { get; set; }


        public static AlignmentRow GetRow(Shape shape)
        {
            AlignmentRow row = new AlignmentRow();

            Section section = shape.Section[(short)VisSectionIndices.visSectionObject];

            if (Helpers.RowExists(shape, VisSectionIndices.visSectionObject, VisRowIndices.visRowAlign))
            {
                Row sectionRow = section[(short)VisRowIndices.visRowAlign];

                row.AlignBottom = sectionRow[VisCellIndices.visAlignBottom].FormulaU;
                row.AlignCenter = sectionRow[VisCellIndices.visAlignCenter].FormulaU;
                row.AlignLeft = sectionRow[VisCellIndices.visAlignLeft].FormulaU;
                row.AlignMiddle = sectionRow[VisCellIndices.visAlignMiddle].FormulaU;
                row.AlignRight = sectionRow[VisCellIndices.visAlignRight].FormulaU;
                row.AlignTop = sectionRow[VisCellIndices.visAlignTop].FormulaU;
            }
            else
            {
                MessageBox.Show("No visRowAlign exists");
            }

            return row;
        }

        public static void SetRow(Shape shape, AlignmentRow alignmentRow)
        {
            try
            {
                Section section = shape.Section[(short)VisSectionIndices.visSectionObject];

                if (Helpers.RowExists(shape, VisSectionIndices.visSectionObject, VisRowIndices.visRowAlign))
                {
                    Row sectionRow = section[(short)VisRowIndices.visRowAlign];

                    sectionRow[VisCellIndices.visAlignBottom].FormulaU = alignmentRow.AlignBottom;
                    sectionRow[VisCellIndices.visAlignCenter].FormulaU = alignmentRow.AlignCenter;
                    sectionRow[VisCellIndices.visAlignLeft].FormulaU = alignmentRow.AlignLeft;
                    sectionRow[VisCellIndices.visAlignMiddle].FormulaU = alignmentRow.AlignMiddle;
                    sectionRow[VisCellIndices.visAlignRight].FormulaU = alignmentRow.AlignRight;
                    sectionRow[VisCellIndices.visAlignTop].FormulaU = alignmentRow.AlignTop;
                }
                else
                {
                    MessageBox.Show("No visRowAlign exists");
                }
            }
            catch (Exception ex)
            {
                Log.ERROR(ex, Common.LOG_CATEGORY);
            }
        }
    }
}
