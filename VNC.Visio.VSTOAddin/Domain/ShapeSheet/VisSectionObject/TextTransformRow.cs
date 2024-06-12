using System;
using System.Windows;


using Microsoft.Office.Interop.Visio;

namespace VNC.Visio.VSTOAddIn.Domain
{
    // TODO(crhodes)
    // Verify Name shyould this be TextXForm
    public class TextTransformRow
    {
        public string TxtWidth { get; set; }
        public string TxtHeight { get; set; }
        public string TxtAngle { get; set; }
        public string TxtPinX { get; set; }
        public string TxtPinY { get; set; }
        public string TxtLocPinX { get; set; }
        public string TxtLocPinY { get; set; }

        public static TextTransformRow GetRow(Shape shape)
        {
            TextTransformRow row = new TextTransformRow();

            Section section = shape.Section[(short)VisSectionIndices.visSectionObject];

            if (Helpers.RowExists(shape, VisSectionIndices.visSectionObject, VisRowIndices.visRowTextXForm))
            {
                MessageBox.Show("No visRowTextXForm exists");
            }
            else
            {
                Row sectionRow = section[(short)VisRowIndices.visRowTextXForm];

                row.TxtWidth = sectionRow[VisCellIndices.visXFormWidth].FormulaU;
                row.TxtHeight = sectionRow[VisCellIndices.visXFormHeight].FormulaU;
                row.TxtAngle = sectionRow[VisCellIndices.visXFormAngle].FormulaU;
                row.TxtPinX = sectionRow[VisCellIndices.visXFormPinX].FormulaU;
                row.TxtPinY = sectionRow[VisCellIndices.visXFormPinY].FormulaU;
                row.TxtLocPinX = sectionRow[VisCellIndices.visXFormLocPinX].FormulaU;
                row.TxtLocPinY = sectionRow[VisCellIndices.visXFormLocPinY].FormulaU;
            }

            return row;
        }

        public static void SetRow(Shape shape, TextTransformRow textTransform)
        {
            try
            {
                Section section = shape.Section[(short)VisSectionIndices.visSectionObject];

                if (Helpers.RowExists(shape, VisSectionIndices.visSectionObject, VisRowIndices.visRowTextXForm))
                {
                    MessageBox.Show("No visRowTextXForm exists");
                }
                else
                {
                    Row sectionRow = section[(short)VisRowIndices.visRowTextXForm];

                    sectionRow[VisCellIndices.visXFormWidth].FormulaForceU = textTransform.TxtWidth;
                    sectionRow[VisCellIndices.visXFormHeight].FormulaForceU = textTransform.TxtHeight;
                    sectionRow[VisCellIndices.visXFormAngle].FormulaForceU = textTransform.TxtAngle;
                    sectionRow[VisCellIndices.visXFormPinX].FormulaForceU = textTransform.TxtPinX;
                    sectionRow[VisCellIndices.visXFormPinY].FormulaForceU = textTransform.TxtPinY;
                    sectionRow[VisCellIndices.visXFormLocPinX].FormulaForceU = textTransform.TxtLocPinX;
                    sectionRow[VisCellIndices.visXFormLocPinY].FormulaForceU = textTransform.TxtLocPinY;
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }
    }
}
