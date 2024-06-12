using System;
using System.Windows;

using Microsoft.Office.Interop.Visio;

namespace VNC.Visio.VSTOAddIn.Domain
{
    public class GradientPropertiesRow
    {
        public string LineGradientDir { get; set; }
        public string LineGradientAngle { get; set; }
        public string FillGradientDir { get; set; }
        public string FillGradientAngle { get; set; }
        public string LineGradientEnabled { get; set; }
        public string FillGradientEnabled { get; set; }
        public string RotateGradientWithShape { get; set; }
        public string UseGroupGradient { get; set; }

        public static GradientPropertiesRow GetRow(Shape shape)
        {
            GradientPropertiesRow row = new GradientPropertiesRow();

            Section section = shape.Section[(short)VisSectionIndices.visSectionObject];

            if (Helpers.RowExists(shape, VisSectionIndices.visSectionObject, VisRowIndices.visRowGradientProperties))
            {
                Row sectionRow = section[(short)VisRowIndices.visRowGradientProperties];

                row.LineGradientDir = sectionRow[VisCellIndices.visLineGradientDir].FormulaU;
                row.LineGradientAngle = sectionRow[VisCellIndices.visLineGradientAngle].FormulaU;
                row.FillGradientDir = sectionRow[VisCellIndices.visFillGradientDir].FormulaU;
                row.FillGradientAngle = sectionRow[VisCellIndices.visFillGradientAngle].FormulaU;
                row.LineGradientEnabled = sectionRow[VisCellIndices.visLineGradientEnabled].FormulaU;
                row.FillGradientEnabled = sectionRow[VisCellIndices.visFillGradientEnabled].FormulaU;
                row.RotateGradientWithShape = sectionRow[VisCellIndices.visRotateGradientWithShape].FormulaU;
                row.UseGroupGradient = sectionRow[VisCellIndices.visUseGroupGradient].FormulaU;
            }
            else
            {
                MessageBox.Show("No visRowGradientProperties exists");
            }

            return row;
        }

        public static void SetRow(Shape shape, GradientPropertiesRow gradientProperties)
        {
            try
            {
                Section section = shape.Section[(short)VisSectionIndices.visSectionObject];

                if (Helpers.RowExists(shape, VisSectionIndices.visSectionObject, VisRowIndices.visRowGradientProperties))
                {
                    Row sectionRow = section[(short)VisRowIndices.visRowGradientProperties];

                    sectionRow[VisCellIndices.visLineGradientDir].FormulaU = gradientProperties.LineGradientDir;
                    sectionRow[VisCellIndices.visLineGradientAngle].FormulaU = gradientProperties.LineGradientAngle;
                    sectionRow[VisCellIndices.visFillGradientDir].FormulaU = gradientProperties.FillGradientDir;
                    sectionRow[VisCellIndices.visFillGradientAngle].FormulaU = gradientProperties.FillGradientAngle;
                    sectionRow[VisCellIndices.visLineGradientEnabled].FormulaU = gradientProperties.LineGradientEnabled;
                    sectionRow[VisCellIndices.visFillGradientEnabled].FormulaU = gradientProperties.FillGradientEnabled;
                    sectionRow[VisCellIndices.visRotateGradientWithShape].FormulaU = gradientProperties.RotateGradientWithShape;
                    sectionRow[VisCellIndices.visUseGroupGradient].FormulaU = gradientProperties.UseGroupGradient;
                }
                else
                {
                    MessageBox.Show("No visRowGradientProperties exists");
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }
    }
}
