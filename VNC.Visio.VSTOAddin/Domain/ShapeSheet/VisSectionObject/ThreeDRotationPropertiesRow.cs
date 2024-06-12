using System;
using System.Windows;

using Microsoft.Office.Interop.Visio;

namespace VNC.Visio.VSTOAddIn.Domain
{
    public class ThreeDRotationPropertiesRow
    {
        public string RotationXAngle { get; set; }
        public string RotationYAngle { get; set; }
        public string RotationZAngle { get; set; }
        public string RotationType { get; set; }
        public string Perspective { get; set; }
        public string DistanceFromGround { get; set; }
        public string KeepTextFlat { get; set; }

        public static ThreeDRotationPropertiesRow GetRow(Shape shape)
        {
            ThreeDRotationPropertiesRow row = new ThreeDRotationPropertiesRow();

            Section section = shape.Section[(short)VisSectionIndices.visSectionObject];

            if (Helpers.RowExists(shape, VisSectionIndices.visSectionObject, VisRowIndices.visRow3DRotationProperties))
            {
                Row sectionRow = section[(short)VisRowIndices.visRow3DRotationProperties];

                row.RotationXAngle = sectionRow[VisCellIndices.visRotationXAngle].FormulaU;
                row.RotationYAngle = sectionRow[VisCellIndices.visRotationYAngle].FormulaU;
                row.RotationZAngle = sectionRow[VisCellIndices.visRotationZAngle].FormulaU;
                row.RotationType = sectionRow[VisCellIndices.visRotationType].FormulaU;
                row.Perspective = sectionRow[VisCellIndices.visPerspective].FormulaU;
                row.DistanceFromGround = sectionRow[VisCellIndices.visDistanceFromGround].FormulaU;
                row.KeepTextFlat = sectionRow[VisCellIndices.visKeepTextFlat].FormulaU;
            }
            else
            {
                MessageBox.Show("No visRow3DRotationProperties exists");
            }

            return row;
        }

        public static void SetRow(Shape shape, ThreeDRotationPropertiesRow threeDRotationProperties)
        {
            try
            {
                Section section = shape.Section[(short)VisSectionIndices.visSectionObject];

                if (Helpers.RowExists(shape, VisSectionIndices.visSectionObject, VisRowIndices.visRow3DRotationProperties))
                {
                    Row sectionRow = section[(short)VisRowIndices.visRow3DRotationProperties];

                    sectionRow[VisCellIndices.visRotationXAngle].FormulaU = threeDRotationProperties.RotationXAngle;
                    sectionRow[VisCellIndices.visRotationYAngle].FormulaU = threeDRotationProperties.RotationYAngle;
                    sectionRow[VisCellIndices.visRotationZAngle].FormulaU = threeDRotationProperties.RotationZAngle;
                    sectionRow[VisCellIndices.visRotationType].FormulaU = threeDRotationProperties.RotationType;
                    sectionRow[VisCellIndices.visPerspective].FormulaU = threeDRotationProperties.Perspective;
                    sectionRow[VisCellIndices.visDistanceFromGround].FormulaU = threeDRotationProperties.DistanceFromGround;
                    sectionRow[VisCellIndices.visKeepTextFlat].FormulaU = threeDRotationProperties.KeepTextFlat;
                }
                else
                {
                    MessageBox.Show("No visRow3DRotationProperties exists");
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }
    }
}
