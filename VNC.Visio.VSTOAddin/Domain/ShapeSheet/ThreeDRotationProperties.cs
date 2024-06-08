using System;

using Microsoft.Office.Interop.Visio;

namespace VNC.Visio.VSTOAddIn.Domain
{
    public class ThreeDRotationProperties
    {
        public string RotationXAngle { get; set; }
        public string RotationYAngle { get; set; }
        public string RotationZAngle { get; set; }
        public string RotationType { get; set; }
        public string Perspective { get; set; }
        public string DistanceFromGround { get; set; }
        public string KeepTextFlat { get; set; }

        public static ThreeDRotationProperties GetRow(Shape shape)
        {
            ThreeDRotationProperties row = new ThreeDRotationProperties();

            Section section = shape.Section[(short)VisSectionIndices.visSectionObject];
            Row sectionRow = section[(short)VisRowIndices.visRow3DRotationProperties];

            row.RotationXAngle = sectionRow[VisCellIndices.visRotationXAngle].FormulaU;
            row.RotationYAngle = sectionRow[VisCellIndices.visRotationYAngle].FormulaU;
            row.RotationZAngle = sectionRow[VisCellIndices.visRotationZAngle].FormulaU;
            row.RotationType = sectionRow[VisCellIndices.visRotationType].FormulaU;
            row.Perspective = sectionRow[VisCellIndices.visPerspective].FormulaU;
            row.DistanceFromGround = sectionRow[VisCellIndices.visDistanceFromGround].FormulaU;
            row.KeepTextFlat = sectionRow[VisCellIndices.visKeepTextFlat].FormulaU;

            return row;
        }

        public static void Set_ThreeDRotationProperties_Section(Shape shape, ThreeDRotationProperties threeDRotationProperties)
        {
            try
            {
                Section section = shape.Section[(short)VisSectionIndices.visSectionObject];
                Row sectionRow = section[(short)VisRowIndices.visRow3DRotationProperties];

                sectionRow[VisCellIndices.visRotationXAngle].FormulaU = threeDRotationProperties.RotationXAngle;
                sectionRow[VisCellIndices.visRotationYAngle].FormulaU = threeDRotationProperties.RotationYAngle;
                sectionRow[VisCellIndices.visRotationZAngle].FormulaU = threeDRotationProperties.RotationZAngle;
                sectionRow[VisCellIndices.visRotationType].FormulaU = threeDRotationProperties.RotationType;
                sectionRow[VisCellIndices.visPerspective].FormulaU = threeDRotationProperties.Perspective;
                sectionRow[VisCellIndices.visDistanceFromGround].FormulaU = threeDRotationProperties.DistanceFromGround;
                sectionRow[VisCellIndices.visKeepTextFlat].FormulaU = threeDRotationProperties.KeepTextFlat;
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }
    }
}
