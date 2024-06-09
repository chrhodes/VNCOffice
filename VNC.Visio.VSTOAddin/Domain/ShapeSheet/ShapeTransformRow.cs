using System;

using Microsoft.Office.Interop.Visio;

namespace VNC.Visio.VSTOAddIn.Domain
{
    public class ShapeTransformRow
    {
        public string Width { get; set; }
        public string Height { get; set; }
        public string Angle { get; set; }
        public string PinX { get; set; }
        public string PinY { get; set; }
        public string LocPinX { get; set; }
        public string LocPinY { get; set; }
        public string FlipX { get; set; }
        public string FlipY { get; set; }
        public string ResizeMode { get; set; }


        public static ShapeTransformRow Get_ShapeTransform(Shape shape)
        {
            ShapeTransformRow row = new ShapeTransformRow();

            Section section = shape.Section[(short)VisSectionIndices.visSectionObject];
            Row sectionRow = section[(short)VisRowIndices.visRowXFormOut];

            row.Width = sectionRow[VisCellIndices.visXFormWidth].FormulaU;
            row.Height = sectionRow[VisCellIndices.visXFormHeight].FormulaU;
            row.Angle = sectionRow[VisCellIndices.visXFormAngle].FormulaU;
            row.PinX = sectionRow[VisCellIndices.visXFormPinX].FormulaU;
            row.PinY = sectionRow[VisCellIndices.visXFormPinY].FormulaU;
            row.LocPinX = sectionRow[VisCellIndices.visXFormLocPinX].FormulaU;
            row.LocPinY = sectionRow[VisCellIndices.visXFormLocPinY].FormulaU;
            row.FlipX = sectionRow[VisCellIndices.visXFormFlipX].FormulaU;
            row.FlipY = sectionRow[VisCellIndices.visXFormFlipY].FormulaU;
            row.ResizeMode = sectionRow[VisCellIndices.visXFormResizeMode].FormulaU;

            return row;
        }

        public static void Set_ShapeTransform_Section(Shape shape, ShapeTransformRow shapeTransform)
        {
            try
            {
                Section section = shape.Section[(short)VisSectionIndices.visSectionObject];
                Row sectionRow = section[(short)VisRowIndices.visRowXFormOut];

                sectionRow[VisCellIndices.visXFormWidth].FormulaU = shapeTransform.Width;
                sectionRow[VisCellIndices.visXFormHeight].FormulaU = shapeTransform.Height;
                sectionRow[VisCellIndices.visXFormAngle].FormulaU = shapeTransform.Angle;
                sectionRow[VisCellIndices.visXFormPinX].FormulaU = shapeTransform.PinX;
                sectionRow[VisCellIndices.visXFormPinY].FormulaU = shapeTransform.PinY;
                sectionRow[VisCellIndices.visXFormLocPinX].FormulaU = shapeTransform.LocPinX;
                sectionRow[VisCellIndices.visXFormLocPinY].FormulaU = shapeTransform.LocPinY;
                sectionRow[VisCellIndices.visXFormFlipX].FormulaU = shapeTransform.FlipX;
                sectionRow[VisCellIndices.visXFormFlipY].FormulaU = shapeTransform.FlipY;
                sectionRow[VisCellIndices.visXFormResizeMode].FormulaU = shapeTransform.ResizeMode;
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }
    }
}
