using Microsoft.Office.Interop.Visio;

namespace VNC.Visio.VSTOAddIn.Domain
{
    public class FillFormatRow
    {
        public string FillForegnd { get; set; }
        public string ShdwForegnd { get; set; }
        public string ShapeShdwType { get; set; }
        public string FillForegndTrans { get; set; }
        public string ShdwForegndTrans { get; set; }
        public string ShapeShdwObliqueAngle { get; set; }
        public string FillBkgnd { get; set; }
        public string ShdwPattern { get; set; }
        public string ShapeShdwScaleFactor { get; set; }
        public string FillBkgndTrans { get; set; }
        public string ShapeShdwOffsetX { get; set; }
        public string ShapeShdwOffsetY { get; set; }
        public string ShapeShdwBlur { get; set; }
        public string FillPattern { get; set; }
        public string ShapeShdwShow { get; set; }

        public static FillFormatRow GetRow(Shape shape)
        {
            FillFormatRow row = new FillFormatRow();

            Section section = shape.Section[(short)VisSectionIndices.visSectionObject];
            Row sectionRow = section[(short)VisRowIndices.visRowFill];

            row.FillForegnd = sectionRow[VisCellIndices.visFillForegnd].FormulaU;
            row.ShdwForegnd = sectionRow[VisCellIndices.visFillShdwForegnd].FormulaU;
            row.ShapeShdwType = sectionRow[VisCellIndices.visFillShdwType].FormulaU;
            row.FillForegndTrans = sectionRow[VisCellIndices.visFillForegndTrans].FormulaU;
            row.ShdwForegndTrans = sectionRow[VisCellIndices.visFillShdwForegndTrans].FormulaU;
            row.ShapeShdwObliqueAngle = sectionRow[VisCellIndices.visFillShdwObliqueAngle].FormulaU;
            row.FillBkgnd = sectionRow[VisCellIndices.visFillBkgnd].FormulaU;
            row.ShdwPattern = sectionRow[VisCellIndices.visFillShdwPattern].FormulaU;
            row.ShapeShdwScaleFactor = sectionRow[VisCellIndices.visFillShdwScaleFactor].FormulaU;
            row.FillBkgndTrans = sectionRow[VisCellIndices.visFillBkgndTrans].FormulaU;
            row.ShapeShdwOffsetX = sectionRow[VisCellIndices.visFillShdwOffsetX].FormulaU;
            row.ShapeShdwOffsetY = sectionRow[VisCellIndices.visFillShdwOffsetY].FormulaU;
            row.ShapeShdwBlur = sectionRow[VisCellIndices.visFillShdwShow].FormulaU;
            row.FillPattern = sectionRow[VisCellIndices.visFillShdwPattern].FormulaU;
            row.ShapeShdwShow = sectionRow[VisCellIndices.visFillShdwShow].FormulaU;

            return row;
        }

        public static void SetRow(Shape shape, FillFormatRow fillFormat)
        {
            Section section = shape.Section[(short)VisSectionIndices.visSectionObject];
            Row sectionRow = section[(short)VisRowIndices.visRowFill];

            sectionRow[VisCellIndices.visFillForegnd].FormulaU = fillFormat.FillForegnd;
            sectionRow[VisCellIndices.visFillShdwForegnd].FormulaU = fillFormat.ShdwForegnd;
            sectionRow[VisCellIndices.visFillShdwType].FormulaU = fillFormat.ShapeShdwType;
            sectionRow[VisCellIndices.visFillForegndTrans].FormulaU = fillFormat.FillForegndTrans;
            sectionRow[VisCellIndices.visFillShdwForegndTrans].FormulaU = fillFormat.ShdwForegndTrans;
            sectionRow[VisCellIndices.visFillShdwObliqueAngle].FormulaU = fillFormat.ShapeShdwObliqueAngle;
            sectionRow[VisCellIndices.visFillBkgnd].FormulaU = fillFormat.FillBkgnd;
            sectionRow[VisCellIndices.visFillShdwPattern].FormulaU = fillFormat.ShdwPattern;
            sectionRow[VisCellIndices.visFillShdwScaleFactor].FormulaU = fillFormat.ShapeShdwScaleFactor;
            sectionRow[VisCellIndices.visFillBkgndTrans].FormulaU = fillFormat.FillBkgndTrans;
            sectionRow[VisCellIndices.visFillShdwOffsetX].FormulaU = fillFormat.ShapeShdwOffsetX;
            sectionRow[VisCellIndices.visFillShdwOffsetY].FormulaU = fillFormat.ShapeShdwOffsetY;
            sectionRow[VisCellIndices.visFillShdwShow].FormulaU = fillFormat.ShapeShdwBlur;
            sectionRow[VisCellIndices.visFillShdwPattern].FormulaU = fillFormat.FillPattern;
            sectionRow[VisCellIndices.visFillShdwShow].FormulaU = fillFormat.ShapeShdwShow;
        }
    }
}
