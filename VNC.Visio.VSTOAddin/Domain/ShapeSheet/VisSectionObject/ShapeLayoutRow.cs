using System;
using System.Windows;


using Microsoft.Office.Interop.Visio;

namespace VNC.Visio.VSTOAddIn.Domain
{
    public class ShapeLayoutRow
    {
        public string ShapePermeableX { get; set; }
        public string ShapePermeableY { get; set; }
        public string ShapeFixedCode { get; set; }
        public string ConLineJumpDirX { get; set; } 
        public string ConLineJumpDirY { get; set; }
        public string ConLineJumpCode { get; set; }
        public string ShapePlaceFlip { get; set; }
        public string ShapePlaceStyle { get; set; }
        public string ShapePlowCode { get; set; }
        public string ConLineJumpStyle { get; set; }
        public string ConLineRouteExt { get; set; }
        public string DisplayLevel { get; set; }
        public string ShapeRouteStyle { get; set; }
        public string ConFixedCode { get; set; }
        public string ShapeSplit { get; set; }
        public string ShapeSplittable { get; set; }
        public string Relationships { get; set; }


        public static ShapeLayoutRow GetRow(Shape shape)
        {
            ShapeLayoutRow row = new ShapeLayoutRow();

            Section section = shape.Section[(short)VisSectionIndices.visSectionObject];

            if (Helpers.RowExists(shape, VisSectionIndices.visSectionObject, VisRowIndices.visRowShapeLayout))
            {
                MessageBox.Show("No visRowShapeLayout exists");
            }
            else
            {
                Row sectionRow = section[(short)VisRowIndices.visRowShapeLayout];

                row.ShapePermeableX = sectionRow[VisCellIndices.visSLOPermX].FormulaU;
                row.ShapePermeableY = sectionRow[VisCellIndices.visSLOPermY].FormulaU;
                row.ShapeFixedCode = sectionRow[VisCellIndices.visSLOFixedCode].FormulaU;
                row.ConLineJumpDirX = sectionRow[VisCellIndices.visSLOJumpDirX].FormulaU;
                row.ConLineJumpDirY = sectionRow[VisCellIndices.visSLOJumpDirY].FormulaU;
                row.ConLineJumpCode = sectionRow[VisCellIndices.visSLOJumpCode].FormulaU;
                row.ShapePlaceFlip = sectionRow[VisCellIndices.visSLOPlaceFlip].FormulaU;
                row.ShapePlaceStyle = sectionRow[VisCellIndices.visSLOPlaceStyle].FormulaU;
                row.ShapePlowCode = sectionRow[VisCellIndices.visSLOPlowCode].FormulaU;
                row.ConLineJumpStyle = sectionRow[VisCellIndices.visSLOJumpStyle].FormulaU;
                row.ConLineRouteExt = sectionRow[VisCellIndices.visSLOLineRouteExt].FormulaU;
                row.DisplayLevel = sectionRow[VisCellIndices.visSLODisplayLevel].FormulaU;
                row.ShapeRouteStyle = sectionRow[VisCellIndices.visSLORouteStyle].FormulaU;
                row.ConFixedCode = sectionRow[VisCellIndices.visSLOConFixedCode].FormulaU;
                row.ShapeSplit = sectionRow[VisCellIndices.visSLOSplit].FormulaU;
                row.ShapeSplittable = sectionRow[VisCellIndices.visSLOSplittable].FormulaU;
                row.Relationships = sectionRow[VisCellIndices.visSLORelationships].FormulaU;
            }

            return row;
        }

        public static void SetRow(Shape shape, ShapeLayoutRow shapeLayout)
        {
            try
            {
                Section section = shape.Section[(short)VisSectionIndices.visSectionObject];

                if (Helpers.RowExists(shape, VisSectionIndices.visSectionObject, VisRowIndices.visRowShapeLayout))
                {
                    MessageBox.Show("No visRowShapeLayout exists");
                }
                else
                {
                    Row sectionRow = section[(short)VisRowIndices.visRowShapeLayout];

                    sectionRow[VisCellIndices.visSLOPermX].FormulaU = shapeLayout.ShapePermeableX;
                    sectionRow[VisCellIndices.visSLOPermY].FormulaU = shapeLayout.ShapePermeableY;
                    sectionRow[VisCellIndices.visSLOFixedCode].FormulaU = shapeLayout.ShapeFixedCode;
                    sectionRow[VisCellIndices.visSLOJumpDirX].FormulaU = shapeLayout.ConLineJumpDirX;
                    sectionRow[VisCellIndices.visSLOJumpDirY].FormulaU = shapeLayout.ConLineJumpDirY;
                    sectionRow[VisCellIndices.visSLOJumpCode].FormulaU = shapeLayout.ConLineJumpCode;
                    sectionRow[VisCellIndices.visSLOPlaceFlip].FormulaU = shapeLayout.ShapePlaceFlip;
                    sectionRow[VisCellIndices.visSLOPlaceStyle].FormulaU = shapeLayout.ShapePlaceStyle;
                    sectionRow[VisCellIndices.visSLOPlowCode].FormulaU = shapeLayout.ShapePlowCode;
                    sectionRow[VisCellIndices.visSLOJumpStyle].FormulaU = shapeLayout.ConLineJumpStyle;
                    sectionRow[VisCellIndices.visSLOLineRouteExt].FormulaU = shapeLayout.ConLineRouteExt;
                    sectionRow[VisCellIndices.visSLODisplayLevel].FormulaU = shapeLayout.DisplayLevel;
                    sectionRow[VisCellIndices.visSLORouteStyle].FormulaU = shapeLayout.ShapeRouteStyle;
                    sectionRow[VisCellIndices.visSLOConFixedCode].FormulaU = shapeLayout.ConFixedCode;
                    sectionRow[VisCellIndices.visSLOSplit].FormulaU = shapeLayout.ShapeSplit;
                    sectionRow[VisCellIndices.visSLOSplittable].FormulaU = shapeLayout.ShapeSplittable;
                    sectionRow[VisCellIndices.visSLORelationships].FormulaU = shapeLayout.Relationships;
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }
    }
}
