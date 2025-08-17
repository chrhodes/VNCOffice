using System;
using System.Windows;

using Microsoft.Office.Interop.Visio;

namespace VNC.Visio.VSTOAddIn.Domain
{
    public class BevelPropertiesRow
    {
        public string BevelTopType { get; set; }
        public string BevelTopWidth { get; set; }
        public string BevelTopHeight { get; set; }
        public string BevelBottomType { get; set; }
        public string BevelBottomWidth { get; set; }
        public string BevelBottomHeight { get; set; }
        public string BevelDepthColor { get; set; }
        public string BevelDepthSize { get; set; }
        public string BevelContourColor { get; set; }
        public string BevelContourSize { get; set; }
        public string BevelMaterialType { get; set; }
        public string BevelLightingType { get; set; }
        public string BevelLightingAngle { get; set; }

        public static BevelPropertiesRow GetRow(Shape shape)
        {
            BevelPropertiesRow row = new BevelPropertiesRow();

            Section section = shape.Section[(short)VisSectionIndices.visSectionObject];

            if (Helpers.RowExists(shape, VisSectionIndices.visSectionObject, VisRowIndices.visRowBevelProperties))
            {
                Row sectionRow = section[(short)VisRowIndices.visRowBevelProperties];

                row.BevelTopType = sectionRow[VisCellIndices.visBevelTopType].FormulaU;
                row.BevelTopWidth = sectionRow[VisCellIndices.visBevelTopWidth].FormulaU;
                row.BevelTopHeight = sectionRow[VisCellIndices.visBevelTopHeight].FormulaU;
                row.BevelBottomType = sectionRow[VisCellIndices.visBevelBottomType].FormulaU;
                row.BevelBottomWidth = sectionRow[VisCellIndices.visBevelBottomWidth].FormulaU;
                row.BevelBottomHeight = sectionRow[VisCellIndices.visBevelBottomHeight].FormulaU;
                row.BevelDepthColor = sectionRow[VisCellIndices.visBevelDepthColor].FormulaU;
                row.BevelDepthSize = sectionRow[VisCellIndices.visBevelDepthSize].FormulaU;
                row.BevelContourColor = sectionRow[VisCellIndices.visBevelContourColor].FormulaU;
                row.BevelContourSize = sectionRow[VisCellIndices.visBevelContourSize].FormulaU;
                row.BevelMaterialType = sectionRow[VisCellIndices.visBevelMaterialType].FormulaU;
                row.BevelLightingType = sectionRow[VisCellIndices.visBevelLightingType].FormulaU;
                row.BevelLightingAngle = sectionRow[VisCellIndices.visBevelLightingAngle].FormulaU;
            }
            else
            {
                MessageBox.Show("No visRowBevelProperties exists");
            }

            return row;
        }

        public static void SetRow(Shape shape, BevelPropertiesRow bevelProperties)
        {
            try
            {
                Section section = shape.Section[(short)VisSectionIndices.visSectionObject];

                if (Helpers.RowExists(shape, VisSectionIndices.visSectionObject, VisRowIndices.visRowBevelProperties))
                {
                    Row sectionRow = section[(short)VisRowIndices.visRowBevelProperties];

                    sectionRow[VisCellIndices.visBevelTopType].FormulaU = bevelProperties.BevelTopType;
                    sectionRow[VisCellIndices.visBevelTopWidth].FormulaU = bevelProperties.BevelTopWidth;
                    sectionRow[VisCellIndices.visBevelTopHeight].FormulaU = bevelProperties.BevelTopHeight;
                    sectionRow[VisCellIndices.visBevelBottomType].FormulaU = bevelProperties.BevelBottomType;
                    sectionRow[VisCellIndices.visBevelBottomWidth].FormulaU = bevelProperties.BevelBottomWidth;
                    sectionRow[VisCellIndices.visBevelBottomHeight].FormulaU = bevelProperties.BevelBottomHeight;
                    sectionRow[VisCellIndices.visBevelDepthColor].FormulaU = bevelProperties.BevelDepthColor;
                    sectionRow[VisCellIndices.visBevelDepthSize].FormulaU = bevelProperties.BevelDepthSize;
                    sectionRow[VisCellIndices.visBevelContourColor].FormulaU = bevelProperties.BevelContourColor;
                    sectionRow[VisCellIndices.visBevelContourSize].FormulaU = bevelProperties.BevelContourSize;
                    sectionRow[VisCellIndices.visBevelMaterialType].FormulaU = bevelProperties.BevelMaterialType;
                    sectionRow[VisCellIndices.visBevelLightingType].FormulaU = bevelProperties.BevelLightingType;
                    sectionRow[VisCellIndices.visBevelLightingAngle].FormulaU = bevelProperties.BevelLightingAngle;
                }
                else
                {
                    MessageBox.Show("No visRowBevelProperties exists");
                }
            }
            catch (Exception ex)
            {
                Log.ERROR(ex, Common.LOG_CATEGORY);
            }
        }
    }
}
