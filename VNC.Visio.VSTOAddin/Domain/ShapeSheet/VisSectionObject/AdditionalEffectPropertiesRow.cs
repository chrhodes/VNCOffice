using System;
using System.Windows;

using Microsoft.Office.Interop.Visio;

namespace VNC.Visio.VSTOAddIn.Domain
{
    public class AdditionalEffectPropertiesRow
    {
        public string ReflectionTrans { get; set; }
        public string ReflectionSize { get; set; }
        public string ReflectionDist { get; set; }
        public string ReflectionBlur { get; set; }
        public string SketchEnabled { get; set; }
        public string SketchSeed { get; set; }
        public string SketchAmount { get; set; }
        public string SketchLineWeight { get; set; }
        public string SketchLineChange { get; set; }
        public string SketchFillChange { get; set; }
        public string GlowColor { get; set; }
        public string GlowColorTrans { get; set; }
        public string GlowSize { get; set; }
        public string SoftEdgesSize { get; set; }

        public static AdditionalEffectPropertiesRow GetRow(Shape shape)
        {
            AdditionalEffectPropertiesRow row = new AdditionalEffectPropertiesRow();

            Section section = shape.Section[(short)VisSectionIndices.visSectionObject];

            if (Helpers.RowExists(shape, VisSectionIndices.visSectionObject, VisRowIndices.visRowOtherEffectProperties))
            {
                Row sectionRow = section[(short)VisRowIndices.visRowOtherEffectProperties];

                row.ReflectionTrans = sectionRow[VisCellIndices.visReflectionTrans].FormulaU;
                row.ReflectionSize = sectionRow[VisCellIndices.visReflectionSize].FormulaU;
                row.ReflectionDist = sectionRow[VisCellIndices.visReflectionDist].FormulaU;
                row.ReflectionBlur = sectionRow[VisCellIndices.visReflectionBlur].FormulaU;
                row.SketchEnabled = sectionRow[VisCellIndices.visSketchEnabled].FormulaU;
                row.SketchSeed = sectionRow[VisCellIndices.visSketchSeed].FormulaU;
                row.SketchAmount = sectionRow[VisCellIndices.visSketchAmount].FormulaU;
                row.SketchLineWeight = sectionRow[VisCellIndices.visSketchLineWeight].FormulaU;
                row.SketchLineChange = sectionRow[VisCellIndices.visSketchLineChange].FormulaU;
                row.SketchFillChange = sectionRow[VisCellIndices.visSketchFillChange].FormulaU;
                row.GlowColor = sectionRow[VisCellIndices.visGlowColor].FormulaU;
                row.GlowColorTrans = sectionRow[VisCellIndices.visGlowColorTrans].FormulaU;
                row.GlowSize = sectionRow[VisCellIndices.visGlowSize].FormulaU;
                row.SoftEdgesSize = sectionRow[VisCellIndices.visSoftEdgesSize].FormulaU;
            }
            else
            {
                MessageBox.Show("No visRowOtherEffectProperties exists");
            }

            return row;
        }

        public static void SetRow(Shape shape, AdditionalEffectPropertiesRow additionalEffectProperties)
        {
            try
            {
                Section section = shape.Section[(short)VisSectionIndices.visSectionObject];

                if (Helpers.RowExists(shape, VisSectionIndices.visSectionObject, VisRowIndices.visRowOtherEffectProperties))
                {
                    Row sectionRow = section[(short)VisRowIndices.visRowOtherEffectProperties];

                    sectionRow[VisCellIndices.visReflectionTrans].FormulaU = additionalEffectProperties.ReflectionTrans;
                    sectionRow[VisCellIndices.visReflectionSize].FormulaU = additionalEffectProperties.ReflectionSize;
                    sectionRow[VisCellIndices.visReflectionDist].FormulaU = additionalEffectProperties.ReflectionDist;
                    sectionRow[VisCellIndices.visReflectionBlur].FormulaU = additionalEffectProperties.ReflectionBlur;
                    sectionRow[VisCellIndices.visSketchEnabled].FormulaU = additionalEffectProperties.SketchEnabled;
                    sectionRow[VisCellIndices.visSketchSeed].FormulaU = additionalEffectProperties.SketchSeed;
                    sectionRow[VisCellIndices.visSketchAmount].FormulaU = additionalEffectProperties.SketchAmount;
                    sectionRow[VisCellIndices.visSketchLineWeight].FormulaU = additionalEffectProperties.SketchLineWeight;
                    sectionRow[VisCellIndices.visSketchLineChange].FormulaU = additionalEffectProperties.SketchLineChange;
                    sectionRow[VisCellIndices.visSketchFillChange].FormulaU = additionalEffectProperties.SketchFillChange;
                    sectionRow[VisCellIndices.visGlowColor].FormulaU = additionalEffectProperties.GlowColor;
                    sectionRow[VisCellIndices.visGlowColorTrans].FormulaU = additionalEffectProperties.GlowColorTrans;
                    sectionRow[VisCellIndices.visGlowSize].FormulaU = additionalEffectProperties.GlowSize;
                    sectionRow[VisCellIndices.visSoftEdgesSize].FormulaU = additionalEffectProperties.SoftEdgesSize;
                }
                else
                {
                    MessageBox.Show("No visRowOtherEffectProperties exists");
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }
    }
}
