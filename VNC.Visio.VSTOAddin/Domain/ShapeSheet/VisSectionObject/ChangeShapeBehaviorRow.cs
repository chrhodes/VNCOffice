using System;

using Microsoft.Office.Interop.Visio;

namespace VNC.Visio.VSTOAddIn.Domain
{
    public class ChangeShapeBehaviorRow
    {
        public string ReplaceLockShapeData { get; set; }
        public string ReplaceLockText { get; set; }
        public string ReplaceLockFormat { get; set; }
        public string ReplaceCopyCells { get; set; }

        public static ChangeShapeBehaviorRow GetRow(Shape shape)
        {
            ChangeShapeBehaviorRow row = new ChangeShapeBehaviorRow();

            Section section = shape.Section[(short)VisSectionIndices.visSectionObject];
            Row sectionRow = section[(short)VisRowIndices.visRowReplaceBehaviors];

            row.ReplaceLockShapeData = sectionRow[VisCellIndices.visReplaceLockShapeData].FormulaU;
            row.ReplaceLockText = sectionRow[VisCellIndices.visReplaceLockText].FormulaU;
            row.ReplaceLockFormat = sectionRow[VisCellIndices.visReplaceLockFormat].FormulaU;
            row.ReplaceCopyCells = sectionRow[VisCellIndices.visReplaceCopyCells].FormulaU;

            return row;
        }

        public static void SetRow(Shape shape, ChangeShapeBehaviorRow changeShapeBehavior)
        {
            try
            {
                Section section = shape.Section[(short)VisSectionIndices.visSectionObject];
                Row sectionRow = section[(short)VisRowIndices.visRowReplaceBehaviors];

                sectionRow[VisCellIndices.visReplaceLockShapeData].FormulaU = changeShapeBehavior.ReplaceLockShapeData;
                sectionRow[VisCellIndices.visReplaceLockText].FormulaU = changeShapeBehavior.ReplaceLockText;
                sectionRow[VisCellIndices.visReplaceLockFormat].FormulaU = changeShapeBehavior.ReplaceLockFormat;
                sectionRow[VisCellIndices.visReplaceCopyCells].FormulaU = changeShapeBehavior.ReplaceCopyCells;
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }
    }
}
