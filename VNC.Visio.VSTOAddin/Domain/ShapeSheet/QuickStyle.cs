using System;

using Microsoft.Office.Interop.Visio;

namespace VNC.Visio.VSTOAddIn.Domain
{
    public class QuickStyle
    {
        public string QuickStyleLineMatrix { get; set; }
        public string QuickStyleLineColor { get; set; }
        public string QuickStyleFontColor { get; set; }
        public string QuickStyleVariation { get; set; }
        public string QuickStyleFillMatrix { get; set; }
        public string QuickStyleFontMatrix { get; set; }
        public string QuickStyleEffectsMatrix { get; set; }
        public string QuickStyleShadowColor { get; set; }
        public string QuickStyleType { get; set; }

        public static QuickStyle GetRow(Shape shape)
        {
            QuickStyle row = new QuickStyle();

            Section section = shape.Section[(short)VisSectionIndices.visSectionObject];
            Row sectionRow = section[(short)VisRowIndices.visRowQuickStyleProperties];

            row.QuickStyleLineMatrix = sectionRow[VisCellIndices.visQuickStyleLineMatrix].FormulaU;
            row.QuickStyleLineColor = sectionRow[VisCellIndices.visQuickStyleLineColor].FormulaU;
            row.QuickStyleFontColor = sectionRow[VisCellIndices.visQuickStyleFontColor].FormulaU;
            row.QuickStyleVariation = sectionRow[VisCellIndices.visQuickStyleVariation].FormulaU;
            row.QuickStyleFillMatrix = sectionRow[VisCellIndices.visQuickStyleFillMatrix].FormulaU;
            row.QuickStyleFontMatrix = sectionRow[VisCellIndices.visQuickStyleFontMatrix].FormulaU;
            row.QuickStyleEffectsMatrix = sectionRow[VisCellIndices.visQuickStyleEffectsMatrix].FormulaU;
            row.QuickStyleShadowColor = sectionRow[VisCellIndices.visQuickStyleShadowColor].FormulaU;
            row.QuickStyleType = sectionRow[VisCellIndices.visQuickStyleType].FormulaU;

            return row;
        }

        public static void Set_QuickStyle_Section(Shape shape, QuickStyle quickStyle)
        {
            try
            {
                Section section = shape.Section[(short)VisSectionIndices.visSectionObject];
                Row sectionRow = section[(short)VisRowIndices.visRowQuickStyleProperties];

                sectionRow[VisCellIndices.visQuickStyleLineMatrix].FormulaU = quickStyle.QuickStyleLineMatrix;
                sectionRow[VisCellIndices.visQuickStyleLineColor].FormulaU = quickStyle.QuickStyleLineColor;
                sectionRow[VisCellIndices.visQuickStyleFontColor].FormulaU = quickStyle.QuickStyleFontColor;
                sectionRow[VisCellIndices.visQuickStyleVariation].FormulaU = quickStyle.QuickStyleVariation;
                sectionRow[VisCellIndices.visQuickStyleFillMatrix].FormulaU = quickStyle.QuickStyleFillMatrix;
                sectionRow[VisCellIndices.visQuickStyleFontMatrix].FormulaU = quickStyle.QuickStyleFontMatrix;
                sectionRow[VisCellIndices.visQuickStyleEffectsMatrix].FormulaU = quickStyle.QuickStyleEffectsMatrix;
                sectionRow[VisCellIndices.visQuickStyleShadowColor].FormulaU = quickStyle.QuickStyleShadowColor;
                sectionRow[VisCellIndices.visQuickStyleType].FormulaU = quickStyle.QuickStyleType;
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }
    }
}
