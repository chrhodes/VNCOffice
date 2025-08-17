using System;
using System.Windows;


using Microsoft.Office.Interop.Visio;

namespace VNC.Visio.VSTOAddIn.Domain
{
    public class ThemePropertiesRow
    {
        public string ConnectorSchemeIndex { get; set; }
        public string EffectSchemeIndex { get; set; }
        public string ColorSchemeIndex { get; set; }
        public string FontSchemeIndex { get; set; }
        public string ThemeIndex { get; set; }
        public string VariationColorIndex { get; set; }
        public string VariationStyleIndex { get; set; }
        public string EmbellishmentIndex { get; set; }

        public static ThemePropertiesRow GetRow(Shape shape)
        {
            ThemePropertiesRow row = new ThemePropertiesRow();

            Section section = shape.Section[(short)VisSectionIndices.visSectionObject];

            if (Helpers.RowExists(shape, VisSectionIndices.visSectionObject, VisRowIndices.visRowThemeProperties))
            {
                Row sectionRow = section[(short)VisRowIndices.visRowThemeProperties];

                row.ConnectorSchemeIndex = sectionRow[VisCellIndices.visConnectorSchemeIndex].FormulaU;
                row.EffectSchemeIndex = sectionRow[VisCellIndices.visEffectSchemeIndex].FormulaU;
                row.ColorSchemeIndex = sectionRow[VisCellIndices.visColorSchemeIndex].FormulaU;
                row.FontSchemeIndex = sectionRow[VisCellIndices.visFontSchemeIndex].FormulaU;
                row.ThemeIndex = sectionRow[VisCellIndices.visThemeIndex].FormulaU;
                row.VariationColorIndex = sectionRow[VisCellIndices.visVariationColorIndex].FormulaU;
                row.VariationStyleIndex = sectionRow[VisCellIndices.visVariationStyleIndex].FormulaU;
                row.EmbellishmentIndex = sectionRow[VisCellIndices.visEmbellishmentIndex].FormulaU;
            }
            else
            {
                MessageBox.Show("No visRowThemeProperties exists");
            }

            return row;
        }

        public static void SetRow(Shape shape, ThemePropertiesRow themeProperties)
        {
            try
            {
                Section section = shape.Section[(short)VisSectionIndices.visSectionObject];

                if (Helpers.RowExists(shape, VisSectionIndices.visSectionObject, VisRowIndices.visRowThemeProperties))
                {
                    Row sectionRow = section[(short)VisRowIndices.visRowThemeProperties];

                    sectionRow[VisCellIndices.visConnectorSchemeIndex].FormulaU = themeProperties.ConnectorSchemeIndex;
                    sectionRow[VisCellIndices.visEffectSchemeIndex].FormulaU = themeProperties.EffectSchemeIndex;
                    sectionRow[VisCellIndices.visColorSchemeIndex].FormulaU = themeProperties.ColorSchemeIndex;
                    sectionRow[VisCellIndices.visFontSchemeIndex].FormulaU = themeProperties.FontSchemeIndex;
                    sectionRow[VisCellIndices.visThemeIndex].FormulaU = themeProperties.ThemeIndex;
                    sectionRow[VisCellIndices.visVariationColorIndex].FormulaU = themeProperties.VariationColorIndex;
                    sectionRow[VisCellIndices.visVariationStyleIndex].FormulaU = themeProperties.VariationStyleIndex;
                    sectionRow[VisCellIndices.visEmbellishmentIndex].FormulaU = themeProperties.EmbellishmentIndex;
                }
                else
                {
                    MessageBox.Show("No visRowThemeProperties exists");
                }
            }
            catch (Exception ex)
            {
                Log.ERROR(ex, Common.LOG_CATEGORY);
            }
        }
    }
}
