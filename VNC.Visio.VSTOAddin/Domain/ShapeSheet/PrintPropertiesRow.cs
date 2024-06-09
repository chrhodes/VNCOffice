using System;

using Microsoft.Office.Interop.Visio;

namespace VNC.Visio.VSTOAddIn.Domain
{
    public class PrintPropertiesRow
    {
        public string PageLeftMargin { get; set; }
        public string PageTopMargin { get; set; }
        public string PageRightMargin { get; set; }
        public string PageBottomMargin { get; set; }
        public string ScaleX { get; set; }
        public string ScaleY { get; set; }
        public string PagesX { get; set; }
        public string PagesY { get; set; }
        public string CenterX { get; set; }
        public string CenterY { get; set; }
        public string OnPage { get; set; }
        public string PrintGrid { get; set; }
        public string PrintPageOrientation { get; set; }
        public string PaperKind { get; set; }
        public string PaperSource { get; set; }

        public static PrintPropertiesRow Get_PrintProperties(Shape shape)
        {
            PrintPropertiesRow row = new PrintPropertiesRow();

            Section section = shape.Section[(short)VisSectionIndices.visSectionObject];
            Row sectionRow = section[(short)VisRowIndices.visRowPrintProperties];

            row.PageLeftMargin = sectionRow[VisCellIndices.visPrintPropertiesLeftMargin].FormulaU;
            row.PageTopMargin = sectionRow[VisCellIndices.visPrintPropertiesTopMargin].FormulaU;
            row.PageRightMargin = sectionRow[VisCellIndices.visPrintPropertiesRightMargin].FormulaU;
            row.PageBottomMargin = sectionRow[VisCellIndices.visPrintPropertiesBottomMargin].FormulaU;
            row.ScaleX = sectionRow[VisCellIndices.visPrintPropertiesScaleX].FormulaU;
            row.ScaleY = sectionRow[VisCellIndices.visPrintPropertiesScaleY].FormulaU;
            row.PagesX = sectionRow[VisCellIndices.visPrintPropertiesPagesX].FormulaU;
            row.PagesY = sectionRow[VisCellIndices.visPrintPropertiesPagesY].FormulaU;
            row.CenterX = sectionRow[VisCellIndices.visPrintPropertiesCenterX].FormulaU;
            row.CenterY = sectionRow[VisCellIndices.visPrintPropertiesCenterY].FormulaU;
            row.OnPage = sectionRow[VisCellIndices.visPrintPropertiesOnPage].FormulaU;
            row.PrintGrid = sectionRow[VisCellIndices.visPrintPropertiesPrintGrid].FormulaU;
            row.PrintPageOrientation = sectionRow[VisCellIndices.visPrintPropertiesPageOrientation].FormulaU;
            row.PaperKind = sectionRow[VisCellIndices.visPrintPropertiesPaperKind].FormulaU;
            row.PaperSource = sectionRow[VisCellIndices.visPrintPropertiesPaperSource].FormulaU;

            return row;
        }

        public static void Set_PrintProperties_Section(Shape shape, PrintPropertiesRow printProperties)
        {
            try
            {
                Section section = shape.Section[(short)VisSectionIndices.visSectionObject];
                Row sectionRow = section[(short)VisRowIndices.visRowPrintProperties];

                sectionRow[VisCellIndices.visPrintPropertiesLeftMargin].FormulaU = printProperties.PageLeftMargin;
                sectionRow[VisCellIndices.visPrintPropertiesTopMargin].FormulaU = printProperties.PageTopMargin;
                sectionRow[VisCellIndices.visPrintPropertiesRightMargin].FormulaU = printProperties.PageRightMargin;
                sectionRow[VisCellIndices.visPrintPropertiesBottomMargin].FormulaU = printProperties.PageBottomMargin;
                sectionRow[VisCellIndices.visPrintPropertiesScaleX].FormulaU = printProperties.ScaleX;
                sectionRow[VisCellIndices.visPrintPropertiesScaleY].FormulaU = printProperties.ScaleY;
                sectionRow[VisCellIndices.visPrintPropertiesPagesX].FormulaU = printProperties.PagesX;
                sectionRow[VisCellIndices.visPrintPropertiesPagesY].FormulaU = printProperties.PagesY;
                sectionRow[VisCellIndices.visPrintPropertiesCenterX].FormulaU = printProperties.CenterX;
                sectionRow[VisCellIndices.visPrintPropertiesCenterY].FormulaU = printProperties.CenterY;
                sectionRow[VisCellIndices.visPrintPropertiesOnPage].FormulaU = printProperties.OnPage;
                sectionRow[VisCellIndices.visPrintPropertiesPrintGrid].FormulaU = printProperties.PrintGrid;
                sectionRow[VisCellIndices.visPrintPropertiesPageOrientation].FormulaU = printProperties.PrintPageOrientation;
                sectionRow[VisCellIndices.visPrintPropertiesPaperKind].FormulaU = printProperties.PaperKind;
                sectionRow[VisCellIndices.visPrintPropertiesPaperSource].FormulaU = printProperties.PaperSource;
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }
    }
}
