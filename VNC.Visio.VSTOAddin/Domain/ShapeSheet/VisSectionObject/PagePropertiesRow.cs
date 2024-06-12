using System;
using System.Windows;


using Microsoft.Office.Interop.Visio;

namespace VNC.Visio.VSTOAddIn.Domain
{
    public class PagePropertiesRow
    {
        public string PageWidth { get; set; }
        public string PageHeight { get; set; }
        public string PageScale { get; set; }
        public string DrawingScale { get; set; }
        public string DrawingSizeType { get; set; }
        public string DrawingScaleType { get; set; }
        public string DrawingResizeType { get; set; }
        public string InhibitSnap { get; set; }
        public string UIVisibility { get; set; }
        public string PageLockReplace { get; set; }
        public string PageLockDuplicate { get; set; }

        public static PagePropertiesRow GetRow(Shape shape)
        {
            PagePropertiesRow row = new PagePropertiesRow();

            Section section = shape.Section[(short)VisSectionIndices.visSectionObject];

            if (Helpers.RowExists(shape, VisSectionIndices.visSectionObject, VisRowIndices.visRowPage))
            {
                Row sectionRow = section[(short)VisRowIndices.visRowPage];

                row.PageWidth = sectionRow[VisCellIndices.visPageWidth].FormulaU;
                row.PageHeight = sectionRow[VisCellIndices.visPageHeight].FormulaU;
                row.PageScale = sectionRow[VisCellIndices.visPageScale].FormulaU;
                row.DrawingScale = sectionRow[VisCellIndices.visPageDrawingScale].FormulaU;
                row.DrawingSizeType = sectionRow[VisCellIndices.visPageDrawSizeType].FormulaU;
                row.DrawingResizeType = sectionRow[VisCellIndices.visPageDrawResizeType].FormulaU;
                row.DrawingScaleType = sectionRow[VisCellIndices.visPageDrawScaleType].FormulaU;
                row.InhibitSnap = sectionRow[VisCellIndices.visPageInhibitSnap].FormulaU;
                row.UIVisibility = sectionRow[VisCellIndices.visPageUIVisibility].FormulaU;
                row.PageLockReplace = sectionRow[VisCellIndices.visPageLockReplace].FormulaU;
                row.PageLockDuplicate = sectionRow[VisCellIndices.visPageLockDuplicate].FormulaU;
            }
            else
            {
                MessageBox.Show("No visRowPage exists");
            }

            return row;
        }

        public static void SetRow(Shape shape, PagePropertiesRow pageProperties)
        {
            try
            {
                Section section = shape.Section[(short)VisSectionIndices.visSectionObject];

                if (Helpers.RowExists(shape, VisSectionIndices.visSectionObject, VisRowIndices.visRowPage))
                {
                    Row sectionRow = section[(short)VisRowIndices.visRowPage];

                    sectionRow[VisCellIndices.visPageWidth].FormulaU = pageProperties.PageWidth;
                    sectionRow[VisCellIndices.visPageHeight].FormulaU = pageProperties.PageHeight;
                    sectionRow[VisCellIndices.visPageScale].FormulaU = pageProperties.PageScale;
                    sectionRow[VisCellIndices.visPageDrawingScale].FormulaU = pageProperties.DrawingScale;
                    sectionRow[VisCellIndices.visPageDrawSizeType].FormulaU = pageProperties.DrawingSizeType;
                    sectionRow[VisCellIndices.visPageDrawScaleType].FormulaU = pageProperties.DrawingScaleType;
                    sectionRow[VisCellIndices.visPageInhibitSnap].FormulaU = pageProperties.InhibitSnap;
                    sectionRow[VisCellIndices.visPageUIVisibility].FormulaU = pageProperties.UIVisibility;
                    sectionRow[VisCellIndices.visPageLockReplace].FormulaU = pageProperties.PageLockReplace;
                    sectionRow[VisCellIndices.visPageLockDuplicate].FormulaU = pageProperties.PageLockDuplicate;
                }
                else
                {
                    MessageBox.Show("No visRowPage exists");
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }
    }
}
