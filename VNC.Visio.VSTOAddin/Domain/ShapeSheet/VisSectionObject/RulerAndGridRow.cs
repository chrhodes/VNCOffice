using System;
using System.Windows;


using Microsoft.Office.Interop.Visio;

namespace VNC.Visio.VSTOAddIn.Domain
{
    public class RulerAndGridRow
    {
        public string XRulerOrigin { get; set; }
        public string YRulerOrigin { get; set; }
        public string XRulerDensity { get; set; }
        public string YRulerDensity { get; set; }
        public string XGridOrigin { get; set; }
        public string YGridOrigin { get; set; }
        public string XGridDensity { get; set; }
        public string YGridDensity { get; set; }
        public string XGridSpacing { get; set; }
        public string YGridSpacing { get; set; }

        public static RulerAndGridRow GetRow(Shape shape)
        {
            RulerAndGridRow row = new RulerAndGridRow();

            Section section = shape.Section[(short)VisSectionIndices.visSectionObject];

            if (Helpers.RowExists(shape, VisSectionIndices.visSectionObject, VisRowIndices.visRowRulerGrid))
            {
                MessageBox.Show("No visRowRulerGrid exists");
            }
            else
            {
                Row sectionRow = section[(short)VisRowIndices.visRowRulerGrid];

                row.XRulerOrigin = sectionRow[VisCellIndices.visXRulerOrigin].FormulaU;
                row.YRulerOrigin = sectionRow[VisCellIndices.visYRulerOrigin].FormulaU;
                row.XRulerDensity = sectionRow[VisCellIndices.visXRulerDensity].FormulaU;
                row.YRulerDensity = sectionRow[VisCellIndices.visYRulerDensity].FormulaU;
                row.XGridOrigin = sectionRow[VisCellIndices.visXGridOrigin].FormulaU;
                row.YGridOrigin = sectionRow[VisCellIndices.visYGridOrigin].FormulaU;
                row.XGridDensity = sectionRow[VisCellIndices.visXGridDensity].FormulaU;
                row.YGridDensity = sectionRow[VisCellIndices.visYGridDensity].FormulaU;
                row.XGridSpacing = sectionRow[VisCellIndices.visXGridSpacing].FormulaU;
                row.YGridSpacing = sectionRow[VisCellIndices.visYGridSpacing].FormulaU;
            }

            return row;
        }

        public static void SetRow(Shape shape, RulerAndGridRow rulerAndGrid)
        {
            try
            {
                Section section = shape.Section[(short)VisSectionIndices.visSectionObject];

                if (Helpers.RowExists(shape, VisSectionIndices.visSectionObject, VisRowIndices.visRowRulerGrid))
                {
                    MessageBox.Show("No visRowRulerGrid exists");
                }
                else
                {
                    Row sectionRow = section[(short)VisRowIndices.visRowRulerGrid];

                sectionRow[VisCellIndices.visXRulerOrigin].FormulaU = rulerAndGrid.XRulerOrigin;
                sectionRow[VisCellIndices.visYRulerOrigin].FormulaU = rulerAndGrid.YRulerOrigin;
                sectionRow[VisCellIndices.visXRulerDensity].FormulaU = rulerAndGrid.XRulerDensity;
                sectionRow[VisCellIndices.visYRulerDensity].FormulaU = rulerAndGrid.YRulerDensity;
                sectionRow[VisCellIndices.visXGridOrigin].FormulaU = rulerAndGrid.XGridOrigin;
                sectionRow[VisCellIndices.visYGridOrigin].FormulaU = rulerAndGrid.YGridOrigin;
                sectionRow[VisCellIndices.visXGridDensity].FormulaU = rulerAndGrid.XGridDensity;
                sectionRow[VisCellIndices.visYGridDensity].FormulaU = rulerAndGrid.YGridDensity;
                sectionRow[VisCellIndices.visXGridSpacing].FormulaU = rulerAndGrid.XGridSpacing;
                sectionRow[VisCellIndices.visYGridSpacing].FormulaU = rulerAndGrid.YGridSpacing;

                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }
    }
}
