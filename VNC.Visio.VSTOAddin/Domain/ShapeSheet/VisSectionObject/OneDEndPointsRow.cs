using System;
using System.Windows;

using Microsoft.Office.Interop.Visio;

namespace VNC.Visio.VSTOAddIn.Domain
{
    public class OneDEndPointsRow
    {
        public string BeginX { get; set; }
        public string BeginY { get; set; }
        public string EndX { get; set; }
        public string EndY { get; set; }


        public static OneDEndPointsRow GetRow(Shape shape)
        {
            OneDEndPointsRow row = new OneDEndPointsRow();

            Section section = shape.Section[(short)VisSectionIndices.visSectionObject];

            if (Helpers.RowExists(shape, VisSectionIndices.visSectionObject, VisRowIndices.visRowXForm1D))
            {
                MessageBox.Show("No visRowXForm1D exists");
            }
            else
            {
                Row sectionRow = section[(short)VisRowIndices.visRowXForm1D];

                row.BeginX = sectionRow[VisCellIndices.vis1DBeginX].FormulaU;
                row.BeginY = sectionRow[VisCellIndices.vis1DBeginY].FormulaU;
                row.EndX = sectionRow[VisCellIndices.vis1DEndX].FormulaU;
                row.EndY = sectionRow[VisCellIndices.vis1DEndY].FormulaU;
            }

            return row;
        }

        public static void SetRow(Shape shape, OneDEndPointsRow oneDEndPoints)
        {
            try
            {
                Section section = shape.Section[(short)VisSectionIndices.visSectionObject];

                if (Helpers.RowExists(shape, VisSectionIndices.visSectionObject, VisRowIndices.visRowXForm1D))
                {
                    MessageBox.Show("No visRowXForm1D exists");
                }
                else
                {
                    Row sectionRow = section[(short)VisRowIndices.visRowXForm1D];

                    sectionRow[VisCellIndices.vis1DBeginX].FormulaU = oneDEndPoints.BeginX;
                    sectionRow[VisCellIndices.vis1DBeginY].FormulaU = oneDEndPoints.BeginY;
                    sectionRow[VisCellIndices.vis1DEndX].FormulaU = oneDEndPoints.EndX;
                    sectionRow[VisCellIndices.vis1DEndY].FormulaU = oneDEndPoints.EndY;
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }
    }
}
