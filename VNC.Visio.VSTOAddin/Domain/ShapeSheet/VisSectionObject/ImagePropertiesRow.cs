using System;
using System.Windows;

using Microsoft.Office.Interop.Visio;

namespace VNC.Visio.VSTOAddIn.Domain
{
    public class ImagePropertiesRow
    {
        public string Contrast { get; set; }
        public string Gamma { get; set; }
        public string Sharpen { get; set; }
        public string Brightness { get; set; }
        public string Blur { get; set; }
        public string Denoise { get; set; }
        public string Transparency { get; set; }

        public static ImagePropertiesRow GetRow(Shape shape)
        {
            ImagePropertiesRow row = new ImagePropertiesRow();

            Section section = shape.Section[(short)VisSectionIndices.visSectionObject];

            if (Helpers.RowExists(shape, VisSectionIndices.visSectionObject, VisRowIndices.visRowImage))
            {
                MessageBox.Show("No visRowImage exists");
            }
            else
            {
                Row sectionRow = section[(short)VisRowIndices.visRowImage];

                row.Contrast = sectionRow[VisCellIndices.visImageContrast].FormulaU;
                row.Gamma = sectionRow[VisCellIndices.visImageGamma].FormulaU;
                row.Sharpen = sectionRow[VisCellIndices.visImageSharpen].FormulaU;
                row.Brightness = sectionRow[VisCellIndices.visImageBrightness].FormulaU;
                row.Blur = sectionRow[VisCellIndices.visImageBlur].FormulaU;
                row.Denoise = sectionRow[VisCellIndices.visImageDenoise].FormulaU;
                row.Transparency = sectionRow[VisCellIndices.visImageTransparency].FormulaU;
            }

            return row;
        }

        public static void SetRow(Shape shape, ImagePropertiesRow imageProperties)
        {
            try
            {
                Section section = shape.Section[(short)VisSectionIndices.visSectionObject];

                if (Helpers.RowExists(shape, VisSectionIndices.visSectionObject, VisRowIndices.visRowImage))
                {
                    MessageBox.Show("No visRowImage exists");
                }
                else
                {
                    Row sectionRow = section[(short)VisRowIndices.visRowImage];

                    sectionRow[VisCellIndices.visImageContrast].FormulaU = imageProperties.Contrast;
                    sectionRow[VisCellIndices.visImageGamma].FormulaU = imageProperties.Gamma;
                    sectionRow[VisCellIndices.visImageSharpen].FormulaU = imageProperties.Sharpen;
                    sectionRow[VisCellIndices.visImageBrightness].FormulaU = imageProperties.Brightness;
                    sectionRow[VisCellIndices.visImageBlur].FormulaU = imageProperties.Blur;
                    sectionRow[VisCellIndices.visImageDenoise].FormulaU = imageProperties.Denoise;
                    imageProperties.Transparency = sectionRow[VisCellIndices.visImageTransparency].FormulaU;
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }
    }
}
