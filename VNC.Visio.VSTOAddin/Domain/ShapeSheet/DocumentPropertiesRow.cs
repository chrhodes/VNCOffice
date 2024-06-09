using System;

using Microsoft.Office.Interop.Visio;

namespace VNC.Visio.VSTOAddIn.Domain
{
    /// <summary>
    /// Document Properties - VisSectionObject - Optional
    /// </summary>
    public class DocumentPropertiesRow
    {
        public string PreviewQuality { get; set; }
        public string OutputFormat { get; set; }
        public string PreviewScope { get; set; }
        public string LockPreview { get; set; }
        public string AddMarkup { get; set; }
        public string ViewMarkup { get; set; }
        public string DocLangID { get; set; }
        public string DocLockReplace { get; set; }
        public string NoCoauth { get; set; }
        public string DocLockDuplicatePage { get; set; }

        public static DocumentPropertiesRow GetRow(Shape shape)
        {
            DocumentPropertiesRow row = new DocumentPropertiesRow();

            Section section = shape.Section[(short)VisSectionIndices.visSectionObject];
            Row sectionRow = section[(short)VisRowIndices.visRowDoc];

            row.PreviewQuality = sectionRow[VisCellIndices.visDocPreviewQuality].FormulaU;
            row.OutputFormat = sectionRow[VisCellIndices.visDocOutputFormat].FormulaU;
            row.PreviewScope = sectionRow[VisCellIndices.visDocPreviewScope].FormulaU;
            row.LockPreview = sectionRow[VisCellIndices.visDocLockPreview].FormulaU;
            row.AddMarkup = sectionRow[VisCellIndices.visDocAddMarkup].FormulaU;
            row.ViewMarkup = sectionRow[VisCellIndices.visDocViewMarkup].FormulaU;
            row.DocLangID = sectionRow[VisCellIndices.visDocLangID].FormulaU;
            row.DocLockReplace = sectionRow[VisCellIndices.visDocLockReplace].FormulaU;
            row.NoCoauth = sectionRow[VisCellIndices.visDocNoCoauth].FormulaU;
            row.DocLockDuplicatePage = sectionRow[VisCellIndices.visDocLockDuplicatePage].FormulaU;

            return row;
        }

        public static void SetRow(Shape shape, DocumentPropertiesRow documentProperties)
        {
            try
            {
                Section section = shape.Section[(short)VisSectionIndices.visSectionObject];
                Row sectionRow = section[(short)VisRowIndices.visRowDoc];

                sectionRow[VisCellIndices.visDocPreviewQuality].FormulaU = documentProperties.PreviewQuality;
                sectionRow[VisCellIndices.visDocOutputFormat].FormulaU = documentProperties.OutputFormat;
                sectionRow[VisCellIndices.visDocPreviewScope].FormulaU = documentProperties.PreviewScope;
                sectionRow[VisCellIndices.visDocLockPreview].FormulaU = documentProperties.LockPreview;
                sectionRow[VisCellIndices.visDocAddMarkup].FormulaU = documentProperties.AddMarkup;
                sectionRow[VisCellIndices.visDocViewMarkup].FormulaU = documentProperties.ViewMarkup;
                sectionRow[VisCellIndices.visDocLangID].FormulaU = documentProperties.DocLangID;
                sectionRow[VisCellIndices.visDocLockReplace].FormulaU = documentProperties.DocLockReplace;
                sectionRow[VisCellIndices.visDocNoCoauth].FormulaU = documentProperties.NoCoauth;
                sectionRow[VisCellIndices.visDocLockDuplicatePage].FormulaU = documentProperties.DocLockDuplicatePage;
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }
    }
}
