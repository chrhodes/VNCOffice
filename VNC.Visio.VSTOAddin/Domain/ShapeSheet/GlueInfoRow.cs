using System;

using Microsoft.Office.Interop.Visio;

namespace VNC.Visio.VSTOAddIn.Domain
{
    public class GlueInfoRow
    {
        public string BegTrigger { get; set; }
        public string EndTrigger { get; set; }
        public string GlueType { get; set; }
        public string WalkPreference { get; set; }

        public static GlueInfoRow Get_GlueInfo(Shape shape)
        {
            GlueInfoRow row = new GlueInfoRow();

            Section section = shape.Section[(short)VisSectionIndices.visSectionObject];
            Row sectionRow = section[(short)VisRowIndices.visRowMisc];

            row.BegTrigger = sectionRow[VisCellIndices.visBegTrigger].FormulaU;
            row.EndTrigger = sectionRow[VisCellIndices.visEndTrigger].FormulaU;
            row.GlueType = sectionRow[VisCellIndices.visGlueType].FormulaU;
            row.WalkPreference = sectionRow[VisCellIndices.visWalkPref].FormulaU;

            return row;
        }

        public static void Set_GlueInfo_Section(Shape shape, GlueInfoRow glueInfo)
        {
            try
            {
                Section section = shape.Section[(short)VisSectionIndices.visSectionObject];
                Row sectionRow = section[(short)VisRowIndices.visRowMisc];

                sectionRow[VisCellIndices.visBegTrigger].FormulaU = glueInfo.BegTrigger;
                sectionRow[VisCellIndices.visEndTrigger].FormulaU = glueInfo.EndTrigger;
                sectionRow[VisCellIndices.visGlueType].FormulaU = glueInfo.GlueType;
                sectionRow[VisCellIndices.visWalkPref].FormulaU = glueInfo.WalkPreference;
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }
    }
}
