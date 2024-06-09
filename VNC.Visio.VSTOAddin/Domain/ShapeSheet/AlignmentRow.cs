using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop.Visio;

namespace VNC.Visio.VSTOAddIn.Domain
{
    public class AlignmentRow
    {
        // TODO(crhodes)
        // populate

        public static AlignmentRow GetRow(Shape shape)
        {
            AlignmentRow row = new AlignmentRow();

            Section section = shape.Section[(short)VisSectionIndices.visSectionObject];
            Row sectionRow = section[(short)VisRowIndices.visRowAlign];

            //row.BevelTopType = sectionRow[VisCellIndices.visBevelTopType].FormulaU;
            //row.BevelTopWidth = sectionRow[VisCellIndices.visBevelTopWidth].FormulaU;
            //row.BevelTopHeight = sectionRow[VisCellIndices.visBevelTopHeight].FormulaU;
            //row.BevelBottomType = sectionRow[VisCellIndices.visBevelBottomType].FormulaU;
            //row.BevelBottomWidth = sectionRow[VisCellIndices.visBevelBottomWidth].FormulaU;
            //row.BevelBottomHeight = sectionRow[VisCellIndices.visBevelBottomHeight].FormulaU;
            //row.BevelDepthColor = sectionRow[VisCellIndices.visBevelDepthColor].FormulaU;
            //row.BevelDepthSize = sectionRow[VisCellIndices.visBevelDepthSize].FormulaU;
            //row.BevelContourColor = sectionRow[VisCellIndices.visBevelContourColor].FormulaU;
            //row.BevelContourSize = sectionRow[VisCellIndices.visBevelContourSize].FormulaU;
            //row.BevelMaterialType = sectionRow[VisCellIndices.visBevelMaterialType].FormulaU;
            //row.BevelLightingType = sectionRow[VisCellIndices.visBevelLightingType].FormulaU;
            //row.BevelLightingAngle = sectionRow[VisCellIndices.visBevelLightingAngle].FormulaU;

            return row;
        }

        public static void SetRow(Shape shape, AlignmentRow alignmentRow)
        {
            // TODO(crhodes)
            // Implement

            //    try
            //    {
            //        Section section = shape.Section[(short)VisSectionIndices.visSectionObject];
            //        Row sectionRow = section[(short)VisRowIndices.visRowMisc];

            //        sectionRow[VisCellIndices.visBegTrigger].FormulaU = glueInfo.BegTrigger;
            //        sectionRow[VisCellIndices.visEndTrigger].FormulaU = glueInfo.EndTrigger;
            //        sectionRow[VisCellIndices.visGlueType].FormulaU = glueInfo.GlueType;
            //        sectionRow[VisCellIndices.visWalkPref].FormulaU = glueInfo.WalkPreference;
            //    }
            //    catch (Exception ex)
            //    {
            //        Log.Error(ex, Common.LOG_CATEGORY);
            //    }
        }
    }
}
