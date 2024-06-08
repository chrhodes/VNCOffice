using System;

using Microsoft.Office.Interop.Visio;

namespace VNC.Visio.VSTOAddIn.Domain
{
    public class LineFormat
    {
        public string LinePattern { get; set; }
        public string LineWeight { get; set; }
        public string LineColor { get; set; }
        public string LineCap { get; set; }
        public string BeginArrow { get; set; }
        public string EndArrow { get; set; }
        public string LineColorTrans { get; set; }
        public string CompoundType { get; set; }
        public string BeginArrowSize { get; set; }
        public string EndArrowSize { get; set; }
        public string Rounding { get; set; }

        public static LineFormat Get_LineFormat(Shape shape)
        {
            LineFormat row = new LineFormat();

            Section section = shape.Section[(short)VisSectionIndices.visSectionObject];
            Row sectionRow = section[(short)VisRowIndices.visRowLine];

            row.LinePattern = sectionRow[VisCellIndices.visLinePattern].FormulaU;
            row.LineWeight = sectionRow[VisCellIndices.visLineWeight].FormulaU;
            row.LineColor = sectionRow[VisCellIndices.visLineColor].FormulaU;
            row.LineCap = sectionRow[VisCellIndices.visLineEndCap].FormulaU;
            row.BeginArrow = sectionRow[VisCellIndices.visLineBeginArrow].FormulaU;
            row.EndArrow = sectionRow[VisCellIndices.visLineEndArrow].FormulaU;
            row.LineColorTrans = sectionRow[VisCellIndices.visLineColorTrans].FormulaU;
            row.CompoundType = sectionRow[VisCellIndices.visCompoundType].FormulaU;
            row.BeginArrowSize = sectionRow[VisCellIndices.visLineBeginArrowSize].FormulaU;
            row.EndArrowSize = sectionRow[VisCellIndices.visLineEndArrowSize].FormulaU;
            row.Rounding = sectionRow[VisCellIndices.visLineRounding].FormulaU;

            return row;
        }

        public static void Set_LineFormat_Section(Shape shape, LineFormat lineFormat)
        {
            try
            {
                Section section = shape.Section[(short)VisSectionIndices.visSectionObject];
                Row sectionRow = section[(short)VisRowIndices.visRowLine];

                sectionRow[VisCellIndices.visLinePattern].FormulaU = lineFormat.LinePattern;
                sectionRow[VisCellIndices.visLineWeight].FormulaU = lineFormat.LineWeight;
                sectionRow[VisCellIndices.visLineColor].FormulaU = lineFormat.LineColor;
                sectionRow[VisCellIndices.visLineEndCap].FormulaU = lineFormat.LineCap;
                sectionRow[VisCellIndices.visLineBeginArrow].FormulaU = lineFormat.BeginArrow;
                sectionRow[VisCellIndices.visLineEndArrow].FormulaU = lineFormat.EndArrow;
                sectionRow[VisCellIndices.visLineColorTrans].FormulaU = lineFormat.LineColorTrans;
                sectionRow[VisCellIndices.visCompoundType].FormulaU = lineFormat.CompoundType;
                sectionRow[VisCellIndices.visLineBeginArrowSize].FormulaU = lineFormat.BeginArrowSize;
                sectionRow[VisCellIndices.visLineEndArrowSize].FormulaU = lineFormat.EndArrowSize;
                sectionRow[VisCellIndices.visLineRounding].FormulaU = lineFormat.Rounding;
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }


    }
}
