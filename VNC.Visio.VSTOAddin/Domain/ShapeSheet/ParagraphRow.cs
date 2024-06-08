using System;

using Microsoft.Office.Interop.Visio;

namespace VNC.Visio.VSTOAddIn.Domain
{
    public class ParagraphRow
    {
        public string IndFirst { get; set; }
        public string IndLeft { get; set; }
        public string IndRight { get; set; }
        public string SpLine { get; set; }
        public string SpBefore { get; set; }
        public string SpAfter { get; set; }
        public string HAlign { get; set; }
        public string Bullet { get; set; }
        public string BulletString { get; set; }
        public string BulletFont { get; set; }
        public string TextPosAfterBullet { get; set; }
        public string BulletSize { get; set; }
        public string Flags { get; set; }

        // TODO(crhodes)
        // Why is this Section and not Row

        public static ParagraphRow Get_ParagraphSection(Shape shape)
        {
            ParagraphRow paragraph = new ParagraphRow();

            Section paragraphSection = shape.Section[(short)VisSectionIndices.visSectionParagraph];
            Row paragraphRow = paragraphSection[0];

            paragraph.IndFirst = paragraphRow[VisCellIndices.visIndentFirst].FormulaU;
            paragraph.IndLeft = paragraphRow[VisCellIndices.visIndentLeft].FormulaU;
            paragraph.IndRight = paragraphRow[VisCellIndices.visIndentRight].FormulaU;
            paragraph.SpLine = paragraphRow[VisCellIndices.visSpaceLine].FormulaU;
            paragraph.SpBefore = paragraphRow[VisCellIndices.visSpaceBefore].FormulaU;
            paragraph.SpAfter = paragraphRow[VisCellIndices.visSpaceAfter].FormulaU;
            paragraph.HAlign = paragraphRow[VisCellIndices.visHorzAlign].FormulaU;
            paragraph.Bullet = paragraphRow[VisCellIndices.visBulletIndex].FormulaU;
            paragraph.BulletString = paragraphRow[VisCellIndices.visBulletString].FormulaU;
            paragraph.BulletFont = paragraphRow[VisCellIndices.visBulletFont].FormulaU;
            paragraph.TextPosAfterBullet = paragraphRow[VisCellIndices.visTextPosAfterBullet].FormulaU;
            paragraph.BulletSize = paragraphRow[VisCellIndices.visBulletFontSize].FormulaU;
            paragraph.Flags = paragraphRow[VisCellIndices.visFlags].FormulaU;

            return paragraph;
        }

        public static void Set_Paragraph_Section(Shape shape, ParagraphRow paragraph)
        {
            try
            {
                Section paragraphSection = shape.Section[(short)VisSectionIndices.visSectionParagraph];
                Row paragraphRow = paragraphSection[0];

                paragraphRow[VisCellIndices.visIndentFirst].FormulaForceU = paragraph.IndFirst;

                paragraphRow[VisCellIndices.visIndentLeft].FormulaForceU = paragraph.IndLeft;
                paragraphRow[VisCellIndices.visIndentRight].FormulaForceU = paragraph.IndRight;
                paragraphRow[VisCellIndices.visSpaceLine].FormulaForceU = paragraph.SpLine;
                paragraphRow[VisCellIndices.visSpaceBefore].FormulaForceU = paragraph.SpBefore;
                paragraphRow[VisCellIndices.visSpaceAfter].FormulaForceU = paragraph.SpAfter;
                paragraphRow[VisCellIndices.visHorzAlign].FormulaForceU = paragraph.HAlign;
                paragraphRow[VisCellIndices.visBulletIndex].FormulaForceU = paragraph.Bullet;
                paragraphRow[VisCellIndices.visBulletString].FormulaForceU = paragraph.BulletString;
                paragraphRow[VisCellIndices.visBulletFont].FormulaForceU = paragraph.BulletFont;
                paragraphRow[VisCellIndices.visTextPosAfterBullet].FormulaForceU = paragraph.TextPosAfterBullet;
                paragraphRow[VisCellIndices.visBulletFontSize].FormulaForceU = paragraph.BulletSize;
                paragraphRow[VisCellIndices.visFlags].FormulaForceU = paragraph.Flags;
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }
    }
}
