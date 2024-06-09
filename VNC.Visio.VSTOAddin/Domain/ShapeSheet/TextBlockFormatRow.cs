using System;

using Microsoft.Office.Interop.Visio;

namespace VNC.Visio.VSTOAddIn.Domain
{
    //public class TextBlockFormat
    //{
    //    public string LeftMargin = "Char.Size/2";
    //    public string RightMargin = "Char.Size/2";
    //    public string TextBkgnd = "0";
    //    public string TopMargin = "Char.Size/4";
    //    public string BottomMargin = "Char.Size/4";
    //    public string TextBkgndTrans = "0%";
    //    public string TextDirection = "0";
    //    public string VerticalAlign = "1";
    //    public string DefaultTabStop = "0.5 in";
    //}

    public class TextBlockFormatRow
    {
        public TextBlockFormatRow()
        {
            // Initialize Defaults
            //LeftMargin = "Char.Size/2";
            //RightMargin = "Char.Size/2";
            //TopMargin = "Char.Size/4";
            //BottomMargin = "Char.Size/4";
        }

        public string LeftMargin { get; set; } = "Char.Size/2";
        public string TopMargin { get; set; } = "Char.Size/2";
        public string RightMargin { get; set; } = "Char.Size/4";
        public string BottomMargin { get; set; } = "Char.Size/4";

        public string TextBkgnd { get; set; } = "0";
        public string TextBkgndTrans { get; set; } = "0%";
        public string TextDirection { get; set; } = "0";
        public string VerticalAlign { get; set; } = "1";
        public string DefaultTabStop { get; set; } = "0.5 in";

        public static TextBlockFormatRow GetRow(Shape shape)
        {
            TextBlockFormatRow row = new TextBlockFormatRow();

            // Shape Transform Section is part of object
            Section section = shape.Section[(short)VisSectionIndices.visSectionObject];
            Row sectionRow = section[(short)VisRowIndices.visRowText];

            // TODO(crhodes)
            // Handle multiple rows

            row.LeftMargin = sectionRow[VisCellIndices.visTxtBlkLeftMargin].FormulaU;
            row.TopMargin = sectionRow[VisCellIndices.visTxtBlkTopMargin].FormulaU;
            row.RightMargin = sectionRow[VisCellIndices.visTxtBlkRightMargin].FormulaU;
            row.BottomMargin = sectionRow[VisCellIndices.visTxtBlkBottomMargin].FormulaU;
            row.TextBkgnd = sectionRow[VisCellIndices.visTxtBlkBkgnd].FormulaU;
            row.TextBkgndTrans = sectionRow[VisCellIndices.visTxtBlkBkgndTrans].FormulaU;
            row.TextDirection = sectionRow[VisCellIndices.visTxtBlkDirection].FormulaU;
            row.VerticalAlign = sectionRow[VisCellIndices.visTxtBlkVerticalAlign].FormulaU;
            row.DefaultTabStop = sectionRow[VisCellIndices.visTxtBlkDefaultTabStop].FormulaU;

            return row;
        }

        public static void SetSection(Shape shape,
            TextBlockFormatRow textBlockFormat = null)
        {
            ValidateSectionExists(shape);

            try
            {
                shape.CellsSRC[
                    (short)VisSectionIndices.visSectionObject,
                    (short)VisRowIndices.visRowText,
                    (short)VisCellIndices.visTxtBlkLeftMargin].FormulaU = textBlockFormat.LeftMargin;
                shape.CellsSRC[
                    (short)VisSectionIndices.visSectionObject,
                    (short)VisRowIndices.visRowText,
                    (short)VisCellIndices.visTxtBlkTopMargin].FormulaU = textBlockFormat.TopMargin;
                shape.CellsSRC[
                    (short)VisSectionIndices.visSectionObject,
                    (short)VisRowIndices.visRowText,
                    (short)VisCellIndices.visTxtBlkRightMargin].FormulaU = textBlockFormat.RightMargin;
                shape.CellsSRC[
                    (short)VisSectionIndices.visSectionObject,
                    (short)VisRowIndices.visRowText,
                    (short)VisCellIndices.visTxtBlkBottomMargin].FormulaU = textBlockFormat.BottomMargin;
                shape.CellsSRC[
                    (short)VisSectionIndices.visSectionObject,
                    (short)VisRowIndices.visRowText,
                    (short)VisCellIndices.visTxtBlkDirection].FormulaU = textBlockFormat.TextDirection;
                shape.CellsSRC[
                    (short)VisSectionIndices.visSectionObject,
                    (short)VisRowIndices.visRowText,
                    (short)VisCellIndices.visTxtBlkVerticalAlign].FormulaU = textBlockFormat.VerticalAlign;
                shape.CellsSRC[
                    (short)VisSectionIndices.visSectionObject,
                    (short)VisRowIndices.visRowText,
                    (short)VisCellIndices.visTxtBlkBkgnd].FormulaU = textBlockFormat.TextBkgnd;
                shape.CellsSRC[
                    (short)VisSectionIndices.visSectionObject,
                    (short)VisRowIndices.visRowText,
                    (short)VisCellIndices.visTxtBlkBkgndTrans].FormulaU = textBlockFormat.TextBkgndTrans;
                shape.CellsSRC[
                    (short)VisSectionIndices.visSectionObject,
                    (short)VisRowIndices.visRowText,
                    (short)VisCellIndices.visTxtBlkDefaultTabStop].FormulaU = textBlockFormat.DefaultTabStop;

            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }

        public static void SetMargins(Shape shape,
            string LeftMargin,
            string TopMargin,
            string RightMargin,
            string BottomMargin)
        {
            // TODO(crhodes):
            // Consider making some of the arguments optional with reasonable defaults

            try
            {
                shape.CellsSRC[
                    (short)VisSectionIndices.visSectionObject,
                    (short)VisRowIndices.visRowText,
                    (short)VisCellIndices.visTxtBlkLeftMargin].FormulaU = LeftMargin;
                shape.CellsSRC[
                    (short)VisSectionIndices.visSectionObject,
                    (short)VisRowIndices.visRowText,
                    (short)VisCellIndices.visTxtBlkTopMargin].FormulaU = TopMargin;
                shape.CellsSRC[
                    (short)VisSectionIndices.visSectionObject,
                    (short)VisRowIndices.visRowText,
                    (short)VisCellIndices.visTxtBlkRightMargin].FormulaU = RightMargin;
                shape.CellsSRC[
                    (short)VisSectionIndices.visSectionObject,
                    (short)VisRowIndices.visRowText,
                    (short)VisCellIndices.visTxtBlkBottomMargin].FormulaU = BottomMargin;
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }

        public static void ValidateSectionExists(Shape shape)
        {
            // TextBlockFormat exists as a row in the SectionObject!

            if (0 == shape.RowExists[
                (short)VisSectionIndices.visSectionObject,
                (short)VisRowIndices.visRowText,
                (short)VisExistsFlags.visExistsAnywhere])
            {
                try
                {
                    shape.AddRow(
                        (short)VisSectionIndices.visSectionObject,
                        (short)VisRowIndices.visRowText,
                        (short)VisRowTags.visTagDefault);
                }
                catch (Exception ex)
                {
                    Log.Error(ex, Common.LOG_CATEGORY);
                }
            }
        }
    }
}
