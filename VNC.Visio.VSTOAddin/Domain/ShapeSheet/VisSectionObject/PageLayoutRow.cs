using System;
using System.Windows;

using Microsoft.Office.Interop.Visio;

namespace VNC.Visio.VSTOAddIn.Domain
{
    public class PageLayoutRow
    {
        public string PlaceStyle { get; set; }
        public string PlaceDepth { get; set; }
        public string PlowCode { get; set; }
        public string ResizePage { get; set; }
        public string DynamicsOff { get; set; }
        public string EnableGrid { get; set; }
        public string CtrlAsInput { get; set; }
        public string LineAdjustFrom { get; set; }
        public string PlaceFlip { get; set; }
        public string AvoidPageBreaks { get; set; }
        public string BlockSizeX { get; set; }
        public string BlockSizeY { get; set; }
        public string AvenueSizeX { get; set; }
        public string AvenueSizeY { get; set; }
        public string RouteStyle { get; set; }
        public string PageLineJumpDirX { get; set; }
        public string PageLineJumpDirY { get; set; }
        public string LineAdjustTo { get; set; }
        public string LineRouteExt { get; set; }
        public string LineToNodeX { get; set; }
        public string LineToNodeY { get; set; }
        public string LineToLineX { get; set; }
        public string LineToLineY { get; set; }
        public string LineJumpFactorX { get; set; }
        public string LineJumpFactorY { get; set; }
        public string LineJumpCode { get; set; }
        public string LineJumpStyle { get; set; }
        public string PageShapeSplit { get; set; }

        public static PageLayoutRow GetRow(Shape shape)
        {
            PageLayoutRow row = new PageLayoutRow();

            Section section = shape.Section[(short)VisSectionIndices.visSectionObject];

            if (Helpers.RowExists(shape, VisSectionIndices.visSectionObject, VisRowIndices.visRowPageLayout))
            {
                Row sectionRow = section[(short)VisRowIndices.visRowPageLayout];

                row.PlaceStyle = sectionRow[VisCellIndices.visPLOPlaceStyle].FormulaU;
                row.PlaceDepth = sectionRow[VisCellIndices.visPLOPlaceDepth].FormulaU;
                row.PlowCode = sectionRow[VisCellIndices.visPLOPlowCode].FormulaU;
                row.ResizePage = sectionRow[VisCellIndices.visPLOResizePage].FormulaU;
                row.DynamicsOff = sectionRow[VisCellIndices.visPLODynamicsOff].FormulaU;
                row.EnableGrid = sectionRow[VisCellIndices.visPLOEnableGrid].FormulaU;
                row.CtrlAsInput = sectionRow[VisCellIndices.visPLOCtrlAsInput].FormulaU;
                row.LineAdjustFrom = sectionRow[VisCellIndices.visPLOLineAdjustFrom].FormulaU;
                row.PlaceFlip = sectionRow[VisCellIndices.visPLOPlaceFlip].FormulaU;
                row.AvoidPageBreaks = sectionRow[VisCellIndices.visPLOAvoidPageBreaks].FormulaU;
                row.BlockSizeX = sectionRow[VisCellIndices.visPLOBlockSizeX].FormulaU;
                row.BlockSizeY = sectionRow[VisCellIndices.visPLOBlockSizeY].FormulaU;
                row.AvenueSizeX = sectionRow[VisCellIndices.visPLOAvenueSizeX].FormulaU;
                row.AvenueSizeY = sectionRow[VisCellIndices.visPLOAvenueSizeY].FormulaU;
                row.RouteStyle = sectionRow[VisCellIndices.visPLORouteStyle].FormulaU;
                row.PageLineJumpDirX = sectionRow[VisCellIndices.visPLOJumpDirX].FormulaU;
                row.PageLineJumpDirY = sectionRow[VisCellIndices.visPLOJumpDirY].FormulaU;
                row.LineAdjustTo = sectionRow[VisCellIndices.visPLOLineAdjustTo].FormulaU;
                row.LineRouteExt = sectionRow[VisCellIndices.visPLOLineRouteExt].FormulaU;
                row.LineToNodeX = sectionRow[VisCellIndices.visPLOLineToNodeX].FormulaU;
                row.LineToNodeY = sectionRow[VisCellIndices.visPLOLineToNodeY].FormulaU;
                row.LineToLineX = sectionRow[VisCellIndices.visPLOLineToLineX].FormulaU;
                row.LineToLineY = sectionRow[VisCellIndices.visPLOLineToLineY].FormulaU;
                row.LineJumpFactorX = sectionRow[VisCellIndices.visPLOJumpFactorX].FormulaU;
                row.LineJumpFactorY = sectionRow[VisCellIndices.visPLOJumpFactorY].FormulaU;
                row.LineJumpCode = sectionRow[VisCellIndices.visPLOJumpCode].FormulaU;
                row.LineJumpStyle = sectionRow[VisCellIndices.visPLOJumpStyle].FormulaU;
                row.PageShapeSplit = sectionRow[VisCellIndices.visPLOSplit].FormulaU;
            }
            else
            {
                MessageBox.Show("No visRowPageLayout exists");
            }

            return row;
        }

        public static void SetRow(Shape shape, PageLayoutRow pageLayout)
        {
            try
            {
                Section section = shape.Section[(short)VisSectionIndices.visSectionObject];

                if (Helpers.RowExists(shape, VisSectionIndices.visSectionObject, VisRowIndices.visRowPageLayout))
                {
                    Row sectionRow = section[(short)VisRowIndices.visRowPageLayout];

                    sectionRow[VisCellIndices.visPLOPlaceStyle].FormulaU = pageLayout.PlaceStyle;
                    sectionRow[VisCellIndices.visPLOPlaceDepth].FormulaU = pageLayout.PlaceDepth;
                    sectionRow[VisCellIndices.visPLOPlowCode].FormulaU = pageLayout.PlowCode;
                    sectionRow[VisCellIndices.visPLOResizePage].FormulaU = pageLayout.ResizePage;
                    sectionRow[VisCellIndices.visPLODynamicsOff].FormulaU = pageLayout.DynamicsOff;
                    sectionRow[VisCellIndices.visPLOEnableGrid].FormulaU = pageLayout.EnableGrid;
                    sectionRow[VisCellIndices.visPLOCtrlAsInput].FormulaU = pageLayout.CtrlAsInput;
                    sectionRow[VisCellIndices.visPLOLineAdjustFrom].FormulaU = pageLayout.LineAdjustFrom;
                    sectionRow[VisCellIndices.visPLOPlaceFlip].FormulaU = pageLayout.PlaceFlip;
                    sectionRow[VisCellIndices.visPLOAvoidPageBreaks].FormulaU = pageLayout.AvoidPageBreaks;
                    sectionRow[VisCellIndices.visPLOBlockSizeX].FormulaU = pageLayout.BlockSizeX;
                    sectionRow[VisCellIndices.visPLOBlockSizeY].FormulaU = pageLayout.BlockSizeY;
                    sectionRow[VisCellIndices.visPLOAvenueSizeX].FormulaU = pageLayout.AvenueSizeX;
                    sectionRow[VisCellIndices.visPLOAvenueSizeY].FormulaU = pageLayout.AvenueSizeY;
                    sectionRow[VisCellIndices.visPLORouteStyle].FormulaU = pageLayout.RouteStyle;
                    sectionRow[VisCellIndices.visPLOJumpDirX].FormulaU = pageLayout.PageLineJumpDirX;
                    sectionRow[VisCellIndices.visPLOJumpDirY].FormulaU = pageLayout.PageLineJumpDirY;
                    sectionRow[VisCellIndices.visPLOLineAdjustTo].FormulaU = pageLayout.LineAdjustTo;
                    sectionRow[VisCellIndices.visPLOLineRouteExt].FormulaU = pageLayout.LineRouteExt;
                    sectionRow[VisCellIndices.visPLOLineToNodeX].FormulaU = pageLayout.LineToNodeX;
                    sectionRow[VisCellIndices.visPLOLineToNodeY].FormulaU = pageLayout.LineToNodeY;
                    sectionRow[VisCellIndices.visPLOLineToLineX].FormulaU = pageLayout.LineToLineX;
                    sectionRow[VisCellIndices.visPLOLineToLineY].FormulaU = pageLayout.LineToLineY;
                    sectionRow[VisCellIndices.visPLOJumpFactorX].FormulaU = pageLayout.LineJumpFactorX;
                    sectionRow[VisCellIndices.visPLOJumpFactorY].FormulaU = pageLayout.LineJumpFactorY;
                    sectionRow[VisCellIndices.visPLOJumpCode].FormulaU = pageLayout.LineJumpCode;
                    sectionRow[VisCellIndices.visPLOJumpStyle].FormulaU = pageLayout.LineJumpStyle;
                    sectionRow[VisCellIndices.visPLOSplit].FormulaU = pageLayout.PageShapeSplit;
                }
                else
                {
                    MessageBox.Show("No visRowPageLayout exists");
                }
            }
            catch (Exception ex)
            {
                Log.ERROR(ex, Common.LOG_CATEGORY);
            }
        }
    }
}
