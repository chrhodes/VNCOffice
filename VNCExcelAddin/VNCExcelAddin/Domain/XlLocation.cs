using System;
using System.Diagnostics;
using System.Drawing;
using System.Reflection;

using Microsoft.Office.Interop.Excel;

namespace VNCExcelAddin.Domain
{
    public partial class Excel
    {
        /// <summary>
        /// Cursor based support for adding information to a worksheet
        /// </summary>
        public class XlLocation
        {
            // TODO(crhodes):
            //	Need to revisit all the offset stuff. It is not clean.  
            // Lots of conditional -1 stuff.  Ugh

            //public int RowStart;

            public int RowOffset = 0;
            public int RowOffsetMax = 0;

            //public int ColumnStart;
            public int ColumnOffset = 0;
            public int ColumnOffsetMax = 0;

            public int RowsAdded = 0;
            public int ColumnsAdded = 0;

            private Range startRange = null;
            private Range nextRange = null;
            public Range CurrentRange { get; set; }

            public Range OffsetRange { get; set; }

            public int GroupStartRow = 0;
            public int GroupStartColumn = 0;
            public int GroupEndRow = 0;
            public int GroupEndColumn = 0;

            public bool OrientVertical = true;

            public int MarkStartRow = 0;
            public int MarkStartColumn = 0;
            public int MarkEndRow = 0;
            public int MarkEndColumn = 0;

            public int TableStartRow = 0;
            public int TableStartColumn = 0;
            public int TableEndRow = 0;
            public int TableEndColumn = 0;

            public Worksheet workSheet = null;

            public CellFormatSpecification ContentFormat = new CellFormatSpecification("ContentFormat");
            public CellFormatSpecification ContentLeftFormat = new CellFormatSpecification("ContentLeftFormat");
            public CellFormatSpecification ContentRightFormat = new CellFormatSpecification("ContentRightFormat");

            public CellFormatSpecification SectionCenterFormat = new CellFormatSpecification("SectionCenterFormat");
            public CellFormatSpecification SectionLeftFormat = new CellFormatSpecification("SectionLeftFormat");
            public CellFormatSpecification SectionRightFormat = new CellFormatSpecification("SectionRightFormat");

            public CellFormatSpecification LabelLargeFormat = new CellFormatSpecification("LabelLargeFormat");
            public CellFormatSpecification LabelLargeLeftFormat = new CellFormatSpecification("LabelLargeLeftFormat");
            public CellFormatSpecification LabelLargeRightFormat = new CellFormatSpecification("LabelLargeRightFormat");

            public CellFormatSpecification LabelFormat = new CellFormatSpecification("LabelFormat");
            public CellFormatSpecification LabelLeftFormat = new CellFormatSpecification("LabelLeftFormat");
            public CellFormatSpecification LabelRightFormat = new CellFormatSpecification("LabelRightFormat");

            public CellFormatSpecification LabelSmallFormat = new CellFormatSpecification("LabelSmallFormat");
            public CellFormatSpecification LabelSmallLeftFormat = new CellFormatSpecification("LabelSmallLeftFormat");
            public CellFormatSpecification LabelSmallRightFormat = new CellFormatSpecification("LabelSmallRightFormat");

            public CellFormatSpecification HeaderFormat = new CellFormatSpecification("HeaderFormat");
            public CellFormatSpecification HeaderLeftFormat = new CellFormatSpecification("HeaderLeftFormat");
            public CellFormatSpecification HeaderRightFormat = new CellFormatSpecification("HeaderRightFormat");

            public CellFormatSpecification HeaderSmFormat = new CellFormatSpecification("HeaderSmFormat");
            public CellFormatSpecification HeaderSmLeftFormat = new CellFormatSpecification("HeaderSmLeftFormat");
            public CellFormatSpecification HeaderSmRightFormat = new CellFormatSpecification("HeaderSmRightFormat");

            #region Constructors

            public XlLocation(int row, int column)
            {

                // TODO(crhodes):
                //	Not sure this makes sense.  May need to always specify a Worksheet
                //RowStart = row;
                //ColumnStart = column;
            }

            public XlLocation(Worksheet ws, int row, int column)
                : this(row, column)
            {
                workSheet = ws;
                startRange = (Range)ws.Cells[row, column];
                nextRange = startRange;

                InitializeFont(startRange.Font, ContentFormat);
                InitializeFont(startRange.Font, ContentLeftFormat);
                ContentLeftFormat.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                InitializeFont(startRange.Font, ContentRightFormat);
                ContentRightFormat.HorizontalAlignment = XlHAlign.xlHAlignRight;

                // Header Format

                InitializeFont(startRange.Font, HeaderFormat);
                HeaderFormat.Font.Size = 12;
                HeaderFormat.Font.Bold = true;

                InitializeFont(HeaderFormat, HeaderLeftFormat);
                HeaderLeftFormat.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                InitializeFont(HeaderFormat, HeaderRightFormat);
                HeaderRightFormat.HorizontalAlignment = XlHAlign.xlHAlignRight;

                // Header Small Format

                InitializeFont(startRange.Font, HeaderSmFormat);
                HeaderSmFormat.Font.Size = 10;
                HeaderSmFormat.Font.Bold = true;

                InitializeFont(HeaderSmFormat, HeaderSmLeftFormat);
                HeaderSmLeftFormat.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                InitializeFont(HeaderSmFormat, HeaderSmRightFormat);
                HeaderSmRightFormat.HorizontalAlignment = XlHAlign.xlHAlignRight;

                // Section Format

                InitializeFont(startRange.Font, SectionCenterFormat);
                SectionCenterFormat.Font.Size = 16;
                SectionCenterFormat.Font.Bold = true;
                SectionCenterFormat.Font.Color = Color.Blue;
                SectionCenterFormat.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                InitializeFont(SectionCenterFormat, SectionLeftFormat);
                SectionLeftFormat.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                InitializeFont(SectionCenterFormat, SectionRightFormat);
                SectionRightFormat.HorizontalAlignment = XlHAlign.xlHAlignRight;

                // Label Large Format

                InitializeFont(startRange.Font, LabelLargeFormat);
                LabelLargeFormat.Font.Size = 14;
                LabelLargeFormat.Font.Bold = true;

                InitializeFont(LabelLargeFormat, LabelLargeLeftFormat);
                LabelLargeLeftFormat.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                InitializeFont(LabelLargeFormat, LabelLargeRightFormat);
                LabelLargeRightFormat.HorizontalAlignment = XlHAlign.xlHAlignRight;

                // Label Format (Default)

                InitializeFont(startRange.Font, LabelFormat);
                LabelFormat.Font.Size = 12;
                LabelFormat.Font.Bold = true;

                InitializeFont(LabelFormat, LabelLeftFormat);
                LabelLeftFormat.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                InitializeFont(LabelFormat, LabelRightFormat);
                LabelRightFormat.HorizontalAlignment = XlHAlign.xlHAlignRight;

                // Label Small Format

                InitializeFont(startRange.Font, LabelSmallFormat);
                LabelSmallFormat.Font.Size = 10;
                LabelSmallFormat.Font.Bold = true;

                InitializeFont(LabelSmallFormat, LabelSmallLeftFormat);
                LabelSmallLeftFormat.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                InitializeFont(LabelSmallFormat, LabelSmallRightFormat);
                LabelSmallRightFormat.HorizontalAlignment = XlHAlign.xlHAlignRight;
            }

            public XlLocation(Worksheet ws, int row, int column, bool orientVertical)
                : this(ws, row, column)
            {
                OrientVertical = orientVertical;
            }

            #endregion

            #region Main Methods

            public XlLocation AddColumn(int columns = 1)
            {
                if (Common.DisplayXlLocationUpdates)
                {
                    StackFrame frame = new StackFrame(1);
                    MethodBase caller = frame.GetMethod();

                    Excel.DisplayInWatchWindow(this);
                }

                // nextRange is currently where new output should be.
                CurrentRange = nextRange;

                // Update to reflect new row.
                ColumnsAdded += columns;
                nextRange = nextRange.Offset[0, columns];

                ColumnOffset = columns;
                //ColumnOffsetMax = ColumnOffset;

                if (columns > ColumnsAdded)
                {
                    ColumnsAdded = columns;
                }

                return this;
            }

            [Obsolete("AddColumnX() has been replaced with AddColumn() which returns an XlLocation instead of a Range")]
            public Range AddColumnX(int columns = 1)
            {
                if (Common.DisplayXlLocationUpdates)
                {
                    StackFrame frame = new StackFrame(1);
                    MethodBase caller = frame.GetMethod();

                    Excel.DisplayInWatchWindow(this);
                }

                // currentRange is currently where it should be.
                Range rngOutput = nextRange;

                // Update to reflect new row.
                ColumnsAdded += columns;
                nextRange = nextRange.Offset[0, columns];

                ColumnOffset = columns;
                //ColumnOffsetMax = ColumnOffset;

                if (columns > ColumnsAdded)
                {
                    ColumnsAdded = columns;
                }

                return rngOutput;
            }

            public XlLocation AddOffsetColumn()
            {
                if (Common.DisplayXlLocationUpdates)
                {
                    StackFrame frame = new StackFrame(1);
                    MethodBase caller = frame.GetMethod();

                    Excel.DisplayInWatchWindow(this);
                }

                // nextRange is currently where new output should be.

                CurrentRange = nextRange;
                OffsetRange = nextRange.Offset[0, ColumnOffset];
                ColumnOffset++;

                //Range rngOutput = nextRange.Offset[0, ColumnOffset++];

                if (ColumnOffsetMax < ColumnOffset)
                {
                    ColumnOffsetMax = ColumnOffset;
                }

                if (ColumnOffset > ColumnsAdded)
                {
                    ColumnsAdded = ColumnOffset;
                }

                return this;
            }

            [Obsolete("\nAddOffsetColumnX has been replaced with AddOffsetColumn()" +
                "\nwhich returns an XlLocation instead of a Range.  " +
                "\nUse AddOffsetContentToCell(AddOffsetColumn(), ...) instead.")]
            public Range AddOffsetColumnX()
            {
                if (Common.DisplayXlLocationUpdates)
                {
                    StackFrame frame = new StackFrame(1);
                    MethodBase caller = frame.GetMethod();

                    //Excel.DisplayInWatchWindow(this);
                }

                Range rngOutput = nextRange.Offset[0, ColumnOffset++];

                if (ColumnOffsetMax < ColumnOffset)
                {
                    ColumnOffsetMax = ColumnOffset;
                }

                if (ColumnOffset > ColumnsAdded)
                {
                    ColumnsAdded = ColumnOffset;
                }

                return rngOutput;
            }

            [Obsolete("AddOffsetRowX has been replaced with AddOffsetRow which returns an XlLocation instead of a Range")]
            public Range AddOffsetRowX()
            {
                if (Common.DisplayXlLocationUpdates)
                {
                    StackFrame frame = new StackFrame(1);
                    MethodBase caller = frame.GetMethod();

                    //Excel.DisplayInWatchWindow(this, caller.Name);
                }

                Range rngOutput = nextRange.Offset[RowOffset++, 0];

                if (RowOffset > RowsAdded)
                {
                    RowsAdded = RowOffset;
                }

                return rngOutput;
            }

            // Should this be AddColumnsToRow?
            // Should column offset be 0 or reflect how many columns have been added?

            // TODO(crhodes):
            //	Maybe we should just use ColumnOffset

            [Obsolete("AddRowX has been replaced with AddRow which returns an XlLocation instead of a Range")]
            public Range AddRowX(int columns = 0)
            {
                if (Common.DisplayXlLocationUpdates)
                {
                    StackFrame frame = new StackFrame(1);
                    MethodBase caller = frame.GetMethod();

                    //Excel.DisplayInWatchWindow(this, caller.Name);
                }

                // currentRange is currently where it should be.
                Range rngOutput = nextRange;

                // Update to reflect new row.
                RowsAdded++;
                nextRange = nextRange.Offset[1, 0];

                ColumnOffset = columns;

                if (columns > ColumnsAdded)
                {
                    ColumnsAdded = columns;
                }

                if (ColumnOffsetMax < ColumnOffset)
                {
                    ColumnOffsetMax = ColumnOffset;
                }

                return rngOutput;
            }

            public XlLocation AddRow(int columns = 0)
            {
                if (Common.DisplayXlLocationUpdates)
                {
                    StackFrame frame = new StackFrame(1);
                    MethodBase caller = frame.GetMethod();

                    //Excel.DisplayInWatchWindow(this, caller.Name);
                }

                // nextRange is currently where new output should be.
                CurrentRange = nextRange;

                //Range rngOutput = currentRange;

                XlLocation currentLocation = this;

                // Update to reflect new row.
                RowsAdded++;
                nextRange = nextRange.Offset[1, 0];

                ColumnOffset = columns;

                if (columns > ColumnsAdded)
                {
                    ColumnsAdded = columns;
                }

                if (ColumnOffsetMax < ColumnOffset)
                {
                    ColumnOffsetMax = ColumnOffset;
                }

                return this;
            }

            public void ClearOffsets(Boolean clearOffsetMax = false)
            {
                if (Common.DisplayXlLocationUpdates)
                {
                    StackFrame frame = new StackFrame(1);
                    MethodBase caller = frame.GetMethod();

                    //Excel.DisplayInWatchWindow(this, caller.Name);
                }

                // HACK(crhodes)
                // 
                RowOffset = 0;
                ColumnOffset = 0;

                if (clearOffsetMax)
                {
                    RowOffsetMax = 0;
                    ColumnOffsetMax = 0;
                }
            }

            public void SetColumnOffset(int column)
            {
                if (Common.DisplayXlLocationUpdates)
                {
                    StackFrame frame = new StackFrame(1);
                    MethodBase caller = frame.GetMethod();

                    //Excel.DisplayInWatchWindow(this, caller.Name);
                }

                ColumnOffset = column;
            }

            public void CreateTable(string tableName)
            {
                if (Common.DisplayXlLocationUpdates)
                {
                    StackFrame frame = new StackFrame(1);
                    MethodBase caller = frame.GetMethod();

                    //Excel.DisplayInWatchWindow(this, caller.Name);
                }

                Worksheet ws = nextRange.Worksheet;

                if (IsValidTableRange(TableStartRow, TableStartColumn, TableEndRow, TableEndColumn))
                {
                    ListObject table = ws.ListObjects.Add(XlListObjectSourceType.xlSrcRange,
                        ws.Range[
                            ws.Cells[TableStartRow, TableStartColumn],
                            ws.Cells[TableEndRow, TableEndColumn]],
                        Type.Missing,
                        XlYesNoGuess.xlYes);

                    if (tableName != null)
                    {
                        table.Name = tableName;
                    }

                    table.ShowTotals = true;
                    table.ListColumns[1].TotalsCalculation = XlTotalsCalculation.xlTotalsCalculationCount;
                }
            }

            public void DecrementColumns(int count = 1)
            {
                if (Common.DisplayXlLocationUpdates)
                {
                    StackFrame frame = new StackFrame(1);
                    MethodBase caller = frame.GetMethod();

                    //Excel.DisplayInWatchWindow(this, caller.Name);
                }

                // TODO(crhodes):
                ////	Decide if need to add to ColumnsAdded or just adjust range.

                //ColumnsAdded -= count;

                nextRange = nextRange.Offset[0, -count];
            }

            public void DecrementRows(int count = 1)
            {
                if (Common.DisplayXlLocationUpdates)
                {
                    StackFrame frame = new StackFrame(1);
                    MethodBase caller = frame.GetMethod();

                    //Excel.DisplayInWatchWindow(this, caller.Name);
                }

                // TODO(crhodes):
                //	Decide if need to add to RowsAdded or just adjust range.

                //RowsAdded -= count;
                nextRange = nextRange.Offset[-count, 0];
            }

            public void EndSectionAndSetNextLocation(bool orientVertical)
            {
                if (!orientVertical)
                {
                    // Skip past the info just added.
                    SetLocation(RowStart, TableEndColumn + 1);
                }
            }

            public Range GetCurrentRange()
            {
                return nextRange;
            }

            public Range GetStartRange()
            {
                return startRange;
            }

            public void Group(bool orientVertical, bool hide = true)
            {
                if (Common.DisplayXlLocationUpdates)
                {
                    StackFrame frame = new StackFrame(1);
                    MethodBase caller = frame.GetMethod();

                    //Excel.DisplayInWatchWindow(this, caller.Name);
                }

                if (orientVertical)
                {
                    Excel.GroupAndHideRows(workSheet, GroupStartRow, GroupEndRow, hide);
                }
                else
                {
                    Excel.GroupAndHideColumns(workSheet, GroupStartColumn, GroupEndColumn, hide);
                }
            }

            public void IncrementColumns(int count = 1)
            {
                if (Common.DisplayXlLocationUpdates)
                {
                    StackFrame frame = new StackFrame(1);
                    MethodBase caller = frame.GetMethod();

                    //Excel.DisplayInWatchWindow(this, caller.Name);
                }

                // TODO(crhodes):
                //	Decide if need to add to ColumnsAdded or just adjust range.
                // Seems odd that Increment adds to ColumnsAdded but Decrement does not.
                ColumnsAdded += count;
                //ColumnOffsetMax += count;
                nextRange = nextRange.Offset[0, count];

                if (nextRange.Column > ColumnOffsetMax)
                {
                    // TODO(crhodes)
                    // Should we bump ColumnOffsetMax
                    // Investigting Category Node issue

                    ColumnOffsetMax = nextRange.Column;
                }
            }

            public void IncrementRows(int count = 1)
            {
                if (Common.DisplayXlLocationUpdates)
                {
                    StackFrame frame = new StackFrame(1);
                    MethodBase caller = frame.GetMethod();

                    //Excel.DisplayInWatchWindow(this, caller.Name);
                }

                // TODO(crhodes):
                //	Decide if need to add to RowsAdded or just adjust range.

                RowsAdded += count;
                nextRange = nextRange.Offset[count, 0];
            }

            public Range InsertRow(int columns)
            {
                StackFrame frame = new StackFrame(1);
                MethodBase caller = frame.GetMethod();

                //Excel.DisplayInWatchWindow(this, caller.Name);

                // currentRange is already where it should be.
                //Range rngOutput = currentRange;
                ////RowsAdded++;
                ////Range rngOutput = currentRange.Offset[RowsAdded++, 0];
                //currentRange = currentRange.Worksheet.Cells[CurrentRow(), CurrentColumn()];

                if (columns > ColumnsAdded)
                {
                    ColumnsAdded = columns;
                }

                return nextRange;
            }

            public void MarkStart(Excel.MarkType type = MarkType.None)
            {
                if (Common.DisplayXlLocationUpdates)
                {
                    StackFrame frame = new StackFrame(1);
                    MethodBase caller = frame.GetMethod();

                    //Excel.DisplayInWatchWindow(this, caller.Name);
                }

                // TODO(crhodes)
                // Previous code called ClearOffsets then DisplayInWatchWindow
                // Decide how to handle.  See above


                // Seems like we always want to do this.
                // 
                ClearOffsets(clearOffsetMax: true);

                //Excel.DisplayInWatchWindow(this, caller.Name);

                switch (type)
                {
                    case Excel.MarkType.None:
                        MarkStartRow = nextRange.Row;
                        MarkStartColumn = nextRange.Column;

                        break;

                    case Excel.MarkType.Group:
                        GroupStartRow = nextRange.Row;
                        GroupStartColumn = nextRange.Column;

                        break;

                    case Excel.MarkType.Table:
                        TableStartRow = nextRange.Row;
                        TableStartColumn = nextRange.Column;

                        break;

                    case Excel.MarkType.GroupTable:
                        GroupStartRow = nextRange.Row;
                        GroupStartColumn = nextRange.Column;
                        TableStartRow = nextRange.Row;
                        TableStartColumn = nextRange.Column;

                        break;
                }

                Excel.DisplayInWatchWindow(this, "End");
            }

            public void MarkEnd(Excel.MarkType type = MarkType.None, string tableName = null)
            {
                if (Common.DisplayXlLocationUpdates)
                {
                    StackFrame frame = new StackFrame(1);
                    MethodBase caller = frame.GetMethod();

                    Excel.DisplayInWatchWindow(this, caller.Name);
                }

                switch (type)
                {
                    case Excel.MarkType.None:
                        MarkEndRow = nextRange.Row + RowsAdded;
                        MarkEndColumn = nextRange.Column + ColumnsAdded;

                        break;

                    case Excel.MarkType.Group:
                        GroupEndRow = nextRange.Row + RowOffset - 1;
                        GroupEndColumn = nextRange.Column + ColumnOffsetMax;

                        if (ColumnOffsetMax != 0)
                        {
                            // We have used ColumnOffsetMax and it was 
                            // incremented past end
                            GroupEndColumn--;
                        }

                        break;

                    case Excel.MarkType.Table:
                        TableEndRow = nextRange.Row + RowOffset - 1;
                        //TableEndColumn = currentRange.Column + ColumnOffset - 1;
                        TableEndColumn = nextRange.Column + ColumnOffsetMax - 1;

                        CreateTable(tableName);
                        //if (tableName != null)
                        //{
                        //    CreateTable(tableName);
                        //}

                        break;

                    case Excel.MarkType.GroupTable:
                        GroupEndRow = nextRange.Row + RowOffset - 1;
                        GroupEndColumn = nextRange.Column + ColumnOffsetMax;

                        if (ColumnOffsetMax != 0)
                        {
                            // We have used ColumnOffsetMax and it was 
                            // incremented past end
                            GroupEndColumn--;
                        }

                        TableEndRow = nextRange.Row + RowOffset - 1;
                        //TableEndColumn = currentRange.Column + ColumnOffset - 1;
                        TableEndColumn = nextRange.Column + ColumnOffsetMax - 1;

                        CreateTable(tableName);
                        //if (tableName != null)
                        //{
                        //    CreateTable(tableName);
                        //}

                        break;
                }

                Excel.DisplayInWatchWindow(this, "End");

            }

            /// <summary>
            /// Increment insertion point based on layout orientation.
            /// TODO(crhodes): Handle all four positions in future.
            /// </summary>
            /// <param name="orientVertical"></param>
            public XlLocation IncrementPosition(bool orientVertical)
            {
                if (orientVertical)
                {
                    IncrementRows();
                }
                else
                {
                    IncrementColumns();
                }

                return this;
            }

            #region Fields and Properties

            public int ColumnCurrent { get { return nextRange.Column; } }
            public int ColumnStart { get { return startRange.Column; } }

            public int RowCurrent { get { return nextRange.Row; } }
            public int RowStart { get { return startRange.Row; } }

            #endregion

            public void SetRow(int row)
            {
                nextRange = (Range)workSheet.Cells[row, nextRange.Column];
            }

            public void SetColumn(int column)
            {
                nextRange = (Range)workSheet.Cells[nextRange.Row, column];
            }

            public void SetLocation(int row, int column)
            {
                nextRange = (Range)workSheet.Cells[row, column];
            }

            public void UpdateOffsets()
            {
                int currentColumnOffset = ColumnCurrent - GroupStartColumn + 1;

                if (currentColumnOffset > ColumnOffsetMax)
                {
                    ColumnOffsetMax = currentColumnOffset;
                }

                int currentRowOffset = RowCurrent - GroupStartRow + 1;

                if (currentRowOffset > RowOffsetMax)
                {
                    RowOffsetMax = currentRowOffset;
                }
            }

            #endregion

            #region Private Methods

            private bool IsValidTableRange(int startRow, int startColumn, int endRow, int endColumn)
            {
                bool isValid = false;

                // Need to have at least one row (+ header) and one column in table.

                if (startRow > 0 && startColumn > 0
                    && (endRow - startRow) > 0
                    && (endColumn - startColumn) >= 0
                    )
                {
                    isValid = true;
                }

                return isValid;
            }

            #endregion

            public static void InitializeFont(Microsoft.Office.Interop.Excel.Font rangeFont, CellFormatSpecification formatSpecification)
            {
                formatSpecification.Font.Background = rangeFont.Background;
                formatSpecification.Font.Bold = rangeFont.Bold;
                formatSpecification.Font.Color = rangeFont.Color;
                formatSpecification.Font.ColorIndex = rangeFont.ColorIndex;
                formatSpecification.Font.FontStyle = rangeFont.FontStyle;
                formatSpecification.Font.Italic = rangeFont.Italic;
                formatSpecification.Font.Name = rangeFont.Name;
                formatSpecification.Font.OutlineFont = rangeFont.OutlineFont;
                formatSpecification.Font.Shadow = rangeFont.Shadow;
                formatSpecification.Font.Size = rangeFont.Size;
                formatSpecification.Font.Strikethrough = rangeFont.Strikethrough;
                formatSpecification.Font.Subscript = rangeFont.Subscript;
                formatSpecification.Font.Superscript = rangeFont.Superscript;
                formatSpecification.Font.Underline = rangeFont.Underline;
                // null by default
                //formatSpecification.Font.ThemeColor = rangeFont.ThemeColor;
                formatSpecification.Font.TintAndShade = rangeFont.TintAndShade;
                formatSpecification.Font.ThemeFont = rangeFont.ThemeFont;
            }

            public static void InitializeFont(CellFormatSpecification source, CellFormatSpecification target)
            {
                target.Font.Background = source.Font.Background;
                target.Font.Bold = source.Font.Bold;
                target.Font.Color = source.Font.Color;
                target.Font.ColorIndex = source.Font.ColorIndex;
                target.Font.FontStyle = source.Font.FontStyle;
                target.Font.Italic = source.Font.Italic;
                target.Font.Name = source.Font.Name;
                target.Font.OutlineFont = source.Font.OutlineFont;
                target.Font.Shadow = source.Font.Shadow;
                target.Font.Size = source.Font.Size;
                target.Font.Strikethrough = source.Font.Strikethrough;
                target.Font.Subscript = source.Font.Subscript;
                target.Font.Superscript = source.Font.Superscript;
                target.Font.Underline = source.Font.Underline;
                // null by default
                //target.Font.ThemeColor = source.Font.ThemeColor;
                target.Font.TintAndShade = source.Font.TintAndShade;
                target.Font.ThemeFont = source.Font.ThemeFont;
            }

            public CellFormatSpecification CreateCellFormat(string name, CellFormatSpecification sourceFormat)
            {
                CellFormatSpecification newCellFormatSpecification = new CellFormatSpecification(name);

                InitializeFont(sourceFormat, newCellFormatSpecification);

                newCellFormatSpecification.HorizontalAlignment = sourceFormat.HorizontalAlignment;
                newCellFormatSpecification.NumberFormat = sourceFormat.NumberFormat;
                newCellFormatSpecification.Orientation = sourceFormat.Orientation;
                newCellFormatSpecification.VerticalAlignment = sourceFormat.VerticalAlignment;
                newCellFormatSpecification.WrapText = sourceFormat.WrapText;

                return newCellFormatSpecification;
            }

            public CellFormatSpecification CreateCellFormat(string name, int fontSize)
            {
                CellFormatSpecification cellFormatSpecification = new CellFormatSpecification(name);

                InitializeFont(this.startRange.Font, cellFormatSpecification);

                cellFormatSpecification.Font.Size = fontSize;

                return cellFormatSpecification;
            }

            public CellFormatSpecification CreateCellFormat(string name, int fontSize, XlHAlign horizontalAlignment)
            {
                 CellFormatSpecification cellFormatSpecification = new CellFormatSpecification(name);

                InitializeFont(this.startRange.Font, cellFormatSpecification);

                cellFormatSpecification.Font.Size = fontSize;
                cellFormatSpecification.HorizontalAlignment = horizontalAlignment;

                return cellFormatSpecification;
            }

            public CellFormatSpecification CreateCellFormat(string name, int fontSize, Excel.MakeBold makeBold)
            {
                CellFormatSpecification cellFormatSpecification = new CellFormatSpecification(name);

                InitializeFont(this.startRange.Font, cellFormatSpecification);

                cellFormatSpecification.Font.Size = fontSize;
                cellFormatSpecification.Font.Bold = makeBold;

                return cellFormatSpecification;
            }

            public CellFormatSpecification CreateCellFormat(string name, int fontSize, Excel.MakeBold makeBold, XlOrientation orientation)
            {
                CellFormatSpecification cellFormatSpecification = new CellFormatSpecification(name);

                InitializeFont(this.startRange.Font, cellFormatSpecification);

                cellFormatSpecification.Font.Size = fontSize;
                cellFormatSpecification.Font.Bold = makeBold;
                cellFormatSpecification.Orientation = orientation;

                return cellFormatSpecification;
            }
        }
    }
}