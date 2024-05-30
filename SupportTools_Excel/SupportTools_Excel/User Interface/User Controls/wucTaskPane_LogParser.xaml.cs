using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;

using XlHlp = VNC.AddinHelper.Excel;
//using System.Windows.Forms;

namespace SupportTools_Excel.User_Interface.User_Controls
{
    /// <summary>
    /// Interaction logic for wucTaskPane_LogParser.xaml
    /// </summary>
    public partial class wucTaskPane_LogParser : UserControl
    {
        #region Enums, Fields, Properties

        enum ActivateMatch
        {
            Yes,
            No
        }

        enum LogColumns
        {
            ID = 1,
            TimeStamp = 2,
            Category = 3,
            Priority = 4,
            Severity = 5,
            EventID = 6,
            ProcessName = 7,
            ProcessID = 8,
            ThreadName = 9,
            ThreadID = 10,
            UserName = 11,
            ClassNameMethodName = 12,
            Duration = 13,
            Message = 14
        }

        enum SignalRColumns
        {
            ID = 1,
            TimeStamp = 2,
            Category = 3,
            Priority = 4,
            //Severity = 5,
            //EventID = 6,
            //ProcessName = 7,
            ProcessID = 5,
            ThreadName = 6,
            ThreadID = 7,
            //UserName = 11,
            ClassNameMethodName = 8,
            Duration = 9,
            Message = 10
        }

        #endregion

        #region Initialization and Load

        public wucTaskPane_LogParser()
        {
            InitializeComponent();
            LoadControlContents();
        }

        #endregion

        #region Event Handlers

        private void btnExtractNamedRange_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                ExtractNamedRegion();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.ToString());
            }
        }

        private void btnClearTableFilters_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                ClearTableFilters();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.ToString());
            }
        }

        void PopulateExtractRowValues()
        {
            XlHlp.DisplayInWatchWindow(string.Format("{0}",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));

            Microsoft.Office.Interop.Excel.Application app = Globals.ThisAddIn.Application;
            Worksheet ws = (Worksheet)app.ActiveSheet;

            Int32 currentRow = app.ActiveCell.Row;

            string message = (string)((Range)Globals.ThisAddIn.Application.Cells[currentRow, LogColumns.Message]).Value;
            string method = (string)((Range)Globals.ThisAddIn.Application.Cells[currentRow, LogColumns.ClassNameMethodName]).Value;

            Int32 startRow = currentRow;
            Int32 endRow = currentRow;
            teEndRow.Text = currentRow.ToString();
            teStartRow.Text = currentRow.ToString();
            teSheetName.Text = method;

            if (message.StartsWith("Enter"))
            {
                endRow = FindMatchingExitForward(ActivateMatch.No);

                teEndRow.Text = startRow == endRow ? "??" : endRow.ToString();
            }
            else if (message.StartsWith("Exit"))
            {
                startRow = FindMatchingEnterBackward(ActivateMatch.No);

                teStartRow.Text = startRow == endRow ? "??" : startRow.ToString();
            }
            else
            {
                MessageBox.Show("Message must start with Enter/Exit");
            }
        }

        private void btnPopulateExtractRowValues_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                PopulateExtractRowValues();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void btnClearTableColumnFilter_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                ClearTableColumnFilter();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.ToString());
            }
        }

        private void btnExtractRows_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                ExtractRows();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.ToString());
            }
        }

        private void btnFilterBySelectedFields_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                FilterBySelectedFields();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.ToString());
            }
        }
        private void btnFindLongestTimeBackward_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                FindLongestTimeBackward();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.ToString());
            }
        }

        private void btnFindLongestTimeForward_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                FindLongestTimeForward();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.ToString());
            }
        }

        private void btnFindMatchBackward_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                FindMatchBackward();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.ToString());
            }
        }

        private void btnFindMatchForward_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                FindMatchForward();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.ToString());
            }
        }

        private void btnFindMatchingEnter_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                FindMatchingEnterBackward();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.ToString());
            }
        }

        private void btnFindMatchingExit_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                FindMatchingExitForward();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void btnGroupMatchingEnter_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                GroupMatchingEnter();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void btnGroupMatchingExit_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                GroupMatchingExit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void btnLoadLogFile_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                XlHlp.ScreenUpdatesOff();

                LoadLogFile();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                XlHlp.ScreenUpdatesOn(true);
            }
        }

        private void btnMoveToKeyRow_Click(object sender, RoutedEventArgs e)
        {
            // User has double clicked on teKeyRow which uses the RowIDs from the table
            // This routine uses Excel Rows.  Adjust.

            const Int32 rowOffset = 5;  // data rows start on row 6

            try
            {
                Int32 keyRow = 0;
                
                if (int.TryParse(teKeyRow.Text, out keyRow))
                {
                    MoveToKeyRow(keyRow + rowOffset);
                }
                else
                {
                    MessageBox.Show("Cannot convert column 1 to row");
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.ToString());
            }
        }

        #endregion

        #region Private Methods

        void ClearTableColumnFilter()
        {
            XlHlp.DisplayInWatchWindow(string.Format("{0}",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));

            Microsoft.Office.Interop.Excel.Application app = Globals.ThisAddIn.Application;
            Worksheet ws = (Worksheet)app.ActiveSheet;
            string listName = "";
            Int32 currentColumn = app.ActiveCell.Column;

            if (ws.ListObjects.Count > 1)
            {
                System.Windows.Forms.MessageBox.Show("Multiple Tables on WorkSheet, aborting.");
            }
            else
            {
                // Odd that we have to use 1 and not 0 for the index value.

                listName = ws.ListObjects[1].Name;
                ws.ListObjects[listName].Range.AutoFilter(Field: currentColumn);
            }
        }

        void ClearTableFilters()
        {
            XlHlp.DisplayInWatchWindow(string.Format("{0}",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));

            Microsoft.Office.Interop.Excel.Application app = Globals.ThisAddIn.Application;
            Worksheet ws = (Worksheet)app.ActiveSheet;

            ws.ShowAllData();

            //string listName = "";

            //if (ws.ListObjects.Count > 1)
            //{
            //    System.Windows.Forms.MessageBox.Show("Multiple Tables on WorkSheet, aborting.");
            //}
            //else
            //{
            //    // Odd that we have to use 1 and not 0 for the index value.

            //    listName = ws.ListObjects[1].Name;
            //    ws.ListObjects[listName].Sort.SortFields.Clear();
            //}
        }

        private static string CopyRowsToNewWorkSheet(Worksheet sourceWs, int startRow, int endRow, string sheetName)
        {
            sheetName = XlHlp.SafeSheetName(sheetName);
            Worksheet destinationWS = XlHlp.NewWorksheet(sheetName, afterSheetName: "LAST");

            XlHlp.XlLocation insertAt = new XlHlp.XlLocation(destinationWS, row: 5, column: 1);

            insertAt.MarkStart(XlHlp.MarkType.GroupTable);

            CreateLogWorkSheetHeader(insertAt);

            insertAt.AddRowX();

            // HACK(crhodes)
            // 
            // There are 14 columns normally, grab five more just in case
            // We can do better than this by finding the last active column. 
            // Extra columns can appear when an exception messes up the formatting of the log file.

            Range sourceRng;

            sourceRng = sourceWs.Range[sourceWs.Cells[startRow, 1], sourceWs.Cells[endRow, 19]];  //.Cells[startRow, 1];

            sourceRng.Copy();

            insertAt.GetCurrentRange().PasteSpecial();

            destinationWS.Activate();

            // Things come over messed up.  Unfreeze
            //Globals.ThisAddIn.Application.ActiveWindow.FreezePanes = false;

            Range rng;
            rng = (Range)destinationWS.Cells[5, 1];

            Int32 lastRow = rng.SpecialCells(XlCellType.xlCellTypeLastCell).Row;
            Int32 lastColumn = rng.SpecialCells(XlCellType.xlCellTypeLastCell).Column;

            XlHlp.DisplayInWatchWindow(string.Format("{0}() row:({1}) col:({2})",
                            MethodBase.GetCurrentMethod().Name, lastRow, lastColumn));

            insertAt.SetRow(lastRow + 1);
            insertAt.SetColumn(lastColumn);

            insertAt.MarkEnd(VNC.AddinHelper.Excel.MarkType.GroupTable, string.Format("tbl_{0}", destinationWS.Name));

            //rng = destinationWS.Columns["E:K"];
            //rng.Columns.Group();
            ((Worksheet)destinationWS.Columns["E:K"]).Columns.Group();

            //destinationWS.Cells[6, 14].Select();

            //destinationWS.Application.ActiveWindow.FreezePanes = true;

            // TODO(crhodes)
            // Figure out what to do to make outlining work.

            //sourceWs.Outline.SummaryColumn = XlSummaryColumn.xlSummaryOnLeft;
            //sourceWs.Outline.SummaryRow = XlSummaryRow.xlSummaryAbove;
            return sheetName;
        }

        //Workbooks.OpenText Filename:= _
        //"C:\Customers\Aston Martin\AML_Logs\EaseWorks_0228.log", Origin:=437, _
        //startRow:=1, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, _
        //ConsecutiveDelimiter:=False, Tab:=False, Semicolon:=False, Comma:=False _
        //, Space:=False, Other:=True, OtherChar:="|", FieldInfo:=Array(Array(1, 1 _
        //), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1), _
        //Array(9, 1), Array(10, 1), Array(11, 1), Array(12, 1), Array(13, 1)), _
        //TrailingMinusNumbers:=True

        //    With ActiveSheet.QueryTables.Add(Connection:= _
        //    "TEXT;C:\Customers\Aston Martin\AML_Logs\AMLViewEase.log", Destination:=Range _
        //    ("$B$5"))
        //    .CommandType = 0
        //    .Name = "AMLViewEase"
        //    .FieldNames = True
        //    .RowNumbers = False
        //    .FillAdjacentFormulas = False
        //    .PreserveFormatting = True
        //    .RefreshOnFileOpen = False
        //    .RefreshStyle = xlInsertDeleteCells
        //    .SavePassword = False
        //    .SaveData = True
        //    .AdjustColumnWidth = True
        //    .RefreshPeriod = 0
        //    .TextFilePromptOnRefresh = False
        //    .TextFilePlatform = 437
        //    .TextFileStartRow = 1
        //    .TextFileParseType = xlDelimited
        //    .TextFileTextQualifier = xlTextQualifierDoubleQuote
        //    .TextFileConsecutiveDelimiter = False
        //    .TextFileTabDelimiter = False
        //    .TextFileSemicolonDelimiter = False
        //    .TextFileCommaDelimiter = False
        //    .TextFileSpaceDelimiter = False
        //    .TextFileOtherDelimiter = "|"
        //    .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
        //    .TextFileTrailingMinusNumbers = True
        //    .Refresh BackgroundQuery:=False
        //End With

        private void CreateLogWorkSheet(string fileName)
        {
            XlHlp.DisplayInWatchWindow(string.Format("{0}()",
                MethodBase.GetCurrentMethod().Name));

            //string sheetName = XlHlp.SafeSheetName("Log-RenameMe");
            //Worksheet ws = XlHlp.NewWorksheet(sheetName, beforeSheetName: "FIRST");

            //XlHlp.XlLocation insertAt = new XlHlp.XlLocation(ws, row: 5, column: 2);

            Globals.ThisAddIn.Application.Workbooks.OpenText(Filename: fileName, StartRow: 1,
                DataType: XlTextParsingType.xlDelimited,
                Tab: false, Semicolon: false, Comma: false, Space: false,
                Other: true, OtherChar: "|");

            Worksheet ws = (Worksheet)Globals.ThisAddIn.Application.ActiveSheet;

            // Rows("1:1").Select
            // Selection.Insert Shift:= xlDown, CopyOrigin:= xlFormatFromLeftOrAbove
            // Columns("A:A").Select
            // Selection.Insert Shift:= xlToRight, CopyOrigin:= xlFormatFromLeftOrAbove

            Range rng = (Range)ws.Rows["1:1"];

            rng.Insert(Shift: XlInsertShiftDirection.xlShiftDown, CopyOrigin: XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
            rng.Insert(Shift: XlInsertShiftDirection.xlShiftDown, CopyOrigin: XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
            rng.Insert(Shift: XlInsertShiftDirection.xlShiftDown, CopyOrigin: XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
            rng.Insert(Shift: XlInsertShiftDirection.xlShiftDown, CopyOrigin: XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
            rng.Insert(Shift: XlInsertShiftDirection.xlShiftDown, CopyOrigin: XlInsertFormatOrigin.xlFormatFromLeftOrAbove);

            rng = ((Range)ws.Columns["A:A"]);

            rng.Insert(Shift: XlInsertShiftDirection.xlShiftToRight, CopyOrigin: XlInsertFormatOrigin.xlFormatFromLeftOrAbove);

            XlHlp.XlLocation insertAt = new XlHlp.XlLocation(ws, row: 5, column: 1);

            insertAt.MarkStart(XlHlp.MarkType.GroupTable);

            CreateLogWorkSheetHeader(insertAt);

            // Do some formatting.  May want to pull this out into a separate method, FormatSheetColumns() if do more

            rng = (Range)ws.Columns["B:B"];

            rng.NumberFormat = "m/d/yyyy h:mm:ss.000";

            //insertAt.ClearOffsets();

            rng = (Range)ws.Cells[5, 1];

            Int32 lastRow = rng.SpecialCells(XlCellType.xlCellTypeLastCell).Row;
            Int32 lastColumn = rng.SpecialCells(XlCellType.xlCellTypeLastCell).Column;

            XlHlp.DisplayInWatchWindow(string.Format("{0}() row:({1}) col:({2})",
                            MethodBase.GetCurrentMethod().Name, lastRow, lastColumn));


            // TODO(crhodes)
            // This would be a good thing to add to ExcelUtilities

            // Add a row number to table so can resort if necessary.  Table starts at 5, first data row at 6

            for (int i = 6, rowNumber = 1; i <= lastRow; i++, rowNumber++)
            {
                ((Range)ws.Cells[i, 1]).Value = rowNumber;
            }

            insertAt.SetRow(lastRow + 1);

            insertAt.MarkEnd(VNC.AddinHelper.Excel.MarkType.GroupTable, string.Format("tbl_{0}", ws.Name));

            rng = (Range)ws.Columns["E:K"];
            rng.Columns.Group();

            ws.Outline.SummaryColumn = XlSummaryColumn.xlSummaryOnLeft;
            ws.Outline.SummaryRow = XlSummaryRow.xlSummaryAbove;


            ((Range)ws.Cells[6, 14]).Select();

            ws.Application.ActiveWindow.FreezePanes = true;

            SaveLogFile(ws.Application.ActiveWorkbook, fileName);

            //Selection.Columns.Group
            //Selection.Columns.Group
            //Range("D3").Select
            //ActiveSheet.Outline.ShowLevels RowLevels:= 0, ColumnLevels:= 3
            //ActiveSheet.Outline.ShowLevels RowLevels:= 0, ColumnLevels:= 2
            //ActiveSheet.Outline.ShowLevels RowLevels:= 0, ColumnLevels:= 1
            //With ActiveSheet.Outline
            //    .AutomaticStyles = False
            //    .SummaryRow = xlAbove
            //    .SummaryColumn = xlLeft
            //End With
        }

        private static void CreateLogWorkSheetHeader(XlHlp.XlLocation insertAt)
        {
            XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(),  8, "ID");
            XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 20, "TimeStamp");
            XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 15, "Category");
            XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 10, "Priority");
            XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 10, "Severity");

            // TODO(crhodes)
            // Add a UI Config for supporting EventID

            //XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 10, "EventID");

            XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 40, "ProcessName");
            XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 12, "ProcessID");
            XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 14, "ThreadName");
            XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 12, "ThreadID");
            XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 15, "UserName");
            XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 50, "ClassName.MethodName");
            XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 20, "Duration");
            XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), 100,"Message");
        }

        void ExtractNamedRegion()
        {
            XlHlp.DisplayInWatchWindow(string.Format("{0}",
                            System.Reflection.MethodInfo.GetCurrentMethod().Name));

            Microsoft.Office.Interop.Excel.Application app = Globals.ThisAddIn.Application;
            Worksheet sourceWs = (Worksheet)app.ActiveSheet;


            Int32 startRow = 0;
            Int32 endRow = 0;
            string sheetName = "";

            string namedRange = teNamedRange.Text;

            FindNamedRangeRows(namedRange, "Start:", "End:", out startRow, out endRow);

            if (ceUseRowNumbers.IsChecked == true)
            {
                sheetName = string.Format("{0}_{1}", startRow, endRow);
            }
            else if ("" == (sheetName = teSheetNameNamedRange.Text))
            {
                System.Windows.Forms.MessageBox.Show("SheetName not valid");
                return;
            }

            sheetName = CopyRowsToNewWorkSheet(sourceWs, startRow, endRow, sheetName);

            teStartRowNamedRange.Clear();
            teEndRowNamedRange.Clear();
            teSheetNameNamedRange.Clear();
        }

        private void ExtractRows()
        {
            XlHlp.DisplayInWatchWindow(string.Format("{0}",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));

            Microsoft.Office.Interop.Excel.Application app = Globals.ThisAddIn.Application;
            Worksheet sourceWs = (Worksheet)app.ActiveSheet;
            Range sourceRng;
            Range rng;

            Int32 startRow = 0;
            Int32 endRow = 0;
            string sheetName = "";

            if (!Int32.TryParse(teStartRow.Text, out startRow))
            {
                System.Windows.Forms.MessageBox.Show("StartRow not valid");
                return;
            }

            if (!Int32.TryParse(teEndRow.Text, out endRow))
            {
                System.Windows.Forms.MessageBox.Show("EndRow not valid");
                return;
            }

            if (ceUseRowNumbers.IsChecked == true)
            {
                sheetName = string.Format("{0}_{1}", startRow, endRow);
            }
            else if ("" == (sheetName = teSheetName.Text))
            {
                System.Windows.Forms.MessageBox.Show("SheetName not valid");
                return;
            }

            sheetName = CopyRowsToNewWorkSheet(sourceWs, startRow, endRow, sheetName);

            teStartRow.Clear();
            teEndRow.Clear();
            teSheetName.Clear();
        }

        void FilterBySelectedFields()
        {
            XlHlp.DisplayInWatchWindow(string.Format("{0}",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));

            Microsoft.Office.Interop.Excel.Application app = Globals.ThisAddIn.Application;
            Worksheet ws = (Worksheet)app.ActiveSheet;
            Range selection = (Range)app.Selection;
            string listName = "";
            string[] criteria;
            List<string> criteria2 = new List<string>();

            Boolean firstItem = true;

            if (null == selection)
            {
                System.Windows.Forms.MessageBox.Show("Must select one or more values");
            }

            criteria = new string[selection.Count];

            int i = 0;

            foreach (Range item in selection)
            {
                XlHlp.DisplayInWatchWindow(string.Format("{0}", item.Value));
                criteria[i++] = item.Value.ToString();
                criteria2.Add(item.Value.ToString());
            }

            foreach (ListObject listObject in ws.ListObjects)
            {
                XlHlp.DisplayInWatchWindow(string.Format("{0}", listObject.Name));
            }

            if (ws.ListObjects.Count > 1)
            {
                System.Windows.Forms.MessageBox.Show("Multiple Tables on WorkSheet, aborting.");
            }
            else
            {
                // Why doesn't ListObjects[0] work?

                //try
                //{
                //    ListObject lo0 = ws.ListObjects[0];
                //}
                //catch (Exception ex)
                //{

                //}
                //try
                //{
                //    ListObject lo1 = ws.ListObjects[1];
                //}
                //catch (Exception ex)
                //{

                //}

                listName = ws.ListObjects[1].Name;

                ws.ListObjects[listName].Range.AutoFilter(Field: 12, Criteria1: criteria, Operator: XlAutoFilterOperator.xlFilterValues);
            }
        }

        void FindLongestTimeBackward()
        {
            XlHlp.DisplayInWatchWindow(string.Format("{0}",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));

            Microsoft.Office.Interop.Excel.Application app = Globals.ThisAddIn.Application;
            Worksheet ws = (Worksheet)app.ActiveSheet;
            Int32 lastRow = app.ActiveCell.Row;
            Int32 currentColumn = app.ActiveCell.Column;
            Int32 durationColumn = currentColumn + 1;

            Range startingCell = (Range)Globals.ThisAddIn.Application.Cells[lastRow, durationColumn];
            //Range startingCell = Globals.ThisAddIn.Application.Cells[lastRow, LogColumns.Duration];

            Double totalTime = (Double)startingCell.Value;
            Double subTime = 0.0;

            Int32 firstRow = FindMatchingEnterBackward(ActivateMatch.No);

            // TODO(crhodes)
            // Search backwards up to the matching enter for the longest time and put the cursor there

            Int32 matchingRow = lastRow;

            for (int row = lastRow - 1; row > firstRow; row--)
            {
                Double? rowTime = (Double)((Range)ws.Cells[row, durationColumn]).Value;
                //Double? rowTime = ws.Cells[row, LogColumns.Duration].Value;

                if (null != rowTime)
                {
                    if (rowTime > subTime)
                    {
                        subTime = (double)rowTime;
                        matchingRow = row;
                    }
                }
            }

            ((Range)ws.Cells[matchingRow, currentColumn]).Activate();
            UpdateTimeControls(totalTime.ToString(), subTime.ToString());
        }

        void FindLongestTimeForward()
        {
            XlHlp.DisplayInWatchWindow(string.Format("{0}",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));

            Microsoft.Office.Interop.Excel.Application app = Globals.ThisAddIn.Application;
            Worksheet ws = (Worksheet)app.ActiveSheet;
            Int32 startingRow = app.ActiveCell.Row;
            Int32 currentColumn = app.ActiveCell.Column;
            Int32 durationColumn = currentColumn + 1;

            Range startingCell = (Range)Globals.ThisAddIn.Application.Cells[startingRow, durationColumn];

            string totalTime = string.Empty;
            string subTime = string.Empty;

            // TODO(crhodes)
            // go the the exit time

            totalTime = GetMatchingExitDuration(startingCell);

            // TODO(crhodes)
            // Search forwards up to the matching exit for the longest time and put the cursor there

            UpdateTimeControls(totalTime, subTime);
        }

        void FindMatchBackward()
        {
            Microsoft.Office.Interop.Excel.Application app = Globals.ThisAddIn.Application;
            Worksheet ws = (Worksheet)app.ActiveSheet;
            Range activeCell = Globals.ThisAddIn.Application.ActiveCell;
            Range currentCell = activeCell;
            //Range currentCell = ws.Cells[activeCell.Row, LogColumns.ClassNameMethodName];


            string pattern = (string)currentCell.Value;
            Int32 currentRow = currentCell.Row;
            Int32 currentColumn = currentCell.Column;
            //Int32 foundRow;
            //Boolean foundCorrectContents = false;
            //string contentValue = "";

            app.Cells.Find(What: pattern, After: currentCell,
                           LookIn: XlFindLookIn.xlFormulas, LookAt: XlLookAt.xlWhole, SearchOrder: XlSearchOrder.xlByRows,
                           SearchDirection: XlSearchDirection.xlPrevious,
                           MatchCase: false, SearchFormat: false).Activate();
        }

        void FindMatchForward()
        {
            Microsoft.Office.Interop.Excel.Application app = Globals.ThisAddIn.Application;
            Worksheet ws = (Worksheet)app.ActiveSheet;
            Range activeCell = Globals.ThisAddIn.Application.ActiveCell;
            Range currentCell = activeCell;
            //Range currentCell = ws.Cells[activeCell.Row, LogColumns.ClassNameMethodName];

            string pattern = (string)currentCell.Value;
            Int32 currentRow = currentCell.Row;
            Int32 currentColumn = currentCell.Column;
            //Int32 foundRow;
            //Boolean foundCorrectContents = false;
            //string contentValue = "";

            app.Cells.Find(What: pattern, After: currentCell,
                            LookIn: XlFindLookIn.xlFormulas, LookAt: XlLookAt.xlWhole, SearchOrder: XlSearchOrder.xlByRows,
                            SearchDirection: XlSearchDirection.xlNext,
                            MatchCase: false, SearchFormat: false).Activate();
        }

        private Int32 FindMatchingEnterBackward(ActivateMatch activateMatch = ActivateMatch.Yes)
        {
            Microsoft.Office.Interop.Excel.Application app = Globals.ThisAddIn.Application;
            Worksheet ws = (Worksheet)app.ActiveSheet;
            Range activeCell = Globals.ThisAddIn.Application.ActiveCell;
            Range currentCell = activeCell;
            //Range currentCell = ws.Cells[activeCell.Row, LogColumns.ClassNameMethodName];

            string pattern = (string)currentCell.Value;
            Int32 currentRow = currentCell.Row;
            Int32 currentColumn = currentCell.Column;
            Int32 foundRow = currentRow;
            Int32 activeRow;
            Boolean foundCorrectContents = false;
            string contentValue = "";

            app.Cells.Find(What: pattern, After: currentCell,
                            LookIn: XlFindLookIn.xlFormulas, LookAt: XlLookAt.xlWhole, SearchOrder: XlSearchOrder.xlByRows,
                            SearchDirection: XlSearchDirection.xlPrevious,
                            MatchCase: false, SearchFormat: false).Activate();

            do
            {
                activeRow = (int)Globals.ThisAddIn.Application.ActiveCell.Row;
                contentValue = (string)Globals.ThisAddIn.Application.ActiveCell.Offset[0, 2].Value;  // Message is two columns to the right.

                if (contentValue != null && contentValue.Contains("Enter"))
                {
                    foundCorrectContents = true;
                    foundRow = activeRow;
                }
                else
                {
                    if (activeRow >= currentRow)
                    {   // Avoid looping if not found
                        foundRow = currentRow;
                        break;
                    }

                    app.Cells.FindPrevious(After: app.ActiveCell).Activate();                     
                }
            } while (!(foundCorrectContents));

            //if (activateMatch == ActivateMatch.Yes)
            //{
            //    matchRange.Activate();
            //}

            return foundRow;
        }

        private Int32 FindMatchingExitForward(ActivateMatch activateMatch = ActivateMatch.Yes)
        {
            Microsoft.Office.Interop.Excel.Application app = Globals.ThisAddIn.Application;
            Worksheet ws = (Worksheet)app.ActiveSheet;
            Range activeCell = Globals.ThisAddIn.Application.ActiveCell;
            Range currentCell = activeCell;
            //Range currentCell = ws.Cells[activeCell.Row, LogColumns.ClassNameMethodName];

            string pattern = (string)currentCell.Value;
            Int32 currentRow = currentCell.Row;
            Int32 currentColumn = currentCell.Column;
            Int32 foundRow = currentRow;
            Int32 activeRow;
            Boolean foundCorrectContents = false;
            string contentValue = "";

            app.Cells.Find(What: pattern, After: currentCell,
                            LookIn: XlFindLookIn.xlFormulas, LookAt: XlLookAt.xlWhole, SearchOrder: XlSearchOrder.xlByRows,
                            SearchDirection: XlSearchDirection.xlNext,
                            MatchCase: false, SearchFormat: false).Activate();

            do
            {
                activeRow = Globals.ThisAddIn.Application.ActiveCell.Row;
                contentValue = (string)app.ActiveCell.Offset[0, 2].Value;  // Message is two columns to the right.

                if (contentValue != null && contentValue.Contains("Exit"))
                {
                    foundCorrectContents = true;
                    foundRow = activeRow;
                }
                else
                {
                    if (activeRow <= currentRow)
                    {   // Avoid looping if not found
                        foundRow = currentRow;
                        break;
                    }
                    app.Cells.FindNext(After: app.ActiveCell).Activate();
                }
            } while (!(foundCorrectContents));

            return foundRow;
        }

        Boolean FindNamedRangeRows(string namedRange, string beginPattern, string endPattern, out int startRow, out int endRow)
        {
            startRow = 0;
            endRow = 0;
            Boolean foundNamedRange = false;

            // TODO(crhodes)
            // Search for namedRange in the log entries for rows that begin with
            // <beginPattern><namdedRange>
            // and end with
            // <endPattern><namedRange>
            // e.g. Start:TestOne    End:TestOne
            //
            // They should be at Log.Info level but that may change in the future

            Microsoft.Office.Interop.Excel.Application app = Globals.ThisAddIn.Application;
            Worksheet ws = (Worksheet)app.ActiveSheet;
            Range startingCell = (Range)Globals.ThisAddIn.Application.Cells[5, LogColumns.Message];

            string pattern = beginPattern + namedRange;
            Int32 currentRow = startingCell.Row;
            Int32 currentColumn = startingCell.Column;
            //Int32 foundRow;
            //Boolean foundCorrectContents = false;
            //string contentValue = "";
            Range foundRangeStart = null;
            Range foundRangeEnd = null;

            foundRangeStart = app.Cells.Find(What: pattern, After: startingCell,
                           LookIn: XlFindLookIn.xlFormulas, LookAt: XlLookAt.xlWhole, SearchOrder: XlSearchOrder.xlByRows,
                           SearchDirection: XlSearchDirection.xlPrevious,
                           MatchCase: false, SearchFormat: false);

            if (null != foundRangeStart)
            {
                startRow = foundRangeStart.Row;
                pattern = endPattern + namedRange;

                foundRangeEnd = app.Cells.Find(What: pattern, After: startingCell,
                    LookIn: XlFindLookIn.xlFormulas, LookAt: XlLookAt.xlWhole, SearchOrder: XlSearchOrder.xlByRows,
                    SearchDirection: XlSearchDirection.xlPrevious,
                    MatchCase: false, SearchFormat: false);

                if (null != foundRangeEnd)
                {
                    endRow = foundRangeEnd.Row;
                    foundNamedRange = true;

                    ((Range)ws.Cells[startRow, LogColumns.Message]).Activate();
                }
            }

            return foundNamedRange;
        }

        string GetCurrentTable(Worksheet ws)
        {
            string activeTableName = string.Empty;
            // TODO(crhodes)
            // Figure out how to find the table (ListObject) the current cell is in

            activeTableName = "tbl_" + ws.Name;
            return activeTableName;
        }
        private bool GetDisplayOrientation()
        {
            return (bool)ceOrientOutputVertically.IsChecked;
        }

        string GetMatchingExitDuration(Range range)
        {
            return "99.9";
        }

        private void GroupMatchingEnter()
        {
            Microsoft.Office.Interop.Excel.Application app = Globals.ThisAddIn.Application;
            Worksheet ws = (Worksheet)app.ActiveSheet;
            Range currentCell = Globals.ThisAddIn.Application.ActiveCell;

            string pattern = (string)currentCell.Value;
            Int32 currentRow = currentCell.Row;
            Int32 currentColumn = currentCell.Column;
            Int32 foundRow;

            foundRow = FindMatchingEnterBackward();

            if (currentRow != foundRow)
            {
                string groupRange = string.Format("{0}:{1}", currentRow - 1, foundRow + 1);
                ((Range)ws.Rows[groupRange]).Group();

                SetOutlineStyle();
            }
        }

        private void GroupMatchingExit()
        {
            Microsoft.Office.Interop.Excel.Application app = Globals.ThisAddIn.Application;
            Worksheet ws = (Worksheet)app.ActiveSheet;
            Range currentCell = Globals.ThisAddIn.Application.ActiveCell;

            string pattern = (string)currentCell.Value;
            Int32 currentRow = currentCell.Row;
            Int32 currentColumn = currentCell.Column;
            Int32 foundRow;

            foundRow = FindMatchingExitForward();

            if (currentRow != foundRow)
            {
                string groupRange = string.Format("{0}:{1}", currentRow + 1, foundRow - 1);
                ((Range)ws.Rows[groupRange]).Group();

                ((Range)app.Cells[currentRow, 1]).Activate();

                SetOutlineStyle();
            }
        }

        private void LoadControlContents()
        {
            try
            {

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        void LoadLogFile()
        {
            XlHlp.DisplayInWatchWindow(string.Format("{0}()",
                MethodBase.GetCurrentMethod().Name));

            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.Filter = "Log files (*.log)|*.log|All files (*.*)|*.*";
            openFileDialog1.FileName = "";

            if (true == openFileDialog1.ShowDialog())
            {
                string fileName = openFileDialog1.FileName;

                CreateLogWorkSheet(fileName);
            }
        }

        private void MoveToKeyRow(Int32 keyRow)
        {
            XlHlp.DisplayInWatchWindow(string.Format("{0}",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));

            Microsoft.Office.Interop.Excel.Application app = Globals.ThisAddIn.Application;
            Worksheet ws = (Worksheet)app.ActiveSheet;

            string currentTableName = GetCurrentTable(ws);

            try
            {
                ListObject currentTable = ws.ListObjects[currentTableName];

                // Put the rows back in ID order.

                //currentTable.Range.Sort(Key1: currentTable.ListColumns[LogColumns.ID].Range);

                currentTable.Sort.SortFields.Clear();

                currentTable.Sort.SortFields.Add(
                    Key: currentTable.ListColumns[1].Range,
                    SortOn: XlSortOn.xlSortOnValues,
                    Order: XlSortOrder.xlAscending,
                    DataOption: XlSortDataOption.xlSortNormal);

                currentTable.Sort.Apply();

                // And move to the key row
                Range targetCell = (Range)ws.Cells[keyRow, LogColumns.Duration];
                targetCell.Activate();
                targetCell.Interior.Pattern = XlPattern.xlPatternSolid;
                targetCell.Interior.Color = System.Drawing.Color.Yellow;
                //targetCell.Interior.Color = Colors.Yellow;
                targetCell.Interior.PatternColor = System.Drawing.Color.Yellow;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        void SaveLogFile(Workbook wb, string fileName)
        {
            XlHlp.DisplayInWatchWindow(string.Format("{0}()",
                MethodBase.GetCurrentMethod().Name));

            SaveFileDialog saveFileDialog = new SaveFileDialog();

            string strOutputFile = null;

            FileInfo fileInfo = new FileInfo(fileName);

            try
            {
                saveFileDialog.FileName = fileInfo.Name;

                saveFileDialog.InitialDirectory = fileInfo.DirectoryName;

                if (false == saveFileDialog.ShowDialog())
                {
                    return;
                }
                else
                {
                    strOutputFile = saveFileDialog.FileName;
                }
                if (string.IsNullOrEmpty(strOutputFile))
                {
                    return;
                }
                Globals.ThisAddIn.Application.ActiveWorkbook.SaveAs(strOutputFile, XlFileFormat.xlWorkbookNormal);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void SetOutlineStyle()
        {
            //Worksheet ws = Globals.ThisAddIn.Application.ActiveSheet;
            //ws.Outline.AutomaticStyles = false;
            //ws.Outline.SummaryRow = XlSummaryRow.xlSummaryAbove;
            //ws.Outline.SummaryColumn = XlSummaryColumn.xlSummaryOnRight;
        }

        private void teEndRow_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application app = Globals.ThisAddIn.Application;
            teEndRow.Text = app.ActiveCell.Row.ToString();
        }

        private void teKeyRow_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application app = Globals.ThisAddIn.Application;
            Worksheet ws = (Worksheet)app.ActiveSheet;
            Range activeCell = app.ActiveCell;

            string targetRow = ((Range)ws.Cells[activeCell.Row, LogColumns.ID]).Value.ToString();

            Int32 keyRow = 0;

            Int32.TryParse(targetRow, out keyRow);

            ((DevExpress.Xpf.Editors.TextEdit)sender).Text = keyRow.ToString();
            //((DevExpress.Xpf.Editors.TextEdit)sender).Text = app.ActiveCell.Row.ToString();
        }

        private void teNamedRange_LostFocus(object sender, RoutedEventArgs e)
        {
            // TODO(crhodes)
            // Maybe do the find here  could also have a button 

            string namedRange = teNamedRange.Text;
            string sheetName = XlHlp.SafeSheetName(namedRange);

            Int32 startRow = 0;
            Int32 endRow = 0;

            if (FindNamedRangeRows(namedRange, "Start:", "End:", out startRow, out endRow))
            {
                teStartRowNamedRange.Text = startRow.ToString();
                teEndRowNamedRange.Text = endRow.ToString();

                teSheetNameNamedRange.Text = sheetName;
            }
            else
            {
                teStartRowNamedRange.Text = "??";
                teEndRowNamedRange.Text = "??";

                teSheetNameNamedRange.Text = "<Range Not Found>";
            }
        }

        private void teStartRow_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application app = Globals.ThisAddIn.Application;
            teStartRow.Text = app.ActiveCell.Row.ToString();
        }


        void UpdateTimeControls(string totalTime, string subTime)
        {
            lilTotalTime.Content = totalTime;
            lilSubTime.Content = subTime;

        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            //wucSQLInstance_Picker1.ControlChanged += WucSQLInstance_Picker1_ControlChanged;
            //wucTFSProvider_Picker.ControlChanged += tfsProvider_Picker1_ControlChanged;
        }

        private bool ValidUISelections()
        {
            //if (cbeTeamProjectCollections.SelectedText.Length > 0)
            //{
            return true;
            //}
            //else
            //{
            //    MessageBox.Show("Must Select Team Project Collection first", "UI Selection Incomplete");
            //    return false;
            //}
        }


        #endregion
    }
}
