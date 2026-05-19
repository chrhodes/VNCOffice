using System;
using System.Configuration;
using System.Diagnostics;
using System.Net.NetworkInformation;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.Remoting.Messaging;
using System.Security.Cryptography;
using System.Windows;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media.Media3D;
using System.Windows.Shapes;

using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;

using Newtonsoft.Json;

using Prism.Common;
using Office = Microsoft.Office.Core;

using VNC.VSTOAddIn.Excel.Domain;

namespace VNC.VSTOAddIn.Excel
{
    public class Helper
    {

        public const int cMaxSheetNameLen = 30;
        #region "Debug Constants"
        //Public Shared ScreenUpdatesToggleEnabled As Boolean = True
        #endregion

        public static object HostApp;
        public static object AddInInstance;
        // Some stuff about who started us and how they started us.
        public static string AppName;

        public static string AppVersion;

        public static short StartMode;

        public const int cHeaderFontSize = 12;
        public const int cHeaderFontSizeMedium = 10;
        public const int cHeaderFontSizeSmall = 8;

        public const string cDefaultFontName = "Calibri";
        public const int cDefaultColumnFontSize = 11;

        public static XlCalculation PriorCalculationState;

        public static bool PriorScreenUpdatingState = false;

        public enum WrapText : byte
        {
            Yes = 1,
            No = 0
        }

        public enum MakeBold : byte
        {
            Yes = 1,
            No = 0
        }

        public enum MarkType : byte
        {
            None = 0,
            Group = 1,
            Table = 2,
            GroupTable = 3
        }

        public enum SectionLocation : byte
        {
            Left = 0,
            Top = 1
        }

        public enum LabelLocation : byte
        {
            Left = 0,
            Top = 1
        }

        public enum UnderLine : byte
        {
            Yes = 1,
            No = 0
        }

        //Private Const cMODULE_NAME As String = Globals.PROJECT_NAME & ".modExcel"

        // Hard coded for Excel? 1/72 Inches.
        public const int cPointsToInch = 72;
        // Times before this
        public const int cStartHour = 8;
        // and after this are highlighted.
        public const int cEndHour = 20;

        // FIX(crhodes)
        // Not sure this is going to be set.  Need to pick one place to hold this.  I think Commmon.ExcelApplication
        public static Microsoft.Office.Interop.Excel.Application? ExcelApplication
        {
            get;
            set;
        }

        #region Main Function Routines

        #region AddColumns

        // FIXME(crhodes)
        // Why does this call a deprecated method?

        public static void AddColumnToSheet(
            Worksheet ws,
            int columnNumber,
            float columnWidth,
            bool columnWrapText,
            string columnNumberFormat,
            XlDirection shiftDirection,
            XlInsertFormatOrigin insertFormatOrigin,
            int headerRow,
            string headerTitle = "",
            CellFormatSpecification cellFormatSpecification = null)
        {
            // Insert the new column and apply things that pertain to all cells in the column
            //((Range)ws.Columns[columnNumber]).Insert(Shift: shiftDirection, CopyOrigin: insertFormatOrigin);

            //You got this error because the code is using a direct cast(Range)
            //on a COM object property(ws.Columns[columnNumber]).
            //When compiling as C# (not dynamic), this requires the C# dynamic runtime binder,
            //which is not available unless you target .NET Framework and reference Microsoft.CSharp.
            //This might be happening because your project is targeting.NET(Core / 5 +/ 6 +)
            //or does not reference Microsoft.Csharp,
            //and / or you are not using the dynamic keyword.
            //The cast(Range) triggers the runtime binder, which is missing.
            //Here's how I fixed the code:
            //I replaced the direct cast(Range) with the as Range operator.
            //This is safer and does not require the dynamic binder.
            //If the cast fails, newColumn will be null instead of throwing an exception at runtime.
            //This approach is compatible with static typing and avoids the missing binder error.

            //Range newColumn = (Range)ws.Columns[columnNumber];
            Range? newColumn = ws.Columns[columnNumber] as Range;

            if (cellFormatSpecification != null)
            {
                newColumn.Insert(Shift: shiftDirection, CopyOrigin: insertFormatOrigin);
                newColumn.WrapText = cellFormatSpecification.WrapText;
                newColumn.NumberFormat = cellFormatSpecification.NumberFormat;
                newColumn.Font.Size = cellFormatSpecification.Font.Size;
                newColumn.Font.Name = cellFormatSpecification.Font.Name;
            }

            // Pass all the rest on
            AddColumnHeaderToSheetX(ws, headerRow, columnNumber, columnWidth, headerTitle, cellFormatSpecification);
        }

        #endregion

        #region AddColumnHeaders

        public static void AddColumnHeaderToSheet(
            XlLocation insertAt,
            float columnWidth,
            string headerTitle,
            CellFormatSpecification headerFormat = null)
        {
            if (headerFormat == null)
            {
                headerFormat = insertAt.HeaderFormat;
            }

            Range header = insertAt.OffsetRange;

            // You got this error because the code is using a direct cast(Range)
            // on a COM object property(header.Worksheet.Columns[header.Column]).
            // When compiling as C# (not dynamic), this requires the C# dynamic runtime binder,
            // which is not available unless you target .NET Framework and reference Microsoft.CSharp.
            // This might be happening because your project is targeting.NET(Core / 5 +/ 6 +)
            // or does not reference Microsoft.CSharp, and / or you are not using the dynamic keyword.
            // The cast(Range) triggers the runtime binder, which is missing.
            // Here's how I fixed the code: I replaced the direct cast (Range) with the as Range operator.
            // This is safer and does not require the dynamic binder.
            // If the cast fails, headerColumn will be null instead of throwing an exception at runtime.
            // This approach is compatible with static typing and avoids the missing binder error.

            //((Range)header.Worksheet.Columns[header.Column]).ColumnWidth = columnWidth;

            Range? headerColumn = header.Worksheet.Columns[header.Column] as Range;

            if (headerColumn != null)
            {
                headerColumn.ColumnWidth = columnWidth;
            }

            //if (!string.IsNullOrEmpty(headerTitle))
            //{
            //Range header = rng;
            header.Value = headerTitle;

            if (headerFormat != null)
            {
                UpdateFormatting(header, headerFormat);
            }
            //}
            //else
            //{
            //    // Who passed in null.?
            //}
        }

        // TODO(crhodes)
        // Why would headerTitle be optional?

        [Obsolete("AddColumnHeaderToSheetX has been replaced with AddColumnHeaderToSheet which takes an XlLocation instead of a Range")]
        public static void AddColumnHeaderToSheetX(
            Range rng,
            float columnWidth,
            string headerTitle,
            CellFormatSpecification headerFormat = null)
        {
            //((Range)rng.Worksheet.Columns[rng.Column]).ColumnWidth = columnWidth;

            Range? rngColumn = rng.Worksheet.Columns[rng.Column] as Range;

            if (rngColumn != null)
            {
                rngColumn.ColumnWidth = columnWidth;
            }

            if (!string.IsNullOrEmpty(headerTitle))
            {
                Range header = rng;
                header.Value = headerTitle;

                if (headerFormat != null)
                {
                    UpdateFormatting(rng, headerFormat);
                }
            }
            else
            {
                // Who passed in null.?
            }
        }

        public static void AddColumnHeaderToSheet(
            Worksheet ws,
            int row,
            int column,
            float columnWidth,
            string headerTitle,
            CellFormatSpecification? headerFormat = null)
        {
            if (headerFormat != null)
            {
                // You got this error because the code is accessing the Name property of headerFormat.Font as if it were a string,
                // but in the Excel Interop API, Font.Name is a dynamic (actually a COM object property, not a native .NET string).
                // Using.Equals directly on a dynamic property triggers the C# dynamic runtime binder,
                // which is not available unless you target .NET Framework and reference Microsoft.CSharp.
                // This might be happening because your project is targeting.NET Core or .NET 5 + (not.NET Framework),
                // and / or does not reference Microsoft.CSharp, and / or you are not using the dynamic keyword explicitly.
                // The direct use of .Equals on a dynamic property triggers the missing runtime binder error.
                // Here's how I fixed the code:
                // I converted headerFormat.Font.Name to a string using .ToString() and then used string.Equals for the comparison.
                // This avoids the dynamic binder and works in all.NET versions,
                // ensuring type safety and compatibility with your project configuration.
                // N.B. Put original code back in as we have put Microsoft.CSharp back in
                // and it is more readable.

                if (!headerFormat.Font.Name.Equals(cDefaultFontName))
                {
                    Range currentColumn = (Range)ws.Columns[column];

                    // Assume we want the whole column to change
                    currentColumn.Font.Name = headerFormat.Font.Name;
                    currentColumn.Font.Size = headerFormat.Font.Size;
                }

                //if (!string.Equals(headerFormat.Font.Name.ToString(), cDefaultFontName, StringComparison.Ordinal))
                //{
                //    //Range currentColumn = (Range)ws.Columns[column];
                //    Range? currentColumn = ws.Columns[column] as Range;

                //    if (currentColumn is not null)
                //    { 
                //        // Assume we want the whole column to change
                //        currentColumn.Font.Name = headerFormat.Font.Name;
                //        currentColumn.Font.Size = headerFormat.Font.Size;
                //    }
                //}
            }

            if (headerFormat == null)
            {
                //h/*eaderFormat = insertAt.;*/
            }

            //AddColumnHeaderToSheetX((Range)ws.Cells[row, column], columnWidth, headerTitle, headerFormat);

            Range? rng = ws.Cells[row, column] as Range;

            if (rng is not null)
            {
                AddColumnHeaderToSheetX(rng, columnWidth, headerTitle, headerFormat);
            }
        }

        [Obsolete("AddColumnHeaderToSheetX has been replaced with AddColumnHeaderToSheet which takes an XlLocation instead of a Range")]
        public static void AddColumnHeaderToSheetX(
            Worksheet ws,
            int row,
            int column,
            float columnWidth,
            string headerTitle,
            CellFormatSpecification? headerFormat = null)
        {
            if (headerFormat != null)
            {
                if (!string.Equals(headerFormat.Font.Name.ToString(), cDefaultFontName, StringComparison.Ordinal))
                {
                    //Range currentColumn = (Range)ws.Columns[column];
                    Range? currentColumn = ws.Columns[column] as Range;

                    if (currentColumn is not null)
                    {
                        // Assume we want the whole column to change
                        currentColumn.Font.Name = headerFormat.Font.Name;
                        currentColumn.Font.Size = headerFormat.Font.Size;
                    }
                }
            }

            if (headerFormat == null)
            {
                //h/*eaderFormat = insertAt.;*/
            }

            //AddColumnHeaderToSheetX((Range)ws.Cells[row, column], columnWidth, headerTitle, headerFormat);

            Range? rng = ws.Cells[row, column] as Range;

            if (rng is not null)
            {
                AddColumnHeaderToSheetX(rng, columnWidth, headerTitle, headerFormat);
            }
        }

        #endregion

        #region AddCommentToCell

        public static void AddCommentToCell(
            Range rng,
            string text,
            CellFormatSpecification? cellFormatSpecification = null)
        {
            rng.AddComment(text);

            // TODO: Determine how to format the text differently.
            //With ws
            //    With .Cells(row, column)
            //        .Value = headerTitle
            //        .Font.Size = headerFontSize
            //        .Font.Bold = headerBold
            //        .Font.Underline = headerUnderline
            //        .WrapText = headerWrapText
            //        .HorizontalAlignment = headerHorizontalAlignment
            //    End With
            //End With
        }

        public static void AddCommentToCell(
            Worksheet ws, int column, int row,
            string text,
            CellFormatSpecification? cellFormatSpecification = null)
        {
            //Range rng = (Range)ws.Cells[row, column];
            Range? rng = ws.Cells[row, column] as Range;

            if (rng is not null)
            {
                AddCommentToCell(rng, text, cellFormatSpecification);
            }
        }

        #endregion

        #region AddContentToCell

        public static void AddContentToCell(
            Range rng,
            string text,
            CellFormatSpecification? cellFormat = null)
        {
            rng.Value = text;

            if (cellFormat != null)
            {
                UpdateFormatting(rng, cellFormat);
            }
        }

        public static void AddOffsetContentToCell(
            XlLocation insertAt,
            string text,
            CellFormatSpecification? cellFormat = null)
        {
            insertAt.OffsetRange.Value = text;

            //rng.Value = text;

            if (cellFormat != null)
            {
                UpdateFormatting(insertAt.OffsetRange, cellFormat);
            }
        }

        public static void AddContentToCell(
            Worksheet ws, int row, int column,
            string text,
            CellFormatSpecification? cellFormat = null)
        {
            //Range rng = (Range)ws.Cells[row, column];
            Range? rng = ws.Cells[row, column] as Range;

            AddContentToCell(rng, text, cellFormat);
        }

        #endregion

        #region AddTitledInfo

        [Obsolete("AddLabledInfoX has been replaced with AddLabledInfo which takes an XlLocation instead of a Range")]
        public static void AddLabeledInfoX(
            Range rng,
            string title,
            string info,
            int columnWidth = -1,
            LabelLocation labelLocation = LabelLocation.Left,
            XlOrientation orientation = XlOrientation.xlHorizontal,
            CellFormatSpecification titleFormat = null,
            CellFormatSpecification contentFormat = null)
        {
            AddContentToCell(rng.Offset[0, 0], title, titleFormat);

            if (labelLocation == LabelLocation.Top)
            {
                AddContentToCell(rng.Offset[1, 0], info, contentFormat);
            }
            else
            {
                if (orientation == XlOrientation.xlUpward)
                {
                    AddContentToCell(rng.Offset[-1, 0], info, contentFormat);
                }
                else
                {
                    AddContentToCell(rng.Offset[0, 1], info, contentFormat);
                }
            }

            if (columnWidth > 0)
            {
                //Range column = (Range)rng.Worksheet.Columns[rng.Column];
                //column.ColumnWidth = columnWidth;
                Range? rngColumn = rng.Worksheet.Columns[rng.Column] as Range;

                if (rngColumn != null)
                {
                    rngColumn.ColumnWidth = columnWidth;
                }
            }
        }

        [Obsolete("AddTitleInfoX has been replaced with AddTitledInfo which takes an XlLocation instead of a Range")]
        public static void AddLabeledInfoX(
            Worksheet ws, int row, int column,
            string title,
            string info,
            int columnWidth = -1,
            LabelLocation lableLocation = LabelLocation.Left,
            XlOrientation orientation = XlOrientation.xlHorizontal,
            CellFormatSpecification? titleFormat = null,
            CellFormatSpecification? contentFormat = null)
        {
            //Range rng = (Range)ws.Cells[row, column];
            Range? rng = ws.Cells[row, column] as Range;

            if (rng != null)
            {
                AddLabeledInfoX(rng, title, info, columnWidth, lableLocation, orientation, titleFormat, contentFormat);
            }
        }

        public static void AddLabeledInfo(
            XlLocation insertAt,
            string title,
            string info,
            int columnWidth = -1,
            LabelLocation labelLocation = LabelLocation.Left,
            XlOrientation orientation = XlOrientation.xlHorizontal,
            CellFormatSpecification? labelFormat = null,
            CellFormatSpecification? contentFormat = null)
        {
            Range rng = insertAt.CurrentRange;
            //Range rng = insertAt.GetCurrentRange();

            if (labelFormat == null)
            {
                labelFormat = insertAt.LabelRightFormat;
            }

            if (contentFormat == null)
            {
                contentFormat = insertAt.ContentLeftFormat;
            }

            AddContentToCell(rng.Offset[0, 0], title, labelFormat);


            if (labelLocation == LabelLocation.Top)
            {
                AddContentToCell(rng.Offset[1, 0], info, contentFormat);
            }
            else
            {
                if (orientation == XlOrientation.xlUpward)
                {
                    AddContentToCell(rng.Offset[-1, 0], info, contentFormat);
                }
                else
                {
                    AddContentToCell(rng.Offset[0, 1], info, contentFormat);
                }
            }

            if (columnWidth > 0)
            {
                //Range column = (Range)rng.Worksheet.Columns[rng.Column];
                //column.ColumnWidth = columnWidth;
                Range? rngColumn = rng.Worksheet.Columns[rng.Column] as Range;

                if (rngColumn != null)
                {
                    rngColumn.ColumnWidth = columnWidth;
                }
            }
        }

        public static void AddSectionInfo(
            XlLocation insertAt,
            string title,
            string info,
            int columnWidth = -1,
            SectionLocation sectionLocation = SectionLocation.Left,
            XlOrientation orientation = XlOrientation.xlHorizontal,
            CellFormatSpecification? sectionFormat = null,
            CellFormatSpecification? contentFormat = null)
        {
            Range rng = insertAt.CurrentRange;
            //Range rng = insertAt.GetCurrentRange();

            if (sectionFormat == null)
            {
                sectionFormat = insertAt.SectionLeftFormat;
            }

            if (contentFormat == null)
            {
                contentFormat = insertAt.ContentLeftFormat;
            }

            AddContentToCell(rng.Offset[0, 0], title, sectionFormat);


            if (sectionLocation == SectionLocation.Top)
            {
                AddContentToCell(rng.Offset[1, 0], info, contentFormat);
            }
            else
            {
                if (orientation == XlOrientation.xlUpward)
                {
                    AddContentToCell(rng.Offset[-1, 0], info, contentFormat);
                }
                else
                {
                    AddContentToCell(rng.Offset[0, 1], info, contentFormat);
                }
            }

            if (columnWidth > 0)
            {
                //Range column = (Range)rng.Worksheet.Columns[rng.Column];
                //column.ColumnWidth = columnWidth;
                Range? rngColumn = rng.Worksheet.Columns[rng.Column] as Range;

                if (rngColumn != null)
                {
                    rngColumn.ColumnWidth = columnWidth;
                }
            }
        }

        #endregion

        //Public Sub ApplicationInfo()
        //    Try
        //        Debug.Print("Application.CommonAppDataPath:" & Application.CommonAppDataPath.ToString)
        //        Debug.Print("Application.CommonAppDataRegistry:" & Application.CommonAppDataRegistry.ToString)
        //        Debug.Print("Application.CompanyName:" & Application.CompanyName.ToString)
        //        Debug.Print("Application.CurrentCulture:" & Application.CurrentCulture.ToString)
        //        Debug.Print("Application.CurrentInputLanguage:" & Application.CurrentInputLanguage.ToString)
        //        Debug.Print("Application.ExecutablePath:" & Application.ExecutablePath.ToString)
        //        Debug.Print("Application.LocalUserAppDataPath:" & Application.LocalUserAppDataPath.ToString)
        //        Debug.Print("Application.ProductName:" & Application.ProductName.ToString)
        //        Debug.Print("Application.ProductVersion:" & Application.ProductVersion.ToString)
        //        Debug.Print("Application.SafeTopLevelCaptionFormat:" & Application.SafeTopLevelCaptionFormat.ToString)
        //        Debug.Print("Application.StartupPath:" & Application.StartupPath.ToString)
        //        Debug.Print("Application.UserAppDataPath:" & Application.UserAppDataPath.ToString)
        //        Debug.Print("Application.UserAppDataRegistry:" & Application.UserAppDataRegistry.ToString)

        //        Debug.Print("ThisAddin.Application.StartupPath:" & Globals.ThisAddIn.Application.StartupPath.ToString)
        //        Debug.Print("ThisAddin.Application.ActiveWorkbook.Name:" & Globals.ThisAddIn.Application.ActiveWorkbook.Name.ToString)
        //        Debug.Print("ThisAddin.Application.ActiveWorkbook.Path:" & Globals.ThisAddIn.Application.ActiveWorkbook.Path.ToString)
        //        Debug.Print("ThisAddin.Application.ActiveWorkbook.FullName:" & Globals.ThisAddIn.Application.ActiveWorkbook.FullName.ToString)
        //        Debug.Print("ThisAddin.Application.ActiveWorkbook.FullNameURLEncoded:" & Globals.ThisAddIn.Application.ActiveWorkbook.FullNameURLEncoded.ToString)

        //        Debug.Print("ThisAddin.Application.DefaultFilePath:" & Globals.ThisAddIn.Application.DefaultFilePath.ToString)
        //        Debug.Print("ThisAddin.Application.Name:" & Globals.ThisAddIn.Application.Name.ToString)
        //        Debug.Print("ThisAddin.Application.NetworkTemplatesPath:" & Globals.ThisAddIn.Application.NetworkTemplatesPath.ToString)
        //        Debug.Print("ThisAddin.Application.Path:" & Globals.ThisAddIn.Application.Path.ToString)
        //    Catch ex As Exception
        //        MessageBox.Show("ApplicationInfo():" & ex.ToString)
        //    End Try
        //End Sub

        public static void CalculationsOff()
        {
            // Don't bother trying to save current if no open workbooks.

            if (ExcelApplication?.Workbooks.Count > 0)
            {
                PriorCalculationState = ExcelApplication.Calculation;
                ExcelApplication.Calculation = XlCalculation.xlCalculationManual;
            }
            else
            {
                // Assume the intent is to run with calculation and screen updates on.
                // Hopefully we never get called with no workbooks open.
                PriorCalculationState = XlCalculation.xlCalculationAutomatic;
            }
        }

        public static void CalculationsOn(bool force = false)
        {
            if (ExcelApplication is not null)
            {
                if (force)
                {
                    ExcelApplication.Calculation = XlCalculation.xlCalculationAutomatic;
                }
                else
                {
                    ExcelApplication.Calculation = PriorCalculationState;
                }
            }
        }

        public static void CustomFooterExists(bool hasCustomFooter)
        {
            try
            {
                //Office.DocumentProperties prps = (Office.DocumentProperties)ExcelApplication.ActiveWorkbook.CustomDocumentProperties;
                Office.DocumentProperties? prps = ExcelApplication?.ActiveWorkbook.CustomDocumentProperties as Office.DocumentProperties;

                if (prps != null)
                {
                    // Add a new property.
                    Office.DocumentProperty prp = prps.Add("HasCustomFooter", false,
                        Office.MsoDocProperties.msoPropertyTypeBoolean, hasCustomFooter);
                }
            }
            catch (Exception ex)
            {
                //PLLog.Error(ex, Globals.PROJECT_NAME)
                MessageBox.Show("CustomFooterExists() Unable to add HasCustomFooter property" + ex.Message);
            }
        }

        public static void CustomTableOfContentsExists(bool hasCustomTableOfContents)
        {
            Office.DocumentProperty? prp = default;
            Office.DocumentProperties? prps = default;

            try
            {
                //prps = (Office.DocumentProperties)ExcelApplication.ActiveWorkbook.CustomDocumentProperties;
                prps = ExcelApplication?.ActiveWorkbook.CustomDocumentProperties as Office.DocumentProperties;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            if (hasCustomTableOfContents)
            {
                try
                {
                    // Add a new property.
                    prp = prps?.Add("HasCustomTableOfContents", false,
                        Office.MsoDocProperties.msoPropertyTypeBoolean, hasCustomTableOfContents);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("CustomTableOfContentsExists() Unable to add HasCustomTableOfContents property" + ex.Message);
                }
            }
            else
            {
                prps = ExcelApplication?.ActiveWorkbook.CustomDocumentProperties as Office.DocumentProperties;
                prps?["HasCustomTableOfContents"]?.Delete();
            }
        }

        public static void DeleteSheet(Worksheet ws, bool prompt = false)
        {
            bool priorState = false;

            priorState = ExcelApplication?.DisplayAlerts ?? false;

            if (ExcelApplication is not null)
            {
                if (prompt)
                {
                    ExcelApplication.DisplayAlerts = true;
                    ws.Delete();
                }
                else
                {
                    ExcelApplication.DisplayAlerts = false;
                    ws.Delete();
                }

                ExcelApplication.DisplayAlerts = priorState;
            }
        }

        public void DisplayExcelRange(Range rng)
        {
            //Debug.Print(rng.Address);
        }

        //public static void DisplayInWatchWindow(string outputLine)
        //{
        //    AddinHelper.Common.WriteToWatchWindow(string.Format("{0}", outputLine));
        //}

        // REVIEW(crhodes)
        // Not sure these are needed.  Compare with what is in Common

        public static long DisplayInWatchWindow(string outputLine, [CallerMemberName] string callingMember = "")
        {
            Common.WriteToWatchWindow(string.Format("{0}: {1}", callingMember, outputLine));

            return Stopwatch.GetTimestamp();
        }

        public static long DisplayInWatchWindow(string outputLine, long startTicks, [CallerMemberName] string callingMember = "")
        {
            Common.WriteToWatchWindow(string.Format("{0}: {1} ({2:0.0000})",
                callingMember, outputLine,
                (Stopwatch.GetTimestamp() - startTicks) / (double)Stopwatch.Frequency));

            return Stopwatch.GetTimestamp();
        }

        public static long DisplayInWatchWindow(long startTicks, [CallerMemberName] string callingMember = "")
        {
            Common.WriteToWatchWindow(string.Format("{0}: ({1:0.0000})",
                callingMember,
                (Stopwatch.GetTimestamp() - startTicks) / (double)Stopwatch.Frequency));

            return Stopwatch.GetTimestamp();
        }

        public static long DisplayInWatchWindow(XlLocation insertAt, string suffix = "", [CallerMemberName] string callingMember = "")
        {
            if (Common.DisplayXlLocationUpdates)
            {
                // Get who called us

                StackFrame frame = new StackFrame(1);
                MethodBase method = frame.GetMethod();

                return DisplayInWatchWindow(string.Format("{0,50}-{21,3} {24,-30} sRC:({1,3}:{2,3}) cRC:({3,3}:{4,3}) oRC:({5,3}:{6,3}) moRC:({22,3}:{23,3}) aRC:({7,3}:{8,3}) msRC:({9,3}:{10,3}) meRC:({11,3}:{12,3}) gsRC:({13,3}:{14,3}) geRC:({15,3}:{16,3}) tsRC:({17,3}:{18,3}) teRC:({19,3}:{20,3})",
                                callingMember,                                      // 0
                                insertAt.RowStart, insertAt.ColumnStart,            // 1, 2
                                insertAt.RowCurrent, insertAt.ColumnCurrent,        // 3, 4
                                insertAt.RowOffset, insertAt.ColumnOffset,          // 5, 6
                                insertAt.RowsAdded, insertAt.ColumnsAdded,          // 7, 8
                                insertAt.MarkStartRow, insertAt.MarkStartColumn,    // 9, 10
                                insertAt.MarkEndRow, insertAt.MarkEndColumn,        // 11, 12
                                insertAt.GroupStartRow, insertAt.GroupStartColumn,  // 13, 14
                                insertAt.GroupEndRow, insertAt.GroupEndColumn,      // 15, 16
                                insertAt.TableStartRow, insertAt.TableStartColumn,  // 17, 18
                                insertAt.TableEndRow, insertAt.TableEndColumn,      // 19, 20
                                suffix,                                             // 21
                                insertAt.RowOffsetMax, insertAt.ColumnOffsetMax,    // 22, 23
                                method.Name));                                      // 24
            }
            else
            {
                return DisplayInWatchWindow(callingMember);
            }
        }

        public static long DisplayInWatchWindow(XlLocation insertAt, long startTicks, string suffix = "", [CallerMemberName] string callingMember = "")
        {
            if (Common.DisplayXlLocationUpdates)
            {
                // Get who called us

                StackFrame frame = new StackFrame(1);
                MethodBase method = frame.GetMethod();

                return DisplayInWatchWindow(string.Format("{0,50}-{21,3} {24,-30} sRC:({1,3}:{2,3}) cRC:({3,3}:{4,3}) oRC:({5,3}:{6,3}) moRC:({22,3}:{23,3}) aRC:({7,3}:{8,3}) msRC:({9,3}:{10,3}) meRC:({11,3}:{12,3}) gsRC:({13,3}:{14,3}) geRC:({15,3}:{16,3}) tsRC:({17,3}:{18,3}) teRC:({19,3}:{20,3})",
                                callingMember,                                      // 0
                                insertAt.RowStart, insertAt.ColumnStart,        // 1, 2
                                insertAt.RowCurrent, insertAt.ColumnCurrent,    // 3, 4
                                insertAt.RowOffset, insertAt.ColumnOffset,          // 5, 6
                                insertAt.RowsAdded, insertAt.ColumnsAdded,          // 7, 8
                                insertAt.MarkStartRow, insertAt.MarkStartColumn,    // 9, 10
                                insertAt.MarkEndRow, insertAt.MarkEndColumn,        // 11, 12
                                insertAt.GroupStartRow, insertAt.GroupStartColumn,  // 13, 14
                                insertAt.GroupEndRow, insertAt.GroupEndColumn,      // 15, 16
                                insertAt.TableStartRow, insertAt.TableStartColumn,  // 17, 18
                                insertAt.TableEndRow, insertAt.TableEndColumn,      // 19, 20
                                suffix,                                             // 21
                                insertAt.RowOffsetMax, insertAt.ColumnOffsetMax,    // 22, 23
                                method.Name),
                                startTicks, callingMember);                                      // 24
            }
            else
            {
                return DisplayInWatchWindow(startTicks, callingMember);
            }
        }

        public static bool DisplayScreenUpdates
        {
            get;
            set;
        }

        public static Worksheet DuplicateWorksheet(string sourceSheetName, string destinationSheetName,
            string beforeSheetName = "", string afterSheetName = "")
        {
            Workbook? wb = ExcelApplication?.ActiveWorkbook;

            foreach (Worksheet ws in wb.Worksheets)
            {
                if (ws.Name == destinationSheetName)
                {
                    // TODO: Sheet exists.  Ask user what to do.
                    MessageBox.Show(string.Format("Destination Sheet: >{0}< already exists.", sourceSheetName));
                    return ws;
                }
            }

            Worksheet? sourceSheet = wb.Sheets[sourceSheetName] as Worksheet;

            if (!string.IsNullOrEmpty(beforeSheetName))
            {
                //((Worksheet)wb.Sheets[sourceSheetName]).Copy(Before: wb.Sheets[beforeSheetName]);
                sourceSheet?.Copy(Before: (object)wb.Sheets[beforeSheetName]);
            }
            else if (!string.IsNullOrEmpty(afterSheetName))
            {
                //((Worksheet)wb.Sheets[sourceSheetName]).Copy(After: wb.Sheets[afterSheetName]);
                sourceSheet?.Copy(After: (object)wb.Sheets[afterSheetName]);
            }
            else
            {
                //((Worksheet)wb.Sheets[sourceSheetName]).Copy();
                sourceSheet?.Copy();
            }

            //((Worksheet)wb.ActiveSheet).Name = destinationSheetName;

            //return (Worksheet)wb.ActiveSheet;

            Worksheet? newSheet = wb.ActiveSheet as Worksheet;

            if (newSheet != null)
            {
                newSheet.Name = destinationSheetName;
            }

            return newSheet;
        }

        //----------------------------------------------------------------------
        //
        // EmptyWorkbook
        //
        // Returns the name of a workbook containing one sheet.
        // Creates new workbook if blnCreateNew
        // All existing sheets except for strWorksheet are removed.
        //
        // ToDo:
        //
        //----------------------------------------------------------------------

        public static string EmptyWorkbook(string strWorksheetName, bool blnCreateNew)
        {
            // REVIEW(crhodes)
            // I think once we figure out ExcelApplication we won't need this check.
            if (ExcelApplication is null)
            {
                throw new InvalidOperationException("ExcelApplication is not initialized.");
            }

            //Worksheet shtWS = default(Worksheet);
            string? strKeepName = null;

            if (true == blnCreateNew)
            {
                ExcelApplication.Workbooks.Add();
                // Keep the name so we don't have to worry about what name is given.
                Worksheet? activeSheet = ExcelApplication?.ActiveSheet as Worksheet;

                strKeepName = activeSheet?.Name;
            }
            else
            {
                strKeepName = strWorksheetName;
            }

            // Yes, delete the damn things!
            ExcelApplication.DisplayAlerts = false;

            // Remove all other worksheets

            foreach (Worksheet ws in ExcelApplication.ActiveWorkbook.Sheets)
            {
                if (strKeepName != ws.Name)
                {
                    ws.Delete();
                }
            }

            ExcelApplication.DisplayAlerts = true;

            Worksheet? activeSheetFinal = ExcelApplication?.ActiveSheet as Worksheet;

            if (activeSheetFinal != null)
            {
                activeSheetFinal.Name = strWorksheetName;
            }

            return strWorksheetName;
        }

        public static int FindFirst_PopulatedColumn_InRow(Range searchFromCell)
        {
            Worksheet searchWorksheet = searchFromCell.Worksheet;
            Range? searchRange = default;
            int columnNumber = searchFromCell.Column;
            int currentRow = searchFromCell.Row;
            int lastColSpecial = searchFromCell.SpecialCells(XlCellType.xlCellTypeLastCell).Column;

            try
            {
                searchRange = searchWorksheet.Range[searchWorksheet.Cells[currentRow, 1], searchWorksheet.Cells[currentRow, lastColSpecial]];

                columnNumber = searchRange.Find("*",
                After: searchWorksheet.Cells[currentRow, lastColSpecial],
                SearchOrder: XlSearchOrder.xlByRows,
                SearchDirection: XlSearchDirection.xlNext).Column;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            //try
            //{
            //    searchRange = searchWorksheet.get_Range((object)searchWorksheet.Cells[currentRow, 1], (object)searchWorksheet.Cells[currentRow, lastColSpecial]);

            //    // You got this error because the code was using a direct property access or method call on a COM object(like searchRange.Find(...))
            //    // and then immediately accessing .Column on the result, assuming it is a Range.
            //    // If the result is not a Range, or if the cast is not explicit, C# tries to use the dynamic runtime binder,
            //    // which is not available unless you reference Microsoft.CSharp and target .NET Framework.
            //    // This might be happening because your project targets.NET Core or.NET 5+ and does not reference Microsoft.CSharp,
            //    // or you are not using the dynamic keyword.The direct property access triggers the missing runtime binder error.
            //    // Here's how I fixed the code: I explicitly cast the result of searchRange.Find(...) to Range using as Range,
            //    // and then checked for null before accessing the .Column property.
            //    // This avoids the dynamic binder and works in all .NET versions, ensuring type safety and compatibility with your project configuration.

            //    Range? foundCell = searchRange.Find("*",
            //        After: (object)searchWorksheet.Cells[currentRow, lastColSpecial],
            //        SearchOrder: XlSearchOrder.xlByRows,
            //        SearchDirection: XlSearchDirection.xlNext) as Range;

            //    if (foundCell != null)
            //    {
            //        columnNumber = foundCell.Column;
            //    }
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.ToString());
            //}

            return columnNumber;
        }

        public static int FindFirst_PopulatedRow_InColumn(Range searchFromCell)
        {
            Worksheet searchWorksheet = searchFromCell.Worksheet;
            Range currentColumnRange = default(Range);
            int rowNumber = 0;
            int currentColumn = searchFromCell.Column;
            int lastRowSpecial = searchFromCell.SpecialCells(XlCellType.xlCellTypeLastCell).Row;

            try
            {
                currentColumnRange = (Range)searchWorksheet.Columns.Item[searchFromCell.Column];
                rowNumber = currentColumnRange.Find("*",
                    After: searchWorksheet.Cells[lastRowSpecial, currentColumn],
                    SearchOrder: XlSearchOrder.xlByColumns,
                    SearchDirection: XlSearchDirection.xlNext).Row;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            return rowNumber;
        }

        public static int FindLast_PopulatedColumn_InRow(Range searchFromCell)
        {
            Worksheet searchWorksheet = searchFromCell.Worksheet;
            Range searchRange = default(Range);
            int columnNumber = searchFromCell.Column;
            int currentRow = searchFromCell.Row;
            int lastColSpecial = searchFromCell.SpecialCells(XlCellType.xlCellTypeLastCell).Column;

            try
            {
                searchRange = searchWorksheet.Range[searchWorksheet.Cells[currentRow, 1], searchWorksheet.Cells[currentRow, lastColSpecial]];

                columnNumber = searchRange.Find("*",
                    After: searchWorksheet.Cells[currentRow, 1],
                    SearchOrder: XlSearchOrder.xlByRows,
                    SearchDirection: XlSearchDirection.xlPrevious).Column;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            return columnNumber;
        }

        public static int FindLast_PopulatedRow_InColumn(Range searchFromCell)
        {
            Worksheet searchWorksheet = searchFromCell.Worksheet;
            Range currentColumnRange = default(Range);
            int lastRow = 0;
            int currentColumn = searchFromCell.Column;

            try
            {
                currentColumnRange = (Range)searchWorksheet.Columns.Item[searchFromCell.Column];
                lastRow = currentColumnRange.Find("*",
                    After: searchWorksheet.Cells[1, currentColumn],
                    SearchOrder: XlSearchOrder.xlByColumns,
                    SearchDirection: XlSearchDirection.xlPrevious).Row;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            return lastRow;
        }


        public static int FindNext_PopulatedColumn_InRow(Range searchFromCell)
        {
            int columnNumber = 0;
            Range nextMatch;

            try
            {
                nextMatch = searchFromCell.Find("*", SearchOrder: XlSearchOrder.xlByRows, SearchDirection: XlSearchDirection.xlNext);

                if (nextMatch.Row != searchFromCell.Row)
                {
                    return searchFromCell.Column;
                }

                columnNumber = nextMatch.Column;

                if (columnNumber < searchFromCell.Column)
                {
                    columnNumber = searchFromCell.Column;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            return columnNumber;
        }

        public static int FindNext_PopulatedRow_InColumn(Range searchFromCell)
        {
            int rowNumber = 0;
            Range nextMatch;

            try
            {
                nextMatch = searchFromCell.Find("*", SearchOrder: XlSearchOrder.xlByColumns, SearchDirection: XlSearchDirection.xlNext);

                if (nextMatch.Column != searchFromCell.Column)
                {
                    return searchFromCell.Row;
                }

                rowNumber = nextMatch.Row;

                if (rowNumber < searchFromCell.Row)
                {
                    rowNumber = searchFromCell.Row;
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            return rowNumber;
        }

        public static int FindPrevious_PopulatedColumn_InRow(Range searchFromCell)
        {
            int columnNumber = 0;
            Range nextMatch;

            try
            {
                nextMatch = searchFromCell.Find("*", SearchOrder: XlSearchOrder.xlByRows, SearchDirection: XlSearchDirection.xlPrevious);

                if (nextMatch.Row != searchFromCell.Row)
                {
                    return searchFromCell.Column;
                }

                columnNumber = nextMatch.Column;

                if (columnNumber > searchFromCell.Column)
                {
                    columnNumber = searchFromCell.Column;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            return columnNumber;
        }

        public static int FindPrevious_PopulatedRow_InColumn(Range searchFromCell)
        {
            int rowNumber = 0;
            Range nextMatch;

            try
            {
                nextMatch = searchFromCell.Find("*", SearchOrder: XlSearchOrder.xlByColumns, SearchDirection: XlSearchDirection.xlPrevious);

                if (nextMatch.Column != searchFromCell.Column)
                {
                    return searchFromCell.Row;
                }

                rowNumber = nextMatch.Row;

                if (rowNumber > searchFromCell.Row)
                {
                    rowNumber = searchFromCell.Row;
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            return rowNumber;
        }

        public static string GetBuiltInPropertyValue(Workbook workBook, string propName)
        {
            string returnValue = "";

            //// This procedure returns the value of the built-in document
            //// property specified in the propName argument for the Office
            //// document object specified in the workBook argument.

            //DocumentProperty prpDocProp = default(DocumentProperty);
            //object varValue = null;

            //const long ERR_BADPROPERTY = 5;
            //const long ERR_BADDOCOBJ = 438;
            //const long ERR_BADCONTEXT = -2147467259;

            Office.DocumentProperties? documentProperties = workBook.BuiltinDocumentProperties as Office.DocumentProperties;

            if (documentProperties != null)
            {
                try
                {
                    Office.DocumentProperty property = documentProperties[propName];
                    returnValue = property.Value.ToString();
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    //Common.WriteToDebugWindow(string.Format("GetBuiltInPropertyValue({0}) COMException ErrorCode:({1})", propName, ex.ErrorCode));
                    returnValue = "Property Not Available Yet";
                }
                catch (NullReferenceException)
                {
                    //Common.WriteToDebugWindow(string.Format("GetBuiltInPropertyValue({0}) NullReferenceException", propName));
                    returnValue = "Property Not Set";
                }
                catch (Exception)
                {
                    //Common.WriteToDebugWindow(string.Format("GetBuiltInPropertyValue({0}) Exception >{1})<", propName, ex));
                }
            }

            return returnValue;
        }

        public static string GetFile(string initialFolder = "", string dialogTitle = "Open", string fileFilter = "All Files (*.*)|*.*")
        {
            // TODO(crhodes)
            // Implement

            //using (OpenFileDialog ofd = new OpenFileDialog())
            //{
            //    DialogResult result = default;
            //    ofd.Multiselect = false;
            //    ofd.InitialDirectory = initialFolder;
            //    ofd.Title = dialogTitle;
            //    ofd.Filter = fileFilter;
            //    result = ofd.ShowDialog();
            //    //Debug.WriteLine(ofd.FileName);
            //    return ofd.FileName;
            //}
            return "";
        }

        public static string GetOpenFileName(string initialFolder = "", string dialogTitle = "Open", string fileFilter = "All Files (*.*)|*.*")
        {
            // TODO(crhodes)
            // Implement

            //using (OpenFileDialog ofd = new OpenFileDialog())
            //{
            //    DialogResult result = default;
            //    ofd.Multiselect = false;
            //    ofd.InitialDirectory = initialFolder;
            //    ofd.Title = dialogTitle;
            //    ofd.Filter = fileFilter;
            //    result = ofd.ShowDialog();
            //    //Debug.WriteLine(ofd.FileName);
            //    return ofd.FileName;
            //}
            return "";
        }

        public static string GetSafeFileName(string fileName)
        {
            fileName = fileName.Replace("/", "-");
            fileName = fileName.Replace(@"\", "-");
            fileName = fileName.Replace("[", "");
            fileName = fileName.Replace("]", "");
            fileName = fileName.Replace(" ", "");
            fileName = fileName.Replace(":", "-");

            return fileName;
        }

        public static string GetSaveFileName(string initialFolder = "", string proposedFileName = "", string dialogTitle = "Open", string fileFilter = "All Files (*.*)|*.*")
        {
            // TODO(crhodes)
            // Implement

            //using (SaveFileDialog sfd = new SaveFileDialog())
            //{
            //    DialogResult result = default;
            //    sfd.InitialDirectory = initialFolder;
            //    sfd.Title = dialogTitle;
            //    sfd.Filter = fileFilter;
            //    sfd.FileName = proposedFileName;
            //    result = sfd.ShowDialog();
            //    //Debug.WriteLine(sfd.FileName);
            //    return sfd.FileName;
            //}
            return "";
        }

        private static string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = string.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (dividend - modulo) / 26;
            }

            return columnName;
        }

        public static int GetEndOfSectionDown(int startRow, int startCol, int lastPopulatedRow, int initialColumn)
        {
            Worksheet activeSheet = (Worksheet)Common.ExcelApplication.ActiveSheet;

            int functionReturnValue = 0;
            int matchingRow = 0;

            // Search down for a matching cell
            matchingRow = ((Range)activeSheet.Cells[startRow, startCol]).End[XlDirection.xlDown].Row;

            if (startCol == initialColumn)
            //if (startCol == _INITIAL_COL)
            {
                // We have back'd all the way back to the first column.
                // Return either the next matching cell down or the last populated row on the sheet.

                if (matchingRow < lastPopulatedRow)
                {
                    // Section ends on the row prior to the match.
                    functionReturnValue = matchingRow - 1;
                }
                else
                {
                    // Return end of populated section
                    functionReturnValue = lastPopulatedRow;
                }
            }
            else
            {
                if (matchingRow <= lastPopulatedRow)
                {
                    // Back up one column and search down for a populated cell.
                    // Treat row prior to matching row as new end.
                    functionReturnValue = GetEndOfSectionDown(startRow, startCol - 1, matchingRow - 1, initialColumn);
                }
                else
                {
                    // Back up one column and search down for a populated cell.
                    // Treat end of worksheet as end.
                    functionReturnValue = GetEndOfSectionDown(startRow, startCol - 1, lastPopulatedRow, initialColumn);
                }
            }

            return functionReturnValue;
        }

        public static void GroupAndHideColumns(Worksheet ws, int startingColumn, int endingColumn, bool hide)
        {
            string startLetter = GetExcelColumnName(startingColumn);
            string endLetter = GetExcelColumnName(endingColumn);

            string groupRange = string.Format("{0}:{1}", startLetter, endLetter);

            //((Range)ws.Columns[groupRange]).Group();
            //((Range)ws.Columns[groupRange]).Hidden = hide;

            Range? rng = ws.Columns[groupRange] as Range;

            if (rng != null)
            {
                rng.Group();
                rng.Hidden = hide;
            }
        }

        public static void GroupAndHideRows(Worksheet ws, int startingRow, int endingRow, bool hide)
        {
            string groupRange = string.Format("{0}:{1}", startingRow, endingRow);

            //((Range)ws.Rows[groupRange]).Group();
            //((Range)ws.Rows[groupRange]).Hidden = hide;

            Range? rng = ws.Rows[groupRange] as Range;

            if (rng != null)
            {
                rng.Group();
                rng.Hidden = hide;
            }
        }

        public static bool HasCustomFooter()
        {
            Office.DocumentProperty? prp = default;
            Office.DocumentProperties? prps = default;

            try
            {
                prps = ExcelApplication.ActiveWorkbook.CustomDocumentProperties as Office.DocumentProperties;
                prp = prps?["HasCustomFooter"];

                // If the property exists we don't really care about the value
                return true;
            }
            catch (Exception)
            {
                // Exception is thrown if property does not exist
                return false;
            }
        }

        public static bool HasCustomTableOfContents()
        {
            Office.DocumentProperty? prp = default;
            Office.DocumentProperties? prps = default;

            try
            {
                prps = ExcelApplication?.ActiveWorkbook.CustomDocumentProperties as Office.DocumentProperties;
                prp = prps?["HasCustomTableOfContents"];

                // If the property exists we don't really care about the value
                return true;
            }
            catch (Exception)
            {
                // Exception is thrown if property does not exist
                return false;
            }
        }

        /// <summary>
        /// Add a new worksheet.
        /// </summary>
        /// <param name="sheetName">
        /// Name of Worksheet. beforeSheetName can use FIRST, afterSheetName can use LAST
        /// </param>
        /// <param name="beforeSheetName">Sheet name or FIRST</param>
        /// <param name="afterSheetName">Sheet name or LAST</param>
        /// <remarks>beforeSheetName can use FIRST, afterSheetName can use LAST</remarks>
        /// <returns></returns>
        public static Worksheet NewWorksheet(String sheetName, String beforeSheetName = "", String afterSheetName = "")
        {
            Workbook? wb = ExcelApplication?.ActiveWorkbook;
            Worksheet? newWs;

            if (wb == null)
            {
                throw new InvalidOperationException("No active workbook found.");
            }

            foreach (Worksheet ws in wb.Worksheets)
            {
                if (ws.Name == sheetName)
                {
                    // Sheet exists.  Ask user what to do.
                    MessageBoxResult answer = MessageBox.Show("Preserve Existing WorkSheet: >" + sheetName + "< ?",
                        "Requested WorkSheet Already Exists",
                        MessageBoxButton.YesNo);

                    if (MessageBoxResult.Yes == answer)
                    {
                        return ws;
                    }
                    else
                    {
                        ws.Delete();
                    }
                }
            }

            //            You got this error because the code is using named/ optional parameters with COM objects(like wb.Sheets.Add(Before: wb.Sheets[1])) in a way that requires the C# dynamic runtime binder. This binder is not available unless you reference Microsoft.CSharp and target .NET Framework, not .NET Core or .NET 5+.
            //This might be happening because your project targets .NET Core or.NET 5 + and does not reference Microsoft.CSharp, or you are not using the dynamic keyword.The use of named/ optional parameters with COM objects triggers the missing runtime binder error.
            //Here's how I fixed the code: I explicitly cast the COM object arguments to object when calling methods like Sheets.Add. This avoids the dynamic binder and works in all .NET versions, ensuring type safety and compatibility with your project configuration.

            if (!string.IsNullOrEmpty(beforeSheetName))
            {
                if ("FIRST" == beforeSheetName)
                {
                    //newWs = (Worksheet)wb.Sheets.Add(Before: wb.Sheets[1]);
                    newWs = wb.Sheets.Add(Before: (object)wb.Sheets[1]) as Worksheet;
                }
                else
                {
                    //newWs = (Worksheet)wb.Sheets.Add(Before: wb.Sheets[beforeSheetName]);
                    newWs = wb.Sheets.Add(Before: (object)wb.Sheets[beforeSheetName]) as Worksheet;
                }
            }
            else if (!string.IsNullOrEmpty(afterSheetName))
            {
                if ("LAST" == afterSheetName)
                {
                    newWs = wb.Sheets.Add(After: (object)wb.Sheets[wb.Sheets.Count]) as Worksheet;
                }
                else
                {
                    newWs = wb.Sheets.Add(After: (object)wb.Sheets[afterSheetName]) as Worksheet;
                }
            }
            else
            {
                newWs = wb.Sheets.Add() as Worksheet;
            }

            if (newWs != null)
            {
                newWs.Name = sheetName;
                // Make the outline things show up the way CHR likes :)
                newWs.Outline.SummaryColumn = XlSummaryColumn.xlSummaryOnLeft;
                newWs.Outline.SummaryRow = XlSummaryRow.xlSummaryAbove;
            }

            return newWs;
        }

        //----------------------------------------------------------------------
        //
        // CreateDummyData
        //
        // Populate the range with some sample data. Used for testing and debugging.
        //
        //----------------------------------------------------------------------
        public static void CreateDummyData(Range dataRange)
        {
            //dataRange.Value2 = "ABC";

            //dataRange.Cells[1, 1].Value = "ABC";
            //dataRange.Cells[1, 2].Value = "DEF";
            //dataRange.Cells[1, 3].Value = "GHI";

            //for (int row = 0; row <= 4; row++)
            //{
            //    for (int col = 0; col <= 4; col++)
            //    {
            //        dataRange.Cells[row + 1, col + 1].Value = (char)('A' + col) + (row + 1).ToString();
            //    }
            //}

            //object[,] data = new object[5, 3];
            //for (int row = 0; row <= 4; row++)
            //{
            //    for (int col = 0; col <= 2; col++)
            //    {
            //        data[row, col] = (char)('A' + col) + (row + 1).ToString();
            //    }
            //}

            //dataRange.Value2 = data;

            //return;

            //try
            //{
            //    dataRange.Cells.ClearContents();
            //}
            //catch (Exception)
            //{
            //}
        }

        //
        // Protect or unprotect the sheet.  Return the current setting before
        // and changes made.
        //
        public static bool ProtectSheet(Worksheet workSheet, bool protectMode)
        {
            bool currentMode = workSheet.ProtectContents;

            if (protectMode == true)
            {
                workSheet.Protect();
            }
            else
            {
                workSheet.Unprotect();
            }

            return currentMode;
        }

        // Do this with regular expressions.

        public static string SafeSheetName(string workSheetName)
        {
            workSheetName = workSheetName.Replace("/", "-");
            workSheetName = workSheetName.Replace(@"\", "-");
            workSheetName = workSheetName.Replace("[", "");
            workSheetName = workSheetName.Replace("]", "");
            workSheetName = workSheetName.Replace(" ", "");
            workSheetName = workSheetName.Replace(":", "-");

            return workSheetName.Substring(0, workSheetName.Length > cMaxSheetNameLen ? cMaxSheetNameLen : workSheetName.Length);
        }

        public static void ScreenUpdatesOff()
        {
            if (false == DisplayScreenUpdates)
            {
                if (ExcelApplication.Workbooks.Count > 0)
                {
                    PriorScreenUpdatingState = ExcelApplication.ScreenUpdating;
                    ExcelApplication.ScreenUpdating = false;
                }
                else
                {
                    // Assume the intent is to run with screen updates on.
                    PriorScreenUpdatingState = true;
                    ExcelApplication.ScreenUpdating = false;
                }
            }
        }

        public static void ScreenUpdatesOn(bool force = false)
        {
            if (force)
            {
                ExcelApplication.ScreenUpdating = true;
            }
            else
            {
                ExcelApplication.ScreenUpdating = PriorScreenUpdatingState;
            }
        }

        public static void SetCellValue(Range rngR, object vntValue, XlHAlign horizontalAlignment = XlHAlign.xlHAlignLeft, string strComment = "")
        {
            rngR.Value = vntValue;
            rngR.HorizontalAlignment = horizontalAlignment;
            if (!string.IsNullOrEmpty(strComment))
            {
                rngR.AddComment();
                rngR.Comment.Visible = false;
                rngR.Comment.Text(Text: strComment);
            }
        }


        //Public Sub TestScreenOff()
        //    Application.ScreenUpdating = False

        //    Application.Workbooks.Add()

        //    Application.ScreenUpdating = True
        //End Sub


        //Private Sub DumpPropertyCollection( _
        // ByVal prps As Office.DocumentProperties, _
        // ByVal rng As Excel.Range, ByRef i As Integer)
        //    Dim prp As Office.DocumentProperty

        //    For Each prp In prps
        //        rng.Offset(i, 0).Value = prp.Name
        //        Try
        //            If Not prp.Value Is Nothing Then
        //                rng.Offset(i, 1).Value = _
        //                 prp.Value.ToString
        //            End If
        //        Catch
        //            ' Do nothing at all.
        //        End Try
        //        i += 1
        //    Next
        //End Sub

        private static void UpdateFormatting(Range targetRange, CellFormatSpecification formatSpec)
        {
            try
            {
                if (formatSpec.Font != null)
                {
                    targetRange.Font.Size = formatSpec.Font.Size;
                    targetRange.Font.Bold = formatSpec.Font.Bold;
                    targetRange.Font.Underline = formatSpec.Font.Underline;
                    targetRange.Font.Color = formatSpec.Font.Color;
                    targetRange.Font.Italic = formatSpec.Font.Italic;
                }
                else
                {
                    // TODO(crhodes)
                    // Why is this null?              
                }

                targetRange.Orientation = formatSpec.Orientation;
                targetRange.HorizontalAlignment = formatSpec.HorizontalAlignment;
                targetRange.VerticalAlignment = formatSpec.VerticalAlignment;
                targetRange.WrapText = formatSpec.WrapText;
            }
            catch (Exception)
            {
                MessageBox.Show($"FormatSpec: {formatSpec} error");
            }
        }

        public static void ZapPageBreaks()
        {
            foreach (Worksheet ws in ExcelApplication.ActiveWorkbook.Sheets)
            {
                ws.PageSetup.PrintArea = "";

                //Debug.Print(ws.Name);
                //        Debug.Print sht.HPageBreaks.Count
                //        Debug.Print sht.VPageBreaks.Count
                // For some reason the page break handling is not clean.
                // There are different types of page breaks, that is clear.
                // Unfortunately the For Each hPB errors out if only Automatic
                // Page breaks.  Wrap in try catch for AddIn
                // ERROR: Not supported in C#: OnErrorStatement

                if (ws.VPageBreaks.Count > 0)
                {
                    foreach (VPageBreak vPB in ws.VPageBreaks)
                    {
                        if (vPB.Type == XlPageBreak.xlPageBreakManual)
                        {
                            vPB.Delete();
                        }
                    }
                }

                if (ws.HPageBreaks.Count > 0)
                {
                    foreach (HPageBreak hPB in ws.HPageBreaks)
                    {
                        if (hPB.Type == XlPageBreak.xlPageBreakManual)
                        {
                            hPB.Delete();
                        }
                    }
                }
            }
        }

        #endregion
    }
}
