using System;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows;

using VNC.VSTOAddIn.Excel.Domain;

using MSExcel = Microsoft.Office.Interop.Excel;

namespace VNC.VSTOAddIn.Excel
{
    public class Common : VNC.VSTOAddIn.Common
    {
        new public const string LOG_CATEGORY = "VNCVSTOAddInExcel";

        // NOTE(crhodes)
        // This is set in VNCVSTOAddInExcel.ThisAddIn_Startup. 
        // It allows access to the Excel Application object without having to pass it around as a parameter.
        // It is only available from Globals in the top level VSTO AddIn, e.g. VNCExcelTools.

        public static MSExcel.Application? ExcelApplication { get; set; }

        // TODO(crhodes)
        // All of this came from VNC.ExcelHelper.
        // We need to decide what we want to keep and what we want to remove.

        //public const string TAG_PREFIX = "VNC";

        //public static Boolean HasAppEvents = true;  // Custom Header and Footer need this enabled.
        //public static Boolean DisplayEvents = false;

        //public static Boolean DebugSQL
        //{
        //    get;
        //    set;
        //}
        //public static Boolean DebugLevel1
        //{
        //    get;
        //    set;
        //}
        //public static Boolean DebugLevel2
        //{
        //    get;
        //    set;
        //}
        //public static Boolean DebugMode
        //{
        //    get;
        //    set;
        //}
        //public static Boolean DeveloperMode
        //{
        //    get;
        //    set;
        //}

        public static Boolean DisplayXlLocationUpdates
        {
            get;
            set;
        }

        //private static void DisplayDataSet(DataSet dataSet)
        //{
        //    DisplayTables(dataSet.Tables);
        //}

        //private static void DisplayTables(DataTableCollection tables)
        //{
        //    foreach (DataTable table in tables)
        //    {
        //        WriteToDebugWindow(string.Format("Table:   >{0}<", table.TableName));
        //        WriteToDebugWindow("Columns:");

        //        foreach (DataColumn column in table.Columns)
        //        {
        //            WriteToDebugWindow(string.Format(" >{0}<", column.ColumnName));
        //        }

        //        WriteToDebugWindow("");
        //        WriteToDebugWindow(string.Format("Rows:{0}", Environment.NewLine));

        //        foreach (DataRow row in table.Rows)
        //        {
        //            foreach (DataColumn column in table.Columns)
        //            {
        //                WriteToDebugWindow(string.Format(" >{0}<", row[column.ColumnName]));
        //            }
        //            WriteToDebugWindow("");
        //        }
        //    }
        //}

        public static long WriteToWatchWindow(XlLocation insertAt, string suffix = "", [CallerMemberName] string callingMember = "")
        {
            if (Common.DisplayXlLocationUpdates)
            {
                // Get who called us

                StackFrame frame = new StackFrame(1);
                MethodBase method = frame.GetMethod();

                return WriteToWatchWindow(string.Format("{0,50}-{21,3} {24,-30} sRC:({1,3}:{2,3}) cRC:({3,3}:{4,3}) oRC:({5,3}:{6,3}) moRC:({22,3}:{23,3}) aRC:({7,3}:{8,3}) msRC:({9,3}:{10,3}) meRC:({11,3}:{12,3}) gsRC:({13,3}:{14,3}) geRC:({15,3}:{16,3}) tsRC:({17,3}:{18,3}) teRC:({19,3}:{20,3})",
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
                                method.Name));                                      // 24
            }
            else
            {
                return WriteToWatchWindow("");
            }
        }

        public static long WriteToWatchWindow(XlLocation insertAt, long startTicks, string suffix = "", [CallerMemberName] string callingMember = "")
        {
            if (Common.DisplayXlLocationUpdates)
            {
                // Get who called us

                StackFrame frame = new StackFrame(1);
                MethodBase method = frame.GetMethod();

                return WriteToWatchWindow(string.Format("{0,50}-{21,3} {24,-30} sRC:({1,3}:{2,3}) cRC:({3,3}:{4,3}) oRC:({5,3}:{6,3}) moRC:({22,3}:{23,3}) aRC:({7,3}:{8,3}) msRC:({9,3}:{10,3}) meRC:({11,3}:{12,3}) gsRC:({13,3}:{14,3}) geRC:({15,3}:{16,3}) tsRC:({17,3}:{18,3}) teRC:({19,3}:{20,3})",
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
                                method.Name),                                       // 24
                                startTicks, callingMember);
            }
            else
            {
                return WriteToWatchWindow("", startTicks);
            }
        }

    }
}
