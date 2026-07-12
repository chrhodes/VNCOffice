using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;

using Microsoft.Office.Interop.Excel;

using Prism.Events;

using VNC;
using VNC.Core.Events;
using VNC.Core.Presentation;
using VNC.WPF.Presentation.Views;

using VNCExcelToolsApplication.Presentation.Views;

using MSExcel = Microsoft.Office.Interop.Excel;
using VNCExcelAddIn = VNC.VSTOAddIn.Excel;

namespace VNCExcelToolsApplication.Actions
{
    public class Excel_Workbook
    {
        public static IEventAggregator eventAggregator;
        public static DeveloperModeEvent developerModeEvent;
        public static StatusMessageEvent statusMessageEvent;

        public static WindowHost _aboutHost = null;

        public static void AddTableOfContents()
        {
            Worksheet tableOfContents;
            int row = 3;        // Starting Row
            int column = 1;     // Starting Column
            Boolean currentSheetProtectionMode;
            Workbook activeWorkbook = Common.ExcelApplication.ActiveWorkbook;

            if (VNCExcelAddIn.Helper.HasCustomTableOfContents())
            {
                // Remove the flag until we rebuild the tableOfContents so the NewWorkSheet event handler
                // does not trigger a recursive loop.
                VNCExcelAddIn.Helper.CustomTableOfContentsExists(false);
            }

            try
            {
                // Use the existing sheet if one exists
                tableOfContents = (Worksheet)activeWorkbook.Sheets["Table of Contents"];
            }
            catch (Exception)
            {
                // else create it
                tableOfContents = VNCExcelAddIn.Helper.NewWorksheet("Table of Contents", beforeSheetName: "FIRST");
                //tableOfContents = activeWorkbook.Sheets.Add();
                //tableOfContents.Name = "Table of Contents";
            }

            ((Range)tableOfContents.Columns["A:A"]).ClearContents();

            foreach (Worksheet sheet in activeWorkbook.Sheets)
            {
                switch (sheet.Name)
                {
                    case "Table of Contents":
                        break;

                    default:
                        // Unprotect the sheet before adding the hyperlink
                        currentSheetProtectionMode = VNCExcelAddIn.Helper.ProtectSheet(sheet, false);
                        ((Range)sheet.Cells[1, 1]).Value = "Table of Contents";
                        sheet.Hyperlinks.Add(Anchor: sheet.Cells[1, 1], Address: "", SubAddress: "'" + tableOfContents.Name + "'!A1", TextToDisplay: tableOfContents.Name);
                        // Then restore the setting
                        VNCExcelAddIn.Helper.ProtectSheet(sheet, currentSheetProtectionMode);

                        // Now update the Table of Contents Sheet
                        ((Range)tableOfContents.Cells[row, column]).Value = sheet.Name;
                        tableOfContents.Hyperlinks.Add(Anchor: tableOfContents.Cells[row, column], Address: "", SubAddress: "'" + sheet.Name + "'!A1", TextToDisplay: sheet.Name);

                        row += 1;
                        break;
                }
            }
            ((Range)tableOfContents.Columns["A:A"]).EntireColumn.AutoFit();

            if (!VNCExcelAddIn.Helper.HasCustomTableOfContents())
            {
                VNCExcelAddIn.Helper.CustomTableOfContentsExists(true);
            }

        }

        public static void LockAllWorksheets()
        {
            Common.WriteToDebugWindow("");

            foreach (Worksheet ws in Common.ExcelApplication.ActiveWorkbook.Sheets)
            {
                ws.Protect(DrawingObjects: false, AllowFormattingCells: true);
            }
        }

        public static void UnlockAllWorksheets()
        {
            foreach (Worksheet ws in Common.ExcelApplication.ActiveWorkbook.Sheets)
            {
                ws.Unprotect();
            }
        }
    }
}
