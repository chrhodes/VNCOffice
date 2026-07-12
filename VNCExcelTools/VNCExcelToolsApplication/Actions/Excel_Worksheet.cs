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
    public class Excel_Worksheet
    {
        public static IEventAggregator eventAggregator;
        public static DeveloperModeEvent developerModeEvent;
        public static StatusMessageEvent statusMessageEvent;

        public static WindowHost _aboutHost = null;

        public static void LockWorksheet()
        {
            Common.WriteToDebugWindow("");

            Worksheet worksheet = Common.ExcelApplication.ActiveSheet as Worksheet;
            worksheet.Protect(DrawingObjects: false, AllowFormattingCells: true);
        }

        public static void UnlockWorksheet()
        {
            Common.WriteToDebugWindow("");

            Worksheet worksheet = Common.ExcelApplication.ActiveSheet as Worksheet;
            worksheet.Unprotect();
        }

    }
}
