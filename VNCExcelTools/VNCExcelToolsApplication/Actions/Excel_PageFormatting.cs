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
    public class Excel_PageFormatting
    {
        public static IEventAggregator eventAggregator;
        public static DeveloperModeEvent developerModeEvent;
        public static StatusMessageEvent statusMessageEvent;

        public static WindowHost _aboutHost = null;

        public static void AddHeaderAllPages()
        {
            Common.WriteToDebugWindow("");

            Common.ExcelApplication.PrintCommunication = false;

            Workbook workBook;
            //Worksheet workSheet;
            StringBuilder sb = new StringBuilder();

            try
            {
                workBook = Common.ExcelApplication.ActiveWorkbook;
                VNCExcelAddIn.Helper.CalculationsOff();

                foreach (Worksheet workSheet in workBook.Sheets)
                {
                    sb.Length = 0;

                    sb.AppendFormat("&A");
                    //sb.Append("&5&Z&F");    // Five point font, path, and filename
                    //sb.AppendFormat("\nCreated: {0} By: {1}",
                    //    VNCExcelAddIn.GetBuiltInPropertyValue(workBook, "Creation Date"),
                    //    VNCExcelAddIn.GetBuiltInPropertyValue(workBook, "Author"));
                    //sb.AppendFormat("\nLast Saved: {0} By: {1}",
                    //    VNCExcelAddIn.GetBuiltInPropertyValue(workBook, "Last Save Time"),
                    //    VNCExcelAddIn.GetBuiltInPropertyValue(workBook, "Last author"));
                    //sb.AppendFormat("\nLast Printed: {0}",
                    //    VNCExcelAddIn.GetBuiltInPropertyValue(workBook, "Last Print Date"));

                    //if (Common.DebugMode)
                    //{
                    //    Common.WriteToDebugWindow("LeftFooter:>" + sb.ToString() + "<");                    	;
                    //}

                    //workSheet.PageSetup.LeftHeader = @"&10 Left &A";
                    workSheet.PageSetup.LeftHeader = "";

                    workSheet.PageSetup.CenterHeader = workSheet.Name;

                    sb.Length = 0;

                    //sb.Append("&5&P - &N"); // Five point font, page of pages
                    //sb.Append("\nTitle :");
                    //sb.Append(VNCExcelAddIn.GetBuiltInPropertyValue(workBook, "Title"));
                    //sb.Append("\nSubject: ");
                    //sb.Append(VNCExcelAddIn.GetBuiltInPropertyValue(workBook, "Subject"));

                    //if (Common.DebugMode)
                    //{
                    //    Common.WriteToDebugWindow("RightFooter:>" + sb.ToString() + "<");                    	;
                    //}

                    workSheet.PageSetup.RightHeader = "";
                    //workSheet.PageSetup.RightHeader = @"&5 Right";

                    // Indicate we have added a custom footer.
                    // This gets checked in the BeforeClose event.

                    //if ( ! VNCExcelAddIn.HasCustomFooter())
                    //{
                    //    VNCExcelAddIn.CustomFooterExists(true);
                    //}
                }

                VNCExcelAddIn.Helper.CalculationsOn();
            }
            catch (Exception ex)
            {
                MessageBox.Show("AddHeader:" + ex.ToString());
            }

            Common.ExcelApplication.PrintCommunication = true;
        }

        public static void AddFooterAllPages()
        {
            Common.WriteToDebugWindow("");

            Common.ExcelApplication.PrintCommunication = false;

            Workbook workBook;
            //Worksheet workSheet;
            StringBuilder sb = new StringBuilder();

            try
            {
                workBook = Common.ExcelApplication.ActiveWorkbook;
                VNCExcelAddIn.Helper.CalculationsOff();

                foreach (Worksheet workSheet in workBook.Sheets)
                {
                    sb.Length = 0;

                    sb.Append("&5&Z&F");    // Five point font, path, and filename
                    sb.AppendFormat("\nCreated: {0} By: {1}",
                        VNCExcelAddIn.Helper.GetBuiltInPropertyValue(workBook, "Creation Date"),
                        VNCExcelAddIn.Helper.GetBuiltInPropertyValue(workBook, "Author"));
                    sb.AppendFormat("\nLast Saved: {0} By: {1}",
                        VNCExcelAddIn.Helper.GetBuiltInPropertyValue(workBook, "Last Save Time"),
                        VNCExcelAddIn.Helper.GetBuiltInPropertyValue(workBook, "Last author"));
                    sb.AppendFormat("\nLast Printed: {0}",
                        VNCExcelAddIn.Helper.GetBuiltInPropertyValue(workBook, "Last Print Date"));

                    //if (Common.DebugMode)
                    //{
                        Common.WriteToDebugWindow("LeftFooter:>" + sb.ToString() + "<"); ;
                    //}

                    workSheet.PageSetup.LeftFooter = sb.ToString();

                    workSheet.PageSetup.CenterFooter = "";

                    sb.Length = 0;

                    sb.Append("&5&P - &N"); // Five point font, page of pages
                    sb.Append("\nTitle :");
                    sb.Append(VNCExcelAddIn.Helper.GetBuiltInPropertyValue(workBook, "Title"));
                    sb.Append("\nSubject: ");
                    sb.Append(VNCExcelAddIn.Helper.GetBuiltInPropertyValue(workBook, "Subject"));

                    //if (Common.DebugMode)
                    //{
                        Common.WriteToDebugWindow("RightFooter:>" + sb.ToString() + "<"); ;
                    //}

                    workSheet.PageSetup.RightFooter = sb.ToString();

                    // Indicate we have added a custom footer.
                    // This gets checked in the BeforeClose event.

                    if (!VNCExcelAddIn.Helper.HasCustomFooter())
                    {
                        VNCExcelAddIn.Helper.CustomFooterExists(true);
                    }
                }

                VNCExcelAddIn.Helper.CalculationsOn();
            }
            catch (Exception ex)
            {
                MessageBox.Show("AddFooter:" + ex.ToString());
            }

            Common.ExcelApplication.PrintCommunication = true;
        }

        public static void FitColumns()
        {
            Common.WriteToDebugWindow("");

            Common.ExcelApplication.PrintCommunication = false;

            foreach (Worksheet sheet in Common.ExcelApplication.ActiveWorkbook.Sheets)
            {
                sheet.PageSetup.FitToPagesWide = 1;
                sheet.PageSetup.FitToPagesTall = 0;
            }

            Common.ExcelApplication.PrintCommunication = true;
        }

        public static void FitRows()
        {
            Common.WriteToDebugWindow("");

            Common.ExcelApplication.PrintCommunication = false;

            foreach (Worksheet sheet in Common.ExcelApplication.ActiveWorkbook.Sheets)
            {
                sheet.PageSetup.FitToPagesWide = 0;
                sheet.PageSetup.FitToPagesTall = 1;
            }

            Common.ExcelApplication.PrintCommunication = true;
        }

        public static void FitToPages(string tag)
        {
            Common.WriteToDebugWindow("");

            string[] flags = tag.Split(',');

            Common.ExcelApplication.PrintCommunication = false;

            foreach (Worksheet sheet in Common.ExcelApplication.ActiveWorkbook.Sheets)
            {
                sheet.PageSetup.FitToPagesWide = flags[0];
                sheet.PageSetup.FitToPagesTall = flags[1];
            }

            Common.ExcelApplication.PrintCommunication = true;
        }

        public static void AllLandscape()
        {
            Common.WriteToDebugWindow("");

            Common.ExcelApplication.PrintCommunication = false;

            foreach (Worksheet sheet in Common.ExcelApplication.ActiveWorkbook.Sheets)
            {
                sheet.PageSetup.Orientation = XlPageOrientation.xlLandscape;
            }

            Common.ExcelApplication.PrintCommunication = true;
        }

        public static void Landscape()
        {
            Common.WriteToDebugWindow("");

            Common.ExcelApplication.PrintCommunication = false;

            Worksheet worksheet = Common.ExcelApplication.ActiveSheet as Worksheet;

            worksheet.PageSetup.Orientation = XlPageOrientation.xlLandscape;

            Common.ExcelApplication.PrintCommunication = true;
        }

        public static void FormatMargins()
        {
            Common.WriteToDebugWindow("");

            Common.ExcelApplication.PrintCommunication = false;

            foreach (Worksheet sheet in Common.ExcelApplication.ActiveWorkbook.Sheets)
            {
                sheet.PageSetup.Orientation = XlPageOrientation.xlPortrait;
            }

            Common.ExcelApplication.PrintCommunication = true;
        }

        public static void AllPortrait()
        {
            Common.WriteToDebugWindow("");

            Common.ExcelApplication.PrintCommunication = false;

            foreach (Worksheet sheet in Common.ExcelApplication.ActiveWorkbook.Sheets)
            {
                sheet.PageSetup.Orientation = XlPageOrientation.xlPortrait;
            }

            Common.ExcelApplication.PrintCommunication = true;
        }

        public static void Portrait()
        {
            Common.WriteToDebugWindow("");

            Common.ExcelApplication.PrintCommunication = false;

            Worksheet worksheet = Common.ExcelApplication.ActiveSheet as Worksheet;

            worksheet.PageSetup.Orientation = XlPageOrientation.xlPortrait;

            Common.ExcelApplication.PrintCommunication = true;
        }
    }
}
