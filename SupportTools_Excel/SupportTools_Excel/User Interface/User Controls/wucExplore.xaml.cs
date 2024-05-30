using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

using Microsoft.Office.Interop.Excel;


using XlHlp = VNC.AddinHelper.Excel;

using System.Text.RegularExpressions;
using Microsoft.Win32;
using System.IO;
using System.Diagnostics;

namespace SupportTools_Excel.User_Interface.User_Controls
{
    public partial class wucExplore : UserControl
    {
        //private static int CLASS_BASE_ERRORNUMBER = ErrorNumbers.APPERROR;
        private const string LOG_APPNAME = Common.LOG_CATEGORY;

        #region Constructors

        public wucExplore()
        {
#if TRACE
            //long startTicks = VNC.Log.Trace5("Start", LOG_APPNAME);
#endif
            try
            {
                InitializeComponent();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                MessageBox.Show(ex.InnerException.ToString());
            }
#if TRACE
            //VNC.Log.Trace5("End", LOG_APPNAME, startTicks);
#endif
        }

        #endregion

        #region Initialization

        private void OnLoaded(object sender, RoutedEventArgs e)
        {

        }

        #endregion

        #region Event Handlers

        private void OnCustomColumnDisplayText(object sender, DevExpress.Xpf.Grid.CustomColumnDisplayTextEventArgs e)
        {
            //CustomFormat.FormatStorageColumns(e);
        }

        #endregion

        private void CustomUnboundColumnData(object sender, DevExpress.Xpf.Grid.GridColumnDataEventArgs e)
        {
            //UnboundColumns.GetEnvironmentInstanceDatabaseColumns(e);
        }

        #region Main Function Routines



        #endregion

        private void XXX_Picker_ControlChanged()
        {

        }

        private void btnOne_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("btnOne_Click");
        }

        private void btnTwo_Click(object sender, RoutedEventArgs e)
        {

        }

        private void teStartRowCol_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application app = Globals.ThisAddIn.Application;
            teStartRow.Text = app.ActiveCell.Row.ToString();
            teStartCol.Text = app.ActiveCell.Column.ToString();
        }

        private void btnTimeRange_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application app = Globals.ThisAddIn.Application;

            teDuration.Clear();

            int insertRow;
            int insertCol;
            int iterations;

            if (!int.TryParse(teStartRow.Text, out insertRow))
            {
                MessageBox.Show("Illegal StartRow");
                teStartRow.Focus();
                return;
            }

            if (!int.TryParse(teStartCol.Text, out insertCol))
            {
                MessageBox.Show("Illegal StartCol");
                teStartCol.Focus();
                return;
            }

            if (!int.TryParse(teIterations.Text, out iterations))
            {
                MessageBox.Show("Illegal Iterations");
                teIterations.Focus();
                return;
            }

            if (!(bool)ceScreenUpdates.IsChecked)
            {
                XlHlp.ScreenUpdatesOff();
            }

            if (!(bool)ceCalculations.IsChecked)
            {
                XlHlp.CalculationsOff();
            }

            long startTicks = XlHlp.DisplayInWatchWindow("Start");

            if ((bool)ceInsertDescending.IsChecked)
            {
                insertRow += iterations - 1;

                for (int i = 0; i < iterations; i++)
                {
                    app.Cells[insertRow, insertCol] = string.Format("{0,5}-1", insertRow);
                    insertRow--;
                }
            }
            else
            {
                for (int i = 0; i < iterations; i++)
                {
                    app.Cells[insertRow, insertCol] = string.Format("{0,5}-1", insertRow);
                    insertRow++;
                }
            }

            long endTicks = XlHlp.DisplayInWatchWindow("End", startTicks);

            teDuration.Text = ((endTicks - startTicks) / ((double)Stopwatch.Frequency)).ToString();

            XlHlp.ScreenUpdatesOn(true);
            XlHlp.CalculationsOn();
        }

        private void btnTimeRangeOffset_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application app = Globals.ThisAddIn.Application;

            teDuration.Text = "";

            int insertRow;
            int insertCol;
            int iterations;

            if (!int.TryParse(teStartRow.Text, out insertRow))
            {
                MessageBox.Show("Illegal StartRow");
                teStartRow.Focus();
                return;
            }

            if (!int.TryParse(teStartCol.Text, out insertCol))
            {
                MessageBox.Show("Illegal StartCol");
                teStartCol.Focus();
                return;
            }

            if (!int.TryParse(teIterations.Text, out iterations))
            {
                MessageBox.Show("Illegal Iterations");
                teIterations.Focus();
                return;
            }

            if (!(bool)ceScreenUpdates.IsChecked)
            {
                XlHlp.ScreenUpdatesOff();
            }

            if (!(bool)ceCalculations.IsChecked)
            {
                XlHlp.CalculationsOff();
            }

            long startTicks = XlHlp.DisplayInWatchWindow("Start");

            if ((bool)ceInsertDescending.IsChecked)
            {
                Range rng = (Range)app.Cells[insertRow, insertCol];
                int rowOffset = iterations - 1;

                for (int i = 0; i < iterations; i++)
                {
                    ((Range)rng.Offset[rowOffset, 0]).Value = string.Format("{0,5}-1", insertRow + rowOffset);
                    rowOffset--;
                }
            }
            else
            {
                Range rng = (Range)app.Cells[insertRow, insertCol];
                int rowOffset = 0;

                for (int i = 0; i < iterations; i++)
                {
                    ((Range)rng.Offset[rowOffset, 0]).Value = string.Format("{0,5}-1", insertRow + rowOffset);
                    rowOffset++;
                }
            }

            long endTicks = XlHlp.DisplayInWatchWindow("End", startTicks);

            teDuration.Text = ((endTicks - startTicks) / ((double)Stopwatch.Frequency)).ToString();

            XlHlp.ScreenUpdatesOn(true);
            XlHlp.CalculationsOn();
        }

        private void btnTimeInsertAt_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application app = Globals.ThisAddIn.Application;

            teDuration.Clear();

            int insertRow;
            int insertCol;
            int iterations;

            if (!int.TryParse(teStartRow.Text, out insertRow))
            {
                MessageBox.Show("Illegal StartRow");
                teStartRow.Focus();
                return;
            }

            if (!int.TryParse(teStartCol.Text, out insertCol))
            {
                MessageBox.Show("Illegal StartCol");
                teStartCol.Focus();
                return;
            }

            if (!int.TryParse(teIterations.Text, out iterations))
            {
                MessageBox.Show("Illegal Iterations");
                teIterations.Focus();
                return;
            }

            if (!(bool)ceScreenUpdates.IsChecked)
            {
                XlHlp.ScreenUpdatesOff();
            }

            if (!(bool)ceCalculations.IsChecked)
            {
                XlHlp.CalculationsOff();
            }

            long startTicks = XlHlp.DisplayInWatchWindow("Start");

            if ((bool)ceInsertDescending.IsChecked)
            {
                XlHlp.XlLocation insertAt = new XlHlp.XlLocation((Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet, insertRow, insertCol);

                insertRow += iterations;

                for (int i = 0; i < iterations; i++)
                {
                    XlHlp.AddContentToCell(insertAt.AddRowX(), string.Format("{0,5}-1", insertRow + i), null);
                }
            }
            else
            {
                XlHlp.XlLocation insertAt = new XlHlp.XlLocation((Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet, insertRow, insertCol);

                for (int i = 0; i < iterations; i++)
                {
                    XlHlp.AddContentToCell(insertAt.AddRowX(), string.Format("{0,5}-1", insertRow + i), null);
                }
            }

            long endTicks = XlHlp.DisplayInWatchWindow("End", startTicks);

            teDuration.Text = ((endTicks - startTicks) / ((double)Stopwatch.Frequency)).ToString();

            XlHlp.ScreenUpdatesOn(true);
            XlHlp.CalculationsOn();
        }

        private void btnTimeInsertAtLight_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application app = Globals.ThisAddIn.Application;

            teDuration.Clear();

            int insertRow;
            int insertCol;
            int iterations;

            if (!int.TryParse(teStartRow.Text, out insertRow))
            {
                MessageBox.Show("Illegal StartRow");
                teStartRow.Focus();
                return;
            }

            if (!int.TryParse(teStartCol.Text, out insertCol))
            {
                MessageBox.Show("Illegal StartCol");
                teStartCol.Focus();
                return;
            }

            if (!int.TryParse(teIterations.Text, out iterations))
            {
                MessageBox.Show("Illegal Iterations");
                teIterations.Focus();
                return;
            }

            if (!(bool)ceScreenUpdates.IsChecked)
            {
                XlHlp.ScreenUpdatesOff();
            }

            if (!(bool)ceCalculations.IsChecked)
            {
                XlHlp.CalculationsOff();
            }

            long startTicks = XlHlp.DisplayInWatchWindow("Start");

            if ((bool)ceInsertDescending.IsChecked)
            {
                XlHlp.XlLocation insertAt = new XlHlp.XlLocation((Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet, insertRow, insertCol);

                insertRow += iterations;

                for (int i = 0; i < iterations; i++)
                {
                    XlHlp.AddContentToCell(insertAt.AddRowX(), string.Format("{0,5}-1", insertRow + i));
                }
            }
            else
            {
                XlHlp.XlLocation insertAt = new XlHlp.XlLocation((Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet, insertRow, insertCol);

                for (int i = 0; i < iterations; i++)
                {
                    XlHlp.AddContentToCell(insertAt.AddRowX(), string.Format("{0,5}-1", insertRow + i));
                }
            }

            long endTicks = XlHlp.DisplayInWatchWindow("End", startTicks);

            teDuration.Text = ((endTicks - startTicks) / ((double)Stopwatch.Frequency)).ToString();

            XlHlp.ScreenUpdatesOn(true);
            XlHlp.CalculationsOn();
        }
    }
}
