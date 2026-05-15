
using Microsoft.Office.Interop.Excel;
using System;
using System.Windows;

using System.Windows.Controls;
using XlHlp = VNC.AddinHelper.Excel;

namespace SupportTools_Excel.User_Interface.User_Controls
{
    /// <summary>
    /// Interaction logic for wucTaskPane_Utilities.xaml
    /// </summary>
    public partial class wucTaskPane_Utilities : UserControl
    {
        #region Fields and Properties


        #endregion

        #region Constructors and Load

        public wucTaskPane_Utilities()
        {
            InitializeComponent();
            LoadControlContents();
        }

        private void LoadControlContents()
        {
            //try
            //{
            //    wucSQLInstance_Picker1.PopulateControlFromFile(Common.cCONFIG_FILE);
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.ToString());
            //}
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            //wucSQLInstance_Picker1.ControlChanged += WucSQLInstance_Picker1_ControlChanged;
            //wucTFSProvider_Picker.ControlChanged += tfsProvider_Picker1_ControlChanged;
        }


        #endregion

        #region Event Handlers
        private void btnFitToPages_Click(object sender, RoutedEventArgs e)
        {
            var s = sender;
            var t = e;

            string tag = "1,0";

            Actions.Excel_PageFormatting.FitToPages(tag);
        }
        private void btnAddHeader_Click(object sender, RoutedEventArgs e)
        {
            Actions.Excel_PageFormatting.AddHeader();
        }
        private void btnAddFooter_Click(object sender, RoutedEventArgs e)
        {
            Actions.Excel_PageFormatting.AddFooter();
        }
        private void btnFormatPortrait_Click(object sender, RoutedEventArgs e)
        {
            Actions.Excel_PageFormatting.FormatPortrait();
        }
        private void btnFormatLandscape_Click(object sender, RoutedEventArgs e)
        {
            Actions.Excel_PageFormatting.FormatLandscape();
        }
        private void btnFormatMargins_Click(object sender, RoutedEventArgs e)
        {

        }
        private void btnFormatFitColumns_Click(object sender, RoutedEventArgs e)
        {

        }
        private void btnFormatFitRows_Click(object sender, RoutedEventArgs e)
        {

        }
        private void btnCreateTableOfContents_Click(object sender, RoutedEventArgs e)
        {
            Actions.Excel_TableOfContents.CreateTableOfContents();
        }
        private void btnAddIDColumn_Click(object sender, RoutedEventArgs e)
        {
            AddIDColumn();
        }

        private void btnCreateDataValidationTable_Click(object sender, RoutedEventArgs e)
        {
            CreateDataValidationTable();
        }

        #endregion

        #region Main Function Routines

        void AddIDColumn()
        {
            Int32 startRow;
            Int32 endRow;
            
            try
            {
                Worksheet ws = (Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
                Range rng = (Range)Globals.ThisAddIn.Application.Selection;
                Range currentCell = Globals.ThisAddIn.Application.ActiveCell;

                startRow = currentCell.Row;
                //endRow = startRow + rng.Rows.Count - 1;
                endRow = XlHlp.FindLast_PopulatedRow_InColumn(currentCell);

                XlHlp.AddColumnToSheet(ws, 1, 10, false, "", XlDirection.xlToRight, XlInsertFormatOrigin.xlFormatFromLeftOrAbove, startRow - 1, null);

                Range idRng = currentCell.Offset[0, -1];

                for (int i = 1; i <= endRow - startRow + 1; i++)
                {
                    ((Range)idRng.Cells[i, 1]).Value = i;
                    //ws.Cells[i + startRow - 1, 1].Value = i;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void CreateDataValidationTable()
        {
            string listName;
            string tableName;
            string validationName;
            string refersTo;

            try
            {
                Range rng = (Range)Globals.ThisAddIn.Application.Selection;

                listName = (string)((Range)rng[1, 1]).Value;
                tableName = GetTableName(listName);
                validationName = GetValidationName(listName);
                refersTo = String.Format("={0}[[{1}]]", tableName, listName);

                CreateTableFromSelection(rng, tableName);
                CreateValidationNameFromTableName(validationName, refersTo);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        #endregion

        #region Utility Routines



        #endregion

        #region Private Methods

        private void CreateTableFromSelection(Range rng, string name)
        {
            Worksheet ws = (Worksheet)Globals.ThisAddIn.Application.ActiveSheet;

            ws.ListObjects.Add(XlListObjectSourceType.xlSrcRange, rng, null, XlYesNoGuess.xlYes).Name = name;
        }

        private void CreateValidationNameFromTableName(string name, string refersTo)
        {
            Globals.ThisAddIn.Application.ActiveWorkbook.Names.Add(Name: name, RefersToR1C1: refersTo);
        }

        private string GetTableName(string name)
        {
            string tableName = "tbl_" + name.Replace(" ", "");

            return tableName;
        }

        private string GetValidationName(string name)
        {
            string validationName = name.Replace(" ", "");

            return validationName;
        }

        //private bool GetDisplayOrientation()
        //{
        //    return (bool)ceOrientOutputVertically.IsChecked;
        //}

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
