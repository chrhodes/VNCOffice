using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

using DevExpress.Xpf.Core;
using DevExpress.Xpf.Editors;
using DevExpress.Xpf.Grid;

namespace SupportTools_Visio.User_Interface
{
    public partial class ToolBoxResources : ResourceDictionary
    {
        protected GridControl dataGrid = null;
        //protected TableView tableView = null;

        protected GridColumn gc_ID = null;
        protected GridColumn gc_SnapShotDate = null;
        protected GridColumn gc_SnapShotDuration = null;
        protected GridColumn gc_SnapShotError = null;

        #region Constructors
        
        public ToolBoxResources()
        {
            InitializeComponent();
        }

        #endregion

        #region Main Function Routines
        
        private static void UpdateVisibilityOfColumns(object sender, string[] columnsToCheck, bool isVisible)
        {
            var vtrootUserControl = VNC.Core.Xaml.PhysicalTree.FindAncestor<UserControl>((DependencyObject)sender);

            if (vtrootUserControl != null)
            {
                foreach (string column in columnsToCheck)
                {
                    var found = vtrootUserControl.FindName(column);

                    if (found != null)
                    {
                        ((GridColumn)found).Visible = isVisible;
                    }
                }
            }
            else
            {
                var vtrootWindow = VNC.Core.Xaml.PhysicalTree.FindAncestor<DXWindow>((DependencyObject)sender);

                foreach (string column in columnsToCheck)
                {
                    var found = vtrootWindow.FindName(column);

                    if (found != null)
                    {
                        ((GridColumn)found).Visible = isVisible;
                    }
                }
            }
        }

        #endregion

        #region Event Handlers

        protected void canAddCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            if (dataGrid != null)
            {
                //dataGrid.CanUserAddRows = true;
            }
        }

        protected void canAddCheckBox_UnChecked(object sender, RoutedEventArgs e)
        {
            if (dataGrid != null)
            {
                //dataGrid.CanUserAddRows = false;
            }
        }

        protected void canDeleteCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            if (dataGrid != null)
            {
                //dataGrid.CanUserDeleteRows = true;
            }
        }

        protected void canDeleteCheckBox_UnChecked(object sender, RoutedEventArgs e)
        {
            if (dataGrid != null)
            {
                //dataGrid.CanUserDeleteRows = false;
            }
        }

        protected void ck_EnableUpdates_Checked(object sender, RoutedEventArgs e)
        {
            if (dataGrid != null)
            {
                //dataGrid.CanUserDeleteRows = true;
            }
        }

        protected void ck_EnableUpdates_UnChecked(object sender, RoutedEventArgs e)
        {
            if (dataGrid != null)
            {
                //dataGrid.CanUserDeleteRows = false;
            }
        }

        private void OnDeleteAllRows(object sender, RoutedEventArgs e)
        {
            // TODO(crhodes): Determine how to get to right place.
        }

        private void OnDeleteRow(object sender, RoutedEventArgs e)
        {
            // TODO(crhodes): Determine how to get to right place.
        }

        protected void OnDisplayEnvironmentColumns(object sender, RoutedEventArgs e)
        {
            var vtroot = VNC.Core.Xaml.PhysicalTree.FindAncestor<UserControl>((DependencyObject)sender);

            string[] columnsToCheck = 
            { 
               "gc_ADDomain", "gc_Environment", "gc_SecurityZone"
            };

            bool isVisible = (bool)((CheckBox)e.Source).IsChecked;

            UpdateVisibilityOfColumns(sender, columnsToCheck, isVisible);
        }

        protected void OnDisplayIDColumns(object sender, RoutedEventArgs e)
        {
            //var foo = sender;
            //var foo2 = ((FrameworkElement)sender).TemplatedParent;
            //ContentControl foo3 = (ContentControl)foo2;

            //// This got us the UserControl hosting this mess.
            //var vtroot = VisualTreeHelperExtensions.FindAncestor<UserControl>((DependencyObject)sender);

            //var found = vtroot.FindName("gc_ID");
            //((GridColumn)found).Visible = true;
            //return;

            //// This returned null
            //var ltroot = LogicalTreeHelperExtensions.FindAncestor<UserControl>((DependencyObject)sender);

            //// Need to find out the controls named gc_ID*

            ////HideIDColumns(((CheckBox)sender).IsChecked);
            //if (gc_ID != null)
            //{
            //    gc_ID.Visible = true;
            //}

            string[] columnsToCheck = 
            { 
              "gc_ID",  "gc_ID1", "gc_ID2", "gc_ID3", "gc_ID4", 
              "gc_ID5", "gc_ID6", "gc_ID7", "gc_ID8", "gc_ID9", 
              "gc_ID10", "gc_ID11"
            };

            bool isVisible = (bool)((CheckBox)e.Source).IsChecked;

            UpdateVisibilityOfColumns(sender, columnsToCheck, isVisible);
        }

        protected void OnDisplayOperatingSystemColumns(object sender, RoutedEventArgs e)
        {
            var vtroot = VNC.Core.Xaml.PhysicalTree.FindAncestor<UserControl>((DependencyObject)sender);

            string[] columnsToCheck = 
            { 
               "gc_OSVersion", "gc_Processors", "gc_PhysicalMemory"
            };

            bool isVisible = (bool)((CheckBox)e.Source).IsChecked;

            UpdateVisibilityOfColumns(sender, columnsToCheck, isVisible);
        }

        protected void OnDisplayExpandColumns(object sender, RoutedEventArgs e)
        {
            var vtroot = VNC.Core.Xaml.PhysicalTree.FindAncestor<UserControl>((DependencyObject)sender);

            string[] columnsToCheck = 
            { 
               "gc_DefaultInstanceExpandMask", 
               "gc_DefaultDatabaseExpandMask", "gc_DefaultJobExpandMask",
               "gc_ExpandContent", "gc_ExpandStorage", "gc_ExpandJobServer", 
               "gc_ExpandLogins", "gc_ExpandRoles", "gc_ExpandServerRoles", "gc_ExpandTriggers",
               "gc_ExpandUserDefinedFunctions", "gc_ExpandUsers", "gc_ExpandViews", "gc_ExpandTables",
               "gc_ExpandDataFiles", "gc_ExpandFileGroups", "gc_ExpandLogFiles", "gc_ExpandStoredProcedures"
            };

            bool isVisible = (bool)((CheckBox)e.Source).IsChecked;

            UpdateVisibilityOfColumns(sender, columnsToCheck, isVisible);
        }

        protected void OnDisplayExtendedPropertyColumns(object sender, RoutedEventArgs e)
        {
            var vtroot = VNC.Core.Xaml.PhysicalTree.FindAncestor<UserControl>((DependencyObject)sender);

            string[] columnsToCheck = 
            { 
               "gc_EP_Area", 
               "gc_EP_Team", "gc_EP_PrimaryDBContact",
               "gc_EP_DBApprover", "gc_EP_DRTier"
            };

            bool isVisible = (bool)((CheckBox)e.Source).IsChecked;

            UpdateVisibilityOfColumns(sender, columnsToCheck, isVisible);
        }

        protected void OnDisplayIOColumns(object sender, RoutedEventArgs e)
        {
            string[] columnsToCheck = 
            { 
              "gc_BytesReadFromDisk",  "gc_BytesWrittenToDisk", 
              "gc_NumberOfDiskReads",  "gc_NumberOfDiskWrites"

            };

            bool isVisible = (bool)((CheckBox)e.Source).IsChecked;

            UpdateVisibilityOfColumns(sender, columnsToCheck, isVisible);
        }

        protected void OnDisplaySizeColumns(object sender, RoutedEventArgs e)
        {
            string[] columnsToCheck = 
            { 
              "gc_AvailableSpace",  "gc_MaxSize", "gc_Size",
              "gc_UsedSpace",  "gc_VolumeFreeSpace"

            };

            bool isVisible = (bool)((CheckBox)e.Source).IsChecked;

            UpdateVisibilityOfColumns(sender, columnsToCheck, isVisible);
        }

        protected void OnDisplaySnapShotColumns(object sender, RoutedEventArgs e)
        {
            string[] columnsToCheck = 
            { 
              "gc_SnapShotDate",  "gc_SnapShotDuration", "gc_SnapShotError",
              "gc_SnapShotDate1",  "gc_SnapShotDuration1", "gc_SnapShotError1",
              "gc_SnapShotDate2",  "gc_SnapShotDuration2", "gc_SnapShotError2",
              "gc_SnapShotDate3",  "gc_SnapShotDuration3", "gc_SnapShotError3",
              "gc_SnapShotDate4",  "gc_SnapShotDuration4", "gc_SnapShotError4",
              "gc_SnapShotDate5",  "gc_SnapShotDuration5", "gc_SnapShotError5",
              "gc_SnapShotDate6",  "gc_SnapShotDuration6", "gc_SnapShotError6",
              "gc_SnapShotDate7",  "gc_SnapShotDuration7", "gc_SnapShotError7",
              "gc_SnapShotDate8",  "gc_SnapShotDuration8", "gc_SnapShotError8",
              "gc_SnapShotDate9",  "gc_SnapShotDuration9", "gc_SnapShotError9",
              "gc_SnapShotDate10",  "gc_SnapShotDuration10", "gc_SnapShotError10"
            };

            bool isVisible = (bool)((CheckBox)e.Source).IsChecked;

            UpdateVisibilityOfColumns(sender, columnsToCheck, isVisible);
        }

        protected void OnDisplaySQLVersionColumns(object sender, RoutedEventArgs e)
        {
            var vtroot = VNC.Core.Xaml.PhysicalTree.FindAncestor<UserControl>((DependencyObject)sender);

            string[] columnsToCheck = 
            { 
               "gc_IsClustered", "gc_Edition", "gc_Version",
               "gc_ProductLevel", "gc_BrowserStartMode", "gc_Collation",
               "gc_EngineEdition", "gc_PerfMonMode", "gc_Product",
               "gc_ServiceAccount", "gc_ServiceName"
            };

            bool isVisible = (bool)((CheckBox)e.Source).IsChecked;

            UpdateVisibilityOfColumns(sender, columnsToCheck, isVisible);
        }

        //protected void OnSaveToCSV(object sender, RoutedEventArgs e)
        //{
        //    TableView tableView = null;
        //    GridControl gridControl = null;

        //    var vtrootUserControl = VNC.Core.Xaml.PhysicalTree.FindAncestor<UserControl>((DependencyObject)sender);

        //    if (vtrootUserControl != null)
        //    {
        //        gridControl = (GridControl)vtrootUserControl.FindName("dataGrid");
        //        tableView = (TableView)vtrootUserControl.FindName("tableView");

        //    }
        //    else
        //    {
        //        var vtrootWindow = VNC.Core.Xaml.PhysicalTree.FindAncestor<DXWindow>((DependencyObject)sender);

        //        gridControl = (GridControl)vtrootWindow.FindName("dataGrid");
        //        tableView = (TableView)vtrootWindow.FindName("tableView");
        //    }
            
        //    if (tableView != null)
        //    {
        //        PostDataToCSV(tableView);
        //    }
        //}

        //protected void OnSaveToXLS(object sender, RoutedEventArgs e)
        //{
        //    TableView tableView = null;
        //    GridControl gridControl = null;

        //    var vtrootUserControl = VNC.Core.Xaml.PhysicalTree.FindAncestor<UserControl>((DependencyObject)sender);

        //    if (vtrootUserControl != null)
        //    {
        //        gridControl = (GridControl)vtrootUserControl.FindName("dataGrid");
        //        tableView = (TableView)vtrootUserControl.FindName("tableView");

        //    }
        //    else
        //    {
        //        var vtrootWindow = VNC.Core.Xaml.PhysicalTree.FindAncestor<DXWindow>((DependencyObject)sender);

        //        gridControl = (GridControl)vtrootWindow.FindName("dataGrid");
        //        tableView = (TableView)vtrootWindow.FindName("tableView");
        //    }

        //    if (tableView != null)
        //    {
        //        PostDataToXLS(tableView);
        //    }
        //}

        //protected void OnSaveToXLSX(object sender, RoutedEventArgs e)
        //{
        //    TableView tableView = null;
        //    GridControl gridControl = null;

        //    var vtrootUserControl = VNC.Core.Xaml.PhysicalTree.FindAncestor<UserControl>((DependencyObject)sender);

        //    if (vtrootUserControl != null)
        //    {
        //        gridControl = (GridControl)vtrootUserControl.FindName("dataGrid");
        //        tableView = (TableView)vtrootUserControl.FindName("tableView");

        //    }
        //    else
        //    {
        //        var vtrootWindow = VNC.Core.Xaml.PhysicalTree.FindAncestor<DXWindow>((DependencyObject)sender);

        //        gridControl = (GridControl)vtrootWindow.FindName("dataGrid");
        //        tableView = (TableView)vtrootWindow.FindName("tableView");
        //    }

        //    if (tableView != null)
        //    {
        //        PostDataToXLSX(tableView);
        //    }
        //}

        protected void readOnlyCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            if (dataGrid != null)
            {
                //dataGrid.IsReadOnly = true;
            }
        }

        protected void readOnlyCheckBox_UnChecked(object sender, RoutedEventArgs e)
        {
            if (dataGrid != null)
            {
                //dataGrid.IsReadOnly = false;
            }
        }

        //protected void PostDataToCSV(TableView tableView)
        //{
        //    var dlg = new Microsoft.Win32.SaveFileDialog();
        //    dlg.FileName = "ServersDataGrid";
        //    dlg.DefaultExt = ".xls";
        //    dlg.Filter = "CSV Documents (*.csv)|*.csv";

        //    Nullable<bool> result = dlg.ShowDialog();

        //    if (result == true)
        //    {
        //        tableView.ExportToCsv(dlg.FileName);
        //    }

        //    //Common.ApplicationDataSet.Instances.WriteXml(outputFile);
        //}

        //protected void PostDataToXLSX(TableView tableView)
        //{
        //    var dlg = new Microsoft.Win32.SaveFileDialog();
        //    dlg.FileName = "ServersDataGrid";
        //    dlg.DefaultExt = ".xlsx";
        //    dlg.Filter = "XLSX Documents (*.xlsx)|*.xlsx";

        //    Nullable<bool> result = dlg.ShowDialog();

        //    if (result == true)
        //    {
        //        tableView.ExportToXlsx(dlg.FileName);
        //    }

        //    //Common.ApplicationDataSet.Instances.WriteXml(outputFile);
        //}

        //protected void PostDataToXLS(TableView tableView)
        //{
        //    var dlg = new Microsoft.Win32.SaveFileDialog();
        //    dlg.FileName = "ServersDataGrid";
        //    dlg.DefaultExt = ".xls";
        //    dlg.Filter = "XLS Documents (*.xls)|*.xls";

        //    Nullable<bool> result = dlg.ShowDialog();

        //    if (result == true)
        //    {
        //        tableView.ExportToXls(dlg.FileName);
        //    }

        //    //Common.ApplicationDataSet.Instances.WriteXml(outputFile);
        //}

        private void RowTemplateComboBox_SelectionChanged(object sender, DevExpress.Xpf.Editors.EditValueChangedEventArgs e)
        {
            ListBoxEdit lbe = (ListBoxEdit)sender;

            TableView tableView = null;
            GridControl gridControl = null;

            var vtrootUserControl = VNC.Core.Xaml.PhysicalTree.FindAncestor<UserControl>((DependencyObject)sender);

            if (vtrootUserControl != null)
            {
                gridControl = (GridControl)vtrootUserControl.FindName("dataGrid");
                tableView = (TableView)vtrootUserControl.FindName("tableView");

            }
            else
            {
                var vtrootWindow = VNC.Core.Xaml.PhysicalTree.FindAncestor<DXWindow>((DependencyObject)sender);

                gridControl = (GridControl)vtrootWindow.FindName("dataGrid");
                tableView = (TableView)vtrootWindow.FindName("tableView");
            }

            switch (lbe.SelectedItem.ToString())
            {
                case "Tooltip":
                    // This uses the resources from the usercontrol/window.
                    //var  tmplt =  (DataTemplate)vtrootUserControl.Resources["rowTooltipDetailTemplate"];
                    var tmplt = (DataTemplate)Application.PrismApp.Current.Resources["rowTooltipDetailTemplate"];
                    //var tmplt = (DataTemplate)Application.Current.Resources["rowTooltipDetailTemplate"];
                    tableView.DataRowTemplate = tmplt;
                    break;

                case "SelectedRowDetails":
                    //tableView.DataRowTemplate = (DataTemplate)Resources["rowSelectedDetailTemplate"];
                    // This uses the Application scope resources.
                    var tmplt1 = (DataTemplate)Application.PrismApp.Current.Resources["rowSelectedDetailTemplate"];
                    //var tmplt1 = (DataTemplate)Application.Current.Resources["rowSelectedDetailTemplate"];
                    tableView.DataRowTemplate = tmplt1;
                    break;

                case "RowDetails":
                    // This uses FindResource.
                    var tmplt2 = (DataTemplate)vtrootUserControl.FindResource("rowDetailTemplate");
                    tableView.DataRowTemplate = tmplt2;
                    break;

                case "None":
                    tableView.ClearValue(DevExpress.Xpf.Grid.TableView.DataRowTemplateProperty);
                    break;

                default:
                    break;
            }
        }

        private void saveButton_Click(object sender, RoutedEventArgs e)
        {
            //Common.ApplicationDataSet.DatabaseInfoTA.Update(Common.ApplicationDataSet.DatabaseInfo);
        }

        private void undoButton_Click(object sender, RoutedEventArgs e)
        {
            //Common.ApplicationDataSet.DatabaseInfo.RejectChanges();
        }

        #endregion

        //public static class VisualTreeHelperExtensions
        //{
        //    public static T FindAncestor<T>(DependencyObject dependencyObject)
        //        where T : class
        //    {
        //        DependencyObject target = dependencyObject;
        //        do
        //        {
        //            target = VisualTreeHelper.GetParent(target);
        //        }
        //        while (target != null && !(target is T));
        //        return target as T;
        //    }
        //}

        //public static class LogicalTreeHelperExtensions
        //{
        //    public static T FindAncestor<T>(DependencyObject dependencyObject)
        //        where T : class
        //    {
        //        DependencyObject target = dependencyObject;
        //        do
        //        {
        //            target = LogicalTreeHelper.GetParent(target);
        //        }
        //        while (target != null && !(target is T));
        //        return target as T;
        //    }
        //}

    }
}
