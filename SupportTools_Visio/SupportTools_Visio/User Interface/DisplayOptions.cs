using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

using DevExpress.Xpf.Core;
using DevExpress.Xpf.Grid;

using VNC;

namespace SupportTools_Visio.User_Interface
{
    public partial class DisplayOptions : ResourceDictionary
    {
        //protected GridControl dataGrid = null;

        //protected GridColumn gc_ID = null;
        //protected GridColumn gc_SnapShotDate = null;
        //protected GridColumn gc_SnapShotDuration = null;
        //protected GridColumn gc_SnapShotError = null;

        bool _isDisplayed_FilterPanel = false;
        bool _isDisplayed_GroupByPanel = false;

        #region Constructors

        public DisplayOptions()
        {
            InitializeComponent();
        }

        #endregion

        #region Event Handlers

        protected void canAddCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            //if (dataGrid != null)
            //{
            //    //dataGrid.CanUserAddRows = true;
            //}
        }

        protected void canAddCheckBox_UnChecked(object sender, RoutedEventArgs e)
        {
            //if (dataGrid != null)
            //{
            //    //dataGrid.CanUserAddRows = false;
            //}
        }

        protected void canDeleteCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            //if (dataGrid != null)
            //{
            //    //dataGrid.CanUserDeleteRows = true;
            //}
        }

        protected void canDeleteCheckBox_UnChecked(object sender, RoutedEventArgs e)
        {
            //if (dataGrid != null)
            //{
            //    //dataGrid.CanUserDeleteRows = false;
            //}
        }

        protected void OnDisplayBackupColumns(object sender, RoutedEventArgs e)
        {
            var vtroot = VNC.Core.Xaml.PhysicalTree.FindAncestor<UserControl>((DependencyObject)sender);

            string[] columnsToCheck = 
            { 
               "gc_LastBackupDate", "gc_LastDifferentialBackupDate", "gc_LastLogBackupDate",
               "gc_RecoveryModel"
            };

            bool isVisible = (bool)((CheckBox)e.Source).IsChecked;

            UpdateVisibilityOfColumns(sender, columnsToCheck, isVisible);
        }

        protected void OnDisplayEnvironmentColumns(object sender, RoutedEventArgs e)
        {
            var vtroot = VNC.Core.Xaml.PhysicalTree.FindAncestor<UserControl>((DependencyObject)sender);

            string[] columnsToCheck = 
            { 
               "gc_ADDomain", "gc_Environment", "gc_SecurityZone",
               "gc_ADDomain2", "gc_Environment2", "gc_SecurityZone2"
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
              "gc_ID",  "gc_ID1", "gc_ID2", "gc_ID3", "gc_ID3a", 
              "gc_ID4", "gc_ID4a", 
              "gc_ID5", "gc_ID6", "gc_ID7", "gc_ID8", "gc_ID9", 
              "gc_ID10", "gc_ID11", "gc_ID12", "gc_ID12a",
              "gc_Server_ID", "gc_Instance_ID", "gc_Database_ID",
              "gc_DBDataFile_ID"
            };

            bool isVisible = (bool)((CheckBox)e.Source).IsChecked;

            UpdateVisibilityOfColumns(sender, columnsToCheck, isVisible);
        }

        protected void OnDisplayMonitoredColumns(object sender, RoutedEventArgs e)
        {
            var vtroot = VNC.Core.Xaml.PhysicalTree.FindAncestor<UserControl>((DependencyObject)sender);

            string[] columnsToCheck = 
            { 
               "gc_IsMonitored", "gc_IsMonitored_Instance"
            };

            bool isVisible = (bool)((CheckBox)e.Source).IsChecked;

            UpdateVisibilityOfColumns(sender, columnsToCheck, isVisible);
        }

        protected void OnDisplayOperatingSystemColumns(object sender, RoutedEventArgs e)
        {
            var vtroot = VNC.Core.Xaml.PhysicalTree.FindAncestor<UserControl>((DependencyObject)sender);

            string[] columnsToCheck = 
            { 
               "gc_Platform", "gc_OSVersion", "gc_Processors", "gc_PhysicalMemory",
               "gc_Platform2", "gc_OSVersion2", "gc_Processors2", "gc_PhysicalMemory2"
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
               "gc_ExpandLinkedServers", "gc_ExpandLogins", "gc_ExpandRoles", "gc_ExpandServerRoles", "gc_ExpandTriggers",
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
               "gc_EP_DBApprover", "gc_EP_DRTier",
               "gc_EP_Purpose"
            };

            bool isVisible = (bool)((CheckBox)e.Source).IsChecked;

            UpdateVisibilityOfColumns(sender, columnsToCheck, isVisible);
        }

        protected void OnDisplayIOColumns(object sender, RoutedEventArgs e)
        {
            string[] columnsToCheck = 
            { 
                "gc_NumberOfDiskReads",  "gc_BytesReadFromDisk",   "gc_NumberOfDiskWrites",  "gc_BytesWrittenToDisk",
                "gc_NumberOfDiskReads1", "gc_BytesReadFromDisk1",  "gc_NumberOfDiskWrites1", "gc_BytesWrittenToDisk1",  
                "gc_NumberOfDiskReads2", "gc_BytesReadFromDisk2",  "gc_NumberOfDiskWrites2", "gc_BytesWrittenToDisk2",  
                "gc_NumberOfDiskReads3", "gc_BytesReadFromDisk3",  "gc_NumberOfDiskWrites3", "gc_BytesWrittenToDisk3",  
                "gc_NumberOfDiskReads4", "gc_BytesReadFromDisk4",  "gc_NumberOfDiskWrites4", "gc_BytesWrittenToDisk4",  
                "gc_NumberOfDiskReads5", "gc_BytesReadFromDisk5",  "gc_NumberOfDiskWrites5", "gc_BytesWrittenToDisk5",  
                "gc_NumberOfDiskReads6", "gc_BytesReadFromDisk6",  "gc_NumberOfDiskWrites6", "gc_BytesWrittenToDisk6",  
                "gc_NumberOfDiskReads7", "gc_BytesReadFromDisk7",  "gc_NumberOfDiskWrites7", "gc_BytesWrittenToDisk7",  
                "gc_NumberOfDiskReads8", "gc_BytesReadFromDisk8",  "gc_NumberOfDiskWrites8", "gc_BytesWrittenToDisk8",  
                "gc_NumberOfDiskReads9", "gc_BytesReadFromDisk9",  "gc_NumberOfDiskWrites9", "gc_BytesWrittenToDisk9"
            };

            bool isVisible = (bool)((CheckBox)e.Source).IsChecked;

            UpdateVisibilityOfColumns(sender, columnsToCheck, isVisible);
        }

        protected void OnDisplayNotesColumns(object sender, RoutedEventArgs e)
        {
            string[] columnsToCheck = 
            { 
              "gc_Notes", "gc_Notes_Instance"
            };

            bool isVisible = (bool)((CheckBox)e.Source).IsChecked;

            UpdateVisibilityOfColumns(sender, columnsToCheck, isVisible);
        }

        protected void OnDisplaySizeColumns(object sender, RoutedEventArgs e)
        {
            string[] columnsToCheck = 
            { 
                "gc_Size",  "gc_UsedSpace",  "gc_AvailableSpace",  "gc_MaxSize",  "gc_VolumeFreeSpace",
                "gc_Size1", "gc_UsedSpace1", "gc_AvailableSpace1", "gc_MaxSize1", "gc_VolumeFreeSpace1",
                "gc_Size2", "gc_UsedSpace2", "gc_AvailableSpace2", "gc_MaxSize2", "gc_VolumeFreeSpace2",
                "gc_Size3", "gc_UsedSpace3", "gc_AvailableSpace3", "gc_MaxSize3", "gc_VolumeFreeSpace3",
                "gc_Size4", "gc_UsedSpace4", "gc_AvailableSpace4", "gc_MaxSize4", "gc_VolumeFreeSpace4",
                "gc_Size5", "gc_UsedSpace5", "gc_AvailableSpace5", "gc_MaxSize5", "gc_VolumeFreeSpace5",
                "gc_Size6", "gc_UsedSpace6", "gc_AvailableSpace6", "gc_MaxSize6", "gc_VolumeFreeSpace6",
                "gc_Size7", "gc_UsedSpace7", "gc_AvailableSpace7", "gc_MaxSize7", "gc_VolumeFreeSpace7",
                "gc_Size8", "gc_UsedSpace8", "gc_AvailableSpace8", "gc_MaxSize8", "gc_VolumeFreeSpace8",
                "gc_Size9", "gc_UsedSpace9", "gc_AvailableSpace9", "gc_MaxSize9", "gc_VolumeFreeSpace9",
                
              "gc_DataSpaceUsage", "gc_IndexSpaceUsage", "gc_SpaceAvailable",
              "gc_DefaultFileGroup",
              "gc_DataSpaceUsed", "gc_RowCount",
              "gc_Total_DataSpaceUsage", "gc_Total_IndexSpaceUsage", "gc_Total_Size", "gc_Total_SpaceAvailable"
            };

            bool isVisible = (bool)((CheckBox)e.Source).IsChecked;

            UpdateVisibilityOfColumns(sender, columnsToCheck, isVisible);
        }

        protected void OnDisplaySnapShotColumns(object sender, RoutedEventArgs e)
        {
            string[] columnsToCheck = 
            { 
              "gc_SnapShotDate",  "gc_SnapShotDuration", "gc_SnapShotError", "gc_NotFound",
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

        protected void readOnlyCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            //if (dataGrid != null)
            //{
            //    //dataGrid.IsReadOnly = true;
            //}
        }

        protected void readOnlyCheckBox_UnChecked(object sender, RoutedEventArgs e)
        {
            //if (dataGrid != null)
            //{
            //    //dataGrid.IsReadOnly = false;
            //}
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
                        if (isVisible)
                        {
                            ((GridColumn)found).Visible = true;
                            ((GridColumn)found).VisibleIndex = Math.Abs(((GridColumn)found).VisibleIndex);
                            ;
                        }
                        else
                        {
                            ((GridColumn)found).Visible = false;
                            ((GridColumn)found).VisibleIndex = - Math.Abs(((GridColumn)found).VisibleIndex);
                        }
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
                        //((GridColumn)found).Visible = isVisible;
                        if (isVisible)
                        {
                            ((GridColumn)found).Visible = true;
                            ((GridColumn)found).VisibleIndex = Math.Abs(((GridColumn)found).VisibleIndex);
                            ;
                        }
                        else
                        {
                            ((GridColumn)found).Visible = false;
                            ((GridColumn)found).VisibleIndex = -Math.Abs(((GridColumn)found).VisibleIndex);
                        }
                    }
                }
            }
        }


        private static void UpdateRowDetailTemplate(object sender, int index)
        {

        }

        private static void UpdateVisibilityOfFilterPanel(object sender, bool isVisible)
        {
            var vtrootUserControl = VNC.Core.Xaml.PhysicalTree.FindAncestor<UserControl>((DependencyObject)sender);

            if (vtrootUserControl != null)
            {
                var tv = VNC.Core.Xaml.PhysicalTree.FindChild<TableView>(vtrootUserControl, "tableView");

                if (isVisible)
                {
                    tv.ShowFilterEditor(null);
                    //tv.ShowFilterPanelMode = ShowFilterPanelMode.ShowAlways;
                }
                else
                {
                    tv.ShowFilterEditor(null);
                    //tv.ShowFilterPanelMode = ShowFilterPanelMode.Never;
                }
            }
            else
            {
                var vtrootWindow = VNC.Core.Xaml.PhysicalTree.FindAncestor<DXWindow>((DependencyObject)sender);

                var tv = VNC.Core.Xaml.PhysicalTree.FindChild<TableView>(vtrootWindow, "tableView");

                if (isVisible)
                {
                    tv.ShowFilterEditor(null);
                    //tv.ShowFilterPanelMode = ShowFilterPanelMode.ShowAlways;
                }
                else
                {
                    tv.ShowFilterEditor(null);
                    //tv.ShowFilterPanelMode = ShowFilterPanelMode.Never;
                }
            }
        }

        private static void UpdateVisibilityOfGroupByPanel(object sender, bool isVisible)
        {
            var vtrootUserControl = VNC.Core.Xaml.PhysicalTree.FindAncestor<UserControl>((DependencyObject)sender);

            if (vtrootUserControl != null)
            {
                var dg = VNC.Core.Xaml.PhysicalTree.FindChild<DataGrid>(vtrootUserControl, "dataGrid");
                var tv = VNC.Core.Xaml.PhysicalTree.FindChild<TableView>(vtrootUserControl, "tableView");

                if (isVisible)
                {
                    tv.ShowGroupPanel = true;
                }
                else
                {
                    tv.ShowGroupPanel = false;
                }
            }
            else
            {
                var vtrootWindow = VNC.Core.Xaml.PhysicalTree.FindAncestor<DXWindow>((DependencyObject)sender);

                var tv = VNC.Core.Xaml.PhysicalTree.FindChild<TableView>(vtrootWindow, "tableView");

                if (isVisible)
                {
                    tv.ShowGroupPanel = true;
                }
                else
                {
                    tv.ShowGroupPanel = false;
                }
            }
        }

        #endregion

        private void OnGetHelpOn(object sender, RoutedEventArgs e)
        {
            string helpTopic = ((Button)sender).Tag.ToString();

            Help.GetHelpOnTopic(helpTopic);

            //switch (helpTopic)
            //{
            //    case "DatabaseExpandTemplate":
            //        MessageBox.Show("Display Help On " + helpTopic);
            //        break;

            //    case "InstanceExpandTemplate":
            //        MessageBox.Show("Display Help On " + helpTopic);
            //        break;

            //    case "Servers":
            //        MessageBox.Show("Display Help On " + helpTopic);
            //        break;

            //    case "SnapShotControlsTemplate":
            //        MessageBox.Show("Display Help On " + helpTopic);
            //        break;

            //}

        }

        private void ShowFilterPanel(object sender, RoutedEventArgs e)
        {
            _isDisplayed_FilterPanel = !_isDisplayed_FilterPanel;

            UpdateVisibilityOfFilterPanel(sender, _isDisplayed_FilterPanel);
        }

        private void ShowGroupByPanel(object sender, RoutedEventArgs e)
        {
            _isDisplayed_GroupByPanel = !_isDisplayed_GroupByPanel;

            UpdateVisibilityOfGroupByPanel(sender, _isDisplayed_GroupByPanel);
        }

    }
}
