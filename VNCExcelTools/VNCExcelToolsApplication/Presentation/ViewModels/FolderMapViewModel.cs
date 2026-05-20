using System;
using System.Windows;
using System.Windows.Input;

using Microsoft.Office.Interop.Excel;

using Prism.Commands;
using Prism.Dialogs;
using Prism.Events;

using VNC;
using VNC.Core.Mvvm;
using VNC.Core.Presentation;
using VNC.WPF.Presentation.Views;

using VNCExcelToolsApplication.Presentation.Views;

using XlHelper = VNC.VSTOAddIn.Excel.Helper;

namespace VNCExcelToolsApplication.Presentation.ViewModels
{
    public class FolderMapViewModel : EventViewModelBase, IFolderMapViewModel, IInstanceCountVM
    {
        #region Fields (none)



        private const string _ERROR_EMPTY_CELL = "Cell is empty.  Must select a populated starting cell first.";

        #endregion

        #region Constructors, Initialization, and Load

        public FolderMapViewModel(
            IEventAggregator eventAggregator,
            IDialogService dialogService) : base(eventAggregator, dialogService)
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.Constructor) startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            InstanceCountVM++;

            // TODO(crhodes)
            // Save constructor parameters here

            InitializeViewModel();

            if (Common.VNCLogging.Constructor) Log.CONSTRUCTOR($"Exit VM:{InstanceCountVM}", Common.LOG_CATEGORY, startTicks);
        }

        private void InitializeViewModel()
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.ViewModelLow) startTicks = Log.VIEWMODEL_LOW("Enter", Common.LOG_CATEGORY);

            // NOTE(crhodes)
            // Put things here that initialize the ViewModel
            // Initialize EventHandlers, Commands, etc.

            SayHelloCommand = new DelegateCommand(
                SayHello, SayHelloCanExecute);

            Message = "FolderMapViewModel says hello";

            CreateFolderMapCommand = new DelegateCommand(CreateFolderMap, CreateFolderMapCanExecute);

            // If using CommandParameter, figure out TYPE here and below
            // and remove above declaration
            //CreateFolderMapCommand = new DelegateCommand<TYPE>(CreateFolderMap, CreateFolderMapCanExecute);

            GroupDownCommand = new DelegateCommand(GroupDown, GroupDownCanExecute);

            // If using CommandParameter, figure out TYPE here and below
            // and remove above declaration
            //GroupDownCommand = new DelegateCommand<TYPE>(GroupDown, GroupDownCanExecute);

            GroupDownAllCommand = new DelegateCommand(GroupDownAll, GroupDownAllCanExecute);

            // If using CommandParameter, figure out TYPE here and below
            // and remove above declaration
            //GroupDownAllCommand = new DelegateCommand<TYPE>(GroupDownAll, GroupDownAllCanExecute);

            UngroupSelectionCommand = new DelegateCommand(UngroupSelection, UngroupSelectionCanExecute);

            // If using CommandParameter, figure out TYPE here and below
            // and remove above declaration
            //UngroupSelectionCommand = new DelegateCommand<TYPE>(UngroupSelection, UngroupSelectionCanExecute);
            // Start Cut Two - Put this in InitializeViewModel or Constructor

            SearchLeftCommand = new DelegateCommand(SearchLeft, SearchLeftCanExecute);
            SearchRightCommand = new DelegateCommand(SearchRight, SearchRightCanExecute);
            SearchUpCommand = new DelegateCommand(SearchUp, SearchUpCanExecute);
            SearchDownCommand = new DelegateCommand(SearchDown, SearchDownCanExecute);

            // If using CommandParameter, figure out TYPE here and below
            // and remove above declaration
            //SearchLeftCommand = new DelegateCommand<TYPE>(SearchLeft, SearchLeftCanExecute);

            if (Common.VNCLogging.ViewModelLow) Log.VIEWMODEL_LOW("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region Enums (none)

        public enum SearchDirection : Int32
        {
            Left = 1,
            Up = 2,
            Right = 3,
            Down = 4,
        }

        #endregion

        #region Structures (none)



        #endregion

        #region Properties (none)



        #endregion

        #region Event Handlers (none)



        #endregion

        #region Public Methods (none)



        #endregion

        #region Commands

        #region SayHello Command

        public ICommand SayHelloCommand { get; private set; }

        private void SayHello()
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.EventHandler) startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Message = "Hello";

            if (Common.VNCLogging.EventHandler) Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private bool SayHelloCanExecute()
        {
            return true;
        }

        #endregion

        #region CreateFolderMap Command

        // If displaying UserControl
        public static WindowHost _CreateFolderMapHost = null;

        // If using CommandParameter, figure out TYPE here and above
        //public DelegateCommand<TYPE>? CreateFolderMapCommand { get; set; }
        public DelegateCommand? CreateFolderMapCommand { get; set; }

        public string CreateFolderMapContent { get; set; } = "CreateFolderMap";
        public string CreateFolderMapToolTip { get; set; } = "CreateFolderMap ToolTip";

        // Can get fancy and use Resources
        //public string CreateFolderMapContent { get; set; } = "ViewName_CreateFolderMapContent";
        //public string CreateFolderMapToolTip { get; set; } = "ViewName_CreateFolderMapContentToolTip";

        // Put these in Resource File
        //    <system:String x:Key="ViewName_CreateFolderMapContent">CreateFolderMap</system:String>
        //    <system:String x:Key="ViewName_CreateFolderMapContentToolTip">CreateFolderMap ToolTip</system:String>  

        // If using CommandParameter, figure out TYPE here
        //public void CreateFolderMap(TYPE value)
        public void CreateFolderMap()
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.EventHandler) startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);
            // TODO(crhodes)
            // Do something amazing.

            Message = "Cool, you called CreateFolderMap";

            PublishStatusMessage(Message);

            // If launching a UserControl

            if (_CreateFolderMapHost is null) _CreateFolderMapHost = new WindowHost();
            //var userControl = new CreateFolderMap();
            var userControl = (CreateFolderMap)Common.ApplicationBootstrapper.Container.Resolve(typeof(CreateFolderMap));

            _CreateFolderMapHost.DisplayUserControlInHost(
                "Create Folder Map",
                //Common.DEFAULT_WINDOW_WIDTH,
                //Common.DEFAULT_WINDOW_HEIGHT,
                (Int32)userControl.Width + Common.WINDOW_HOSTING_USER_CONTROL_WIDTH_PAD,
                (Int32)userControl.Height + Common.WINDOW_HOSTING_USER_CONTROL_HEIGHT_PAD,
                ShowWindowMode.Modal_ShowDialog,
                userControl);

            // Uncomment this if you are telling someone else to handle this

            // Common.EventAggregator.GetEvent<CreateFolderMapEvent>().Publish();

            // May want EventArgs

            //  EventAggregator.GetEvent<CreateFolderMapEvent>().Publish(
            //      new CreateFolderMapEventArgs()
            //      {
            //            Organization = _collectionMainViewModel.SelectedCollection.Organization,
            //            Process = _contextMainViewModel.Context.SelectedProcess
            //      });

            // Start Cut Four - Put this in PrismEvents

            // public class CreateFolderMapEvent : PubSubEvent { }

            // End Cut Four

            // Start Cut Five - Put this in places that listen for event

            //Common.EventAggregator.GetEvent<CreateFolderMapEvent>().Subscribe(CreateFolderMap);

            // End Cut Five

            if (Common.VNCLogging.EventHandler) Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // If using CommandParameter, figure out TYPE and fix above
        //public bool CreateFolderMapCanExecute(TYPE value)
        public bool CreateFolderMapCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion

        #region GroupDown Command

        public DelegateCommand? GroupDownCommand { get; set; }
        // If using CommandParameter, figure out TYPE here and above
        // and remove above declaration
        //public DelegateCommand<TYPE>? GroupDownCommand { get; set; }

        // If displaying UserControl
        // public static WindowHost _GroupDownHost = null;

        // If using CommandParameter, figure out TYPE here
        //public TYPE GroupDownCommandParameter;

        public string GroupDownContent { get; set; } = "GroupDown";
        public string GroupDownToolTip { get; set; } = "GroupDown ToolTip";

        // Can get fancy and use Resources
        //public string GroupDownContent { get; set; } = "ViewName_GroupDownContent";
        //public string GroupDownToolTip { get; set; } = "ViewName_GroupDownContentToolTip";

        // Put these in Resource File
        //    <system:String x:Key="ViewName_GroupDownContent">GroupDown</system:String>
        //    <system:String x:Key="ViewName_GroupDownContentToolTip">GroupDown ToolTip</system:String>  

        // If using CommandParameter, figure out TYPE here
        //public void GroupDown(TYPE value)
        public void GroupDown()
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.EventHandler) startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);
            // TODO(crhodes)
            // Do something amazing.

            Message = "Cool, you called GroupDown";

            PublishStatusMessage(Message);

            Worksheet activeSheet = (Worksheet)Common.ExcelApplication.ActiveSheet;
            Range activeCell = Common.ExcelApplication.ActiveCell;

            int currentRow;
            int currentColumn;
            int lastPopulatedRow;
            int lastPopulatedColumn;
            int endRowOfSection;

            if (activeCell.Value == null)
            {
                MessageBox.Show(_ERROR_EMPTY_CELL);
            }
            else
            {
                // Get the last populated cell on the worksheet

                lastPopulatedRow = activeCell.SpecialCells(XlCellType.xlCellTypeLastCell).Row;
                lastPopulatedColumn = activeCell.SpecialCells(XlCellType.xlCellTypeLastCell).Column;

                // Save where we currently are located
                currentRow = activeCell.Row;
                currentColumn = activeCell.Column;

                endRowOfSection = XlHelper.GetEndOfSectionDown(currentRow, currentColumn, lastPopulatedRow, currentColumn);
                ((Range)activeSheet.Rows[currentRow + 1 + ":" + endRowOfSection]).Group();

                //activeSheet.Cells[endRowOfSection, startColumn].Select();
                ((Range)activeSheet.Rows[currentRow + 1 + ":" + endRowOfSection]).Hidden = true;
            }

            // If launching a UserControl

            // if (_GroupDownHost is null) _GroupDownHost = new WindowHost();
            // var userControl = new USERCONTROL();

            // _loggingConfigurationHost.DisplayUserControlInHost(
            //     "TITLE GOES HERE",
            //     //Common.DEFAULT_WINDOW_WIDTH,
            //     //Common.DEFAULT_WINDOW_HEIGHT,
            //     (Int32)userControl.Width + Common.WINDOW_HOSTING_USER_CONTROL_WIDTH_PAD,
            //     (Int32)userControl.Height + Common.WINDOW_HOSTING_USER_CONTROL_HEIGHT_PAD,
            //     ShowWindowMode.Modeless_Show,
            //     userControl);

            // Uncomment this if you are telling someone else to handle this

            // Common.EventAggregator.GetEvent<GroupDownEvent>().Publish();

            // May want EventArgs

            //  EventAggregator.GetEvent<GroupDownEvent>().Publish(
            //      new GroupDownEventArgs()
            //      {
            //            Organization = _collectionMainViewModel.SelectedCollection.Organization,
            //            Process = _contextMainViewModel.Context.SelectedProcess
            //      });

            // Start Cut Four - Put this in PrismEvents

            // public class GroupDownEvent : PubSubEvent { }

            // End Cut Four

            // Start Cut Five - Put this in places that listen for event

            //Common.EventAggregator.GetEvent<GroupDownEvent>().Subscribe(GroupDown);

            // End Cut Five

            if (Common.VNCLogging.EventHandler) Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // If using CommandParameter, figure out TYPE and fix above
        //public bool GroupDownCanExecute(TYPE value)
        public bool GroupDownCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion

        #region GroupDownAll Command

        // If displaying UserControl
        // public static WindowHost _GroupDownAllHost = null;

        public DelegateCommand? GroupDownAllCommand { get; set; }
        // If using CommandParameter, figure out TYPE here and above
        // and remove above declaration
        //public DelegateCommand<TYPE>? GroupDownAllCommand { get; set; }

        // If using CommandParameter, figure out TYPE here
        //public TYPE GroupDownAllCommandParameter;

        public string GroupDownAllContent { get; set; } = "GroupDownAll";
        public string GroupDownAllToolTip { get; set; } = "GroupDownAll ToolTip";

        // Can get fancy and use Resources
        //public string GroupDownAllContent { get; set; } = "ViewName_GroupDownAllContent";
        //public string GroupDownAllToolTip { get; set; } = "ViewName_GroupDownAllContentToolTip";

        // Put these in Resource File
        //    <system:String x:Key="ViewName_GroupDownAllContent">GroupDownAll</system:String>
        //    <system:String x:Key="ViewName_GroupDownAllContentToolTip">GroupDownAll ToolTip</system:String>  

        // If using CommandParameter, figure out TYPE here
        //public void GroupDownAll(TYPE value)
        public void GroupDownAll()
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.EventHandler) startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);
            // TODO(crhodes)
            // Do something amazing.

            Message = "Cool, you called GroupDownAll";

            PublishStatusMessage(Message);

            Worksheet activeSheet = (Worksheet)Common.ExcelApplication.ActiveSheet;
            Range activeCell = Common.ExcelApplication.ActiveCell;

            int currentRow = activeCell.Row;
            int currentColumn = activeCell.Column;

            int lastPopulatedRow;
            int lastPopulatedColumn;
            int endRowOfSection;

            if (activeCell.Value == null)
            {
                MessageBox.Show(_ERROR_EMPTY_CELL);
            }
            else
            {
                // Get the last populated cell on the worksheet

                lastPopulatedRow = activeCell.SpecialCells(XlCellType.xlCellTypeLastCell).Row;
                lastPopulatedColumn = activeCell.SpecialCells(XlCellType.xlCellTypeLastCell).Column;

                while (currentRow < lastPopulatedRow)
                {
                    endRowOfSection = XlHelper.GetEndOfSectionDown(currentRow, currentColumn, lastPopulatedRow, currentColumn);
                    ((Range)activeSheet.Rows[currentRow + 1 + ":" + endRowOfSection]).Group();
                    ((Range)activeSheet.Rows[currentRow + 1 + ":" + endRowOfSection]).Hidden = true;

                    // Move to the next possible row to collapse.  
                    currentRow = endRowOfSection + 1;

                    // Keep going if next row does not have an empty cell or past end of data.
                    while (((Range)activeSheet.Cells[currentRow + 1, currentColumn]).Value != null)
                    {
                        if (currentRow >= lastPopulatedRow)
                        {
                            break;
                        }

                        currentRow++;
                    }
                }
            }

            // If launching a UserControl

            // if (_GroupDownAllHost is null) _GroupDownAllHost = new WindowHost();
            // var userControl = new USERCONTROL();

            // _loggingConfigurationHost.DisplayUserControlInHost(
            //     "TITLE GOES HERE",
            //     //Common.DEFAULT_WINDOW_WIDTH,
            //     //Common.DEFAULT_WINDOW_HEIGHT,
            //     (Int32)userControl.Width + Common.WINDOW_HOSTING_USER_CONTROL_WIDTH_PAD,
            //     (Int32)userControl.Height + Common.WINDOW_HOSTING_USER_CONTROL_HEIGHT_PAD,
            //     ShowWindowMode.Modeless_Show,
            //     userControl);

            // Uncomment this if you are telling someone else to handle this

            // Common.EventAggregator.GetEvent<GroupDownAllEvent>().Publish();

            // May want EventArgs

            //  EventAggregator.GetEvent<GroupDownAllEvent>().Publish(
            //      new GroupDownAllEventArgs()
            //      {
            //            Organization = _collectionMainViewModel.SelectedCollection.Organization,
            //            Process = _contextMainViewModel.Context.SelectedProcess
            //      });

            // Start Cut Four - Put this in PrismEvents

            // public class GroupDownAllEvent : PubSubEvent { }

            // End Cut Four

            // Start Cut Five - Put this in places that listen for event

            //Common.EventAggregator.GetEvent<GroupDownAllEvent>().Subscribe(GroupDownAll);

            // End Cut Five

            if (Common.VNCLogging.EventHandler) Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // If using CommandParameter, figure out TYPE and fix above
        //public bool GroupDownAllCanExecute(TYPE value)
        public bool GroupDownAllCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion

        #region UngroupSelection Command

        // If displaying UserControl
        // public static WindowHost _UngroupSelectionHost = null;

        public DelegateCommand? UngroupSelectionCommand { get; set; }
        // If using CommandParameter, figure out TYPE here and above
        // and remove above declaration
        //public DelegateCommand<TYPE>? UngroupSelectionCommand { get; set; }

        // If using CommandParameter, figure out TYPE here
        //public TYPE UngroupSelectionCommandParameter;

        public string UngroupSelectionContent { get; set; } = "UngroupSelection";
        public string UngroupSelectionToolTip { get; set; } = "UngroupSelection ToolTip";

        // Can get fancy and use Resources
        //public string UngroupSelectionContent { get; set; } = "ViewName_UngroupSelectionContent";
        //public string UngroupSelectionToolTip { get; set; } = "ViewName_UngroupSelectionContentToolTip";

        // Put these in Resource File
        //    <system:String x:Key="ViewName_UngroupSelectionContent">UngroupSelection</system:String>
        //    <system:String x:Key="ViewName_UngroupSelectionContentToolTip">UngroupSelection ToolTip</system:String>  

        // If using CommandParameter, figure out TYPE here
        //public void UngroupSelection(TYPE value)
        public void UngroupSelection()
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.EventHandler) startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);
            // TODO(crhodes)
            // Do something amazing.

            Message = "Cool, you called UngroupSelection";

            PublishStatusMessage(Message);

            // If launching a UserControl

            // if (_UngroupSelectionHost is null) _UngroupSelectionHost = new WindowHost();
            // var userControl = new USERCONTROL();

            // _loggingConfigurationHost.DisplayUserControlInHost(
            //     "TITLE GOES HERE",
            //     //Common.DEFAULT_WINDOW_WIDTH,
            //     //Common.DEFAULT_WINDOW_HEIGHT,
            //     (Int32)userControl.Width + Common.WINDOW_HOSTING_USER_CONTROL_WIDTH_PAD,
            //     (Int32)userControl.Height + Common.WINDOW_HOSTING_USER_CONTROL_HEIGHT_PAD,
            //     ShowWindowMode.Modeless_Show,
            //     userControl);

            // Uncomment this if you are telling someone else to handle this

            // Common.EventAggregator.GetEvent<UngroupSelectionEvent>().Publish();

            // May want EventArgs

            //  EventAggregator.GetEvent<UngroupSelectionEvent>().Publish(
            //      new UngroupSelectionEventArgs()
            //      {
            //            Organization = _collectionMainViewModel.SelectedCollection.Organization,
            //            Process = _contextMainViewModel.Context.SelectedProcess
            //      });

            // Start Cut Four - Put this in PrismEvents

            // public class UngroupSelectionEvent : PubSubEvent { }

            // End Cut Four

            // Start Cut Five - Put this in places that listen for event

            //Common.EventAggregator.GetEvent<UngroupSelectionEvent>().Subscribe(UngroupSelection);

            // End Cut Five

            if (Common.VNCLogging.EventHandler) Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // If using CommandParameter, figure out TYPE and fix above
        //public bool UngroupSelectionCanExecute(TYPE value)
        public bool UngroupSelectionCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion

        #region SearchLeft Command

        // If displaying UserControl
        // public static WindowHost _SearchLeftHost = null;

        public DelegateCommand? SearchLeftCommand { get; set; }
        // If using CommandParameter, figure out TYPE here and above
        // and remove above declaration
        //public DelegateCommand<TYPE>? SearchLeftCommand { get; set; }

        // If using CommandParameter, figure out TYPE here
        //public TYPE SearchLeftCommandParameter;

        public string SearchLeftContent { get; set; } = "SearchLeft";
        public string SearchLeftToolTip { get; set; } = "SearchLeft ToolTip";

        // Can get fancy and use Resources
        //public string SearchLeftContent { get; set; } = "ViewName_SearchLeftContent";
        //public string SearchLeftToolTip { get; set; } = "ViewName_SearchLeftContentToolTip";

        // Put these in Resource File
        //    <system:String x:Key="ViewName_SearchLeftContent">SearchLeft</system:String>
        //    <system:String x:Key="ViewName_SearchLeftContentToolTip">SearchLeft ToolTip</system:String>  

        // If using CommandParameter, figure out TYPE here
        //public void SearchLeft(TYPE value)
        public void SearchLeft()
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.EventHandler) startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);
            // TODO(crhodes)
            // Do something amazing.

            Message = "Cool, you called SearchLeft";

            PublishStatusMessage(Message);

            Worksheet activeSheet = (Worksheet)Common.ExcelApplication.ActiveSheet;
            Range activeCell = Common.ExcelApplication.ActiveCell;

            int matchRow = activeCell.Row;
            int matchColumn = XlHelper.FindPrevious_PopulatedColumn_InRow(activeCell);

            ((Range)activeSheet.Cells[matchRow, matchColumn]).Select();

            // If launching a UserControl

            // if (_SearchLeftHost is null) _SearchLeftHost = new WindowHost();
            // var userControl = new USERCONTROL();

            // _loggingConfigurationHost.DisplayUserControlInHost(
            //     "TITLE GOES HERE",
            //     //Common.DEFAULT_WINDOW_WIDTH,
            //     //Common.DEFAULT_WINDOW_HEIGHT,
            //     (Int32)userControl.Width + Common.WINDOW_HOSTING_USER_CONTROL_WIDTH_PAD,
            //     (Int32)userControl.Height + Common.WINDOW_HOSTING_USER_CONTROL_HEIGHT_PAD,
            //     ShowWindowMode.Modeless_Show,
            //     userControl);

            // Uncomment this if you are telling someone else to handle this

            // Common.EventAggregator.GetEvent<SearchLeftEvent>().Publish();

            // May want EventArgs

            //  EventAggregator.GetEvent<SearchLeftEvent>().Publish(
            //      new SearchLeftEventArgs()
            //      {
            //            Organization = _collectionMainViewModel.SelectedCollection.Organization,
            //            Process = _contextMainViewModel.Context.SelectedProcess
            //      });

            // Start Cut Four - Put this in PrismEvents

            // public class SearchLeftEvent : PubSubEvent { }

            // End Cut Four

            // Start Cut Five - Put this in places that listen for event

            //Common.EventAggregator.GetEvent<SearchLeftEvent>().Subscribe(SearchLeft);

            // End Cut Five

            if (Common.VNCLogging.EventHandler) Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // If using CommandParameter, figure out TYPE and fix above
        //public bool SearchLeftCanExecute(TYPE value)
        public bool SearchLeftCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion

        #region SearchRight Command

        // If displaying UserControl
        // public static WindowHost _SearchRightHost = null;

        public DelegateCommand? SearchRightCommand { get; set; }
        // If using CommandParameter, figure out TYPE here and above
        // and remove above declaration
        //public DelegateCommand<TYPE>? SearchRightCommand { get; set; }

        // If using CommandParameter, figure out TYPE here
        //public TYPE SearchRightCommandParameter;

        public string SearchRightContent { get; set; } = "SearchRight";
        public string SearchRightToolTip { get; set; } = "SearchRight ToolTip";

        // Can get fancy and use Resources
        //public string SearchRightContent { get; set; } = "ViewName_SearchRightContent";
        //public string SearchRightToolTip { get; set; } = "ViewName_SearchRightContentToolTip";

        // Put these in Resource File
        //    <system:String x:Key="ViewName_SearchRightContent">SearchRight</system:String>
        //    <system:String x:Key="ViewName_SearchRightContentToolTip">SearchRight ToolTip</system:String>  

        // If using CommandParameter, figure out TYPE here
        //public void SearchRight(TYPE value)
        public void SearchRight()
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.EventHandler) startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);
            // TODO(crhodes)
            // Do something amazing.

            Message = "Cool, you called SearchRight";

            PublishStatusMessage(Message);

            Worksheet activeSheet = (Worksheet)Common.ExcelApplication.ActiveSheet;
            Range activeCell = Common.ExcelApplication.ActiveCell;

            int matchRow = activeCell.Row;
            int matchColumn = XlHelper.FindNext_PopulatedColumn_InRow(activeCell);

            ((Range)activeSheet.Cells[matchRow, matchColumn]).Select();

            // If launching a UserControl

            // if (_SearchRightHost is null) _SearchRightHost = new WindowHost();
            // var userControl = new USERCONTROL();

            // _loggingConfigurationHost.DisplayUserControlInHost(
            //     "TITLE GOES HERE",
            //     //Common.DEFAULT_WINDOW_WIDTH,
            //     //Common.DEFAULT_WINDOW_HEIGHT,
            //     (Int32)userControl.Width + Common.WINDOW_HOSTING_USER_CONTROL_WIDTH_PAD,
            //     (Int32)userControl.Height + Common.WINDOW_HOSTING_USER_CONTROL_HEIGHT_PAD,
            //     ShowWindowMode.Modeless_Show,
            //     userControl);

            // Uncomment this if you are telling someone else to handle this

            // Common.EventAggregator.GetEvent<SearchRightEvent>().Publish();

            // May want EventArgs

            //  EventAggregator.GetEvent<SearchRightEvent>().Publish(
            //      new SearchRightEventArgs()
            //      {
            //            Organization = _collectionMainViewModel.SelectedCollection.Organization,
            //            Process = _contextMainViewModel.Context.SelectedProcess
            //      });

            // Start Cut Four - Put this in PrismEvents

            // public class SearchRightEvent : PubSubEvent { }

            // End Cut Four

            // Start Cut Five - Put this in places that listen for event

            //Common.EventAggregator.GetEvent<SearchRightEvent>().Subscribe(SearchRight);

            // End Cut Five

            if (Common.VNCLogging.EventHandler) Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // If using CommandParameter, figure out TYPE and fix above
        //public bool SearchRightCanExecute(TYPE value)
        public bool SearchRightCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion

        #region SearchUp Command

        // If displaying UserControl
        // public static WindowHost _SearchUpHost = null;

        public DelegateCommand? SearchUpCommand { get; set; }
        // If using CommandParameter, figure out TYPE here and above
        // and remove above declaration
        //public DelegateCommand<TYPE>? SearchUpCommand { get; set; }

        // If using CommandParameter, figure out TYPE here
        //public TYPE SearchUpCommandParameter;

        public string SearchUpContent { get; set; } = "SearchUp";
        public string SearchUpToolTip { get; set; } = "SearchUp ToolTip";

        // Can get fancy and use Resources
        //public string SearchUpContent { get; set; } = "ViewName_SearchUpContent";
        //public string SearchUpToolTip { get; set; } = "ViewName_SearchUpContentToolTip";

        // Put these in Resource File
        //    <system:String x:Key="ViewName_SearchUpContent">SearchUp</system:String>
        //    <system:String x:Key="ViewName_SearchUpContentToolTip">SearchUp ToolTip</system:String>  

        // If using CommandParameter, figure out TYPE here
        //public void SearchUp(TYPE value)
        public void SearchUp()
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.EventHandler) startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);
            // TODO(crhodes)
            // Do something amazing.

            Message = "Cool, you called SearchUp";

            PublishStatusMessage(Message);

            Worksheet activeSheet = (Worksheet)Common.ExcelApplication.ActiveSheet;
            Range activeCell = Common.ExcelApplication.ActiveCell;

            int matchRow = XlHelper.FindPrevious_PopulatedRow_InColumn(activeCell);
            int matchColumn = activeCell.Column;

            ((Range)activeSheet.Cells[matchRow, matchColumn]).Select();

            // If launching a UserControl

            // if (_SearchUpHost is null) _SearchUpHost = new WindowHost();
            // var userControl = new USERCONTROL();

            // _loggingConfigurationHost.DisplayUserControlInHost(
            //     "TITLE GOES HERE",
            //     //Common.DEFAULT_WINDOW_WIDTH,
            //     //Common.DEFAULT_WINDOW_HEIGHT,
            //     (Int32)userControl.Width + Common.WINDOW_HOSTING_USER_CONTROL_WIDTH_PAD,
            //     (Int32)userControl.Height + Common.WINDOW_HOSTING_USER_CONTROL_HEIGHT_PAD,
            //     ShowWindowMode.Modeless_Show,
            //     userControl);

            // Uncomment this if you are telling someone else to handle this

            // Common.EventAggregator.GetEvent<SearchUpEvent>().Publish();

            // May want EventArgs

            //  EventAggregator.GetEvent<SearchUpEvent>().Publish(
            //      new SearchUpEventArgs()
            //      {
            //            Organization = _collectionMainViewModel.SelectedCollection.Organization,
            //            Process = _contextMainViewModel.Context.SelectedProcess
            //      });

            // Start Cut Four - Put this in PrismEvents

            // public class SearchUpEvent : PubSubEvent { }

            // End Cut Four

            // Start Cut Five - Put this in places that listen for event

            //Common.EventAggregator.GetEvent<SearchUpEvent>().Subscribe(SearchUp);

            // End Cut Five

            if (Common.VNCLogging.EventHandler) Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // If using CommandParameter, figure out TYPE and fix above
        //public bool SearchUpCanExecute(TYPE value)
        public bool SearchUpCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion

        #region SearchDown Command

        // If displaying UserControl
        // public static WindowHost _SearchDownHost = null;

        public DelegateCommand? SearchDownCommand { get; set; }
        // If using CommandParameter, figure out TYPE here and above
        // and remove above declaration
        //public DelegateCommand<TYPE>? SearchDownCommand { get; set; }

        // If using CommandParameter, figure out TYPE here
        //public TYPE SearchDownCommandParameter;

        public string SearchDownContent { get; set; } = "SearchDown";
        public string SearchDownToolTip { get; set; } = "SearchDown ToolTip";

        // Can get fancy and use Resources
        //public string SearchDownContent { get; set; } = "ViewName_SearchDownContent";
        //public string SearchDownToolTip { get; set; } = "ViewName_SearchDownContentToolTip";

        // Put these in Resource File
        //    <system:String x:Key="ViewName_SearchDownContent">SearchDown</system:String>
        //    <system:String x:Key="ViewName_SearchDownContentToolTip">SearchDown ToolTip</system:String>  

        // If using CommandParameter, figure out TYPE here
        //public void SearchDown(TYPE value)
        public void SearchDown()
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.EventHandler) startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);
            // TODO(crhodes)
            // Do something amazing.

            Message = "Cool, you called SearchDown";

            PublishStatusMessage(Message);

            Worksheet activeSheet = (Worksheet)Common.ExcelApplication.ActiveSheet;
            Range activeCell = Common.ExcelApplication.ActiveCell;

            int matchRow = XlHelper.FindNext_PopulatedRow_InColumn(activeCell);
            int matchColumn = activeCell.Column;

            ((Range)activeSheet.Cells[matchRow, matchColumn]).Select();

            // If launching a UserControl

            // if (_SearchDownHost is null) _SearchDownHost = new WindowHost();
            // var userControl = new USERCONTROL();

            // _loggingConfigurationHost.DisplayUserControlInHost(
            //     "TITLE GOES HERE",
            //     //Common.DEFAULT_WINDOW_WIDTH,
            //     //Common.DEFAULT_WINDOW_HEIGHT,
            //     (Int32)userControl.Width + Common.WINDOW_HOSTING_USER_CONTROL_WIDTH_PAD,
            //     (Int32)userControl.Height + Common.WINDOW_HOSTING_USER_CONTROL_HEIGHT_PAD,
            //     ShowWindowMode.Modeless_Show,
            //     userControl);

            // Uncomment this if you are telling someone else to handle this

            // Common.EventAggregator.GetEvent<SearchDownEvent>().Publish();

            // May want EventArgs

            //  EventAggregator.GetEvent<SearchDownEvent>().Publish(
            //      new SearchDownEventArgs()
            //      {
            //            Organization = _collectionMainViewModel.SelectedCollection.Organization,
            //            Process = _contextMainViewModel.Context.SelectedProcess
            //      });

            // Start Cut Four - Put this in PrismEvents

            // public class SearchDownEvent : PubSubEvent { }

            // End Cut Four

            // Start Cut Five - Put this in places that listen for event

            //Common.EventAggregator.GetEvent<SearchDownEvent>().Subscribe(SearchDown);

            // End Cut Five

            if (Common.VNCLogging.EventHandler) Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // If using CommandParameter, figure out TYPE and fix above
        //public bool SearchDownCanExecute(TYPE value)
        public bool SearchDownCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion

        #endregion

        #region Protected Methods (none)



        #endregion

        #region Private Methods (none)

        private void SearchForPopulatedCell(SearchDirection searchDirection)
        {
            Worksheet activeSheet = (Worksheet)Common.ExcelApplication.ActiveSheet;
            Range activeCell = (Range)Common.ExcelApplication.ActiveCell;

            Int32 matchRow = 0;
            Int32 matchColumn = 0;

            switch (searchDirection)
            {
                case SearchDirection.Left:
                    matchRow = activeCell.Row;
                    matchColumn = XlHelper.FindPrevious_PopulatedColumn_InRow(activeCell);
                    break;

                case SearchDirection.Up:
                    matchRow = XlHelper.FindPrevious_PopulatedRow_InColumn(activeCell);
                    matchColumn = activeCell.Column;
                    break;

                case SearchDirection.Right:
                    matchRow = activeCell.Row;
                    matchColumn = XlHelper.FindNext_PopulatedColumn_InRow(activeCell);
                    break;

                case SearchDirection.Down:
                    matchRow = XlHelper.FindNext_PopulatedRow_InColumn(activeCell);
                    matchColumn = activeCell.Column;
                    break;
            }


            ((Range)activeSheet.Cells[matchRow, matchColumn]).Select();
        }

        #endregion

        #region IInstanceCountVM

        private static int _instanceCountVM;

        public int InstanceCountVM
        {
            get => _instanceCountVM;
            set => _instanceCountVM = value;
        }

        #endregion
    }
}
