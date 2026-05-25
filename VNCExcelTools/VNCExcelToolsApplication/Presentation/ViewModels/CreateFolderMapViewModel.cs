using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Windows.Input;

using System.Windows.Media;

using DevExpress.CodeParser;

using Microsoft.Office.Interop.Excel;

using Prism.Commands;
using Prism.Dialogs;
using Prism.Events;

using VNC;
using VNC.Core.Mvvm;

namespace VNCExcelToolsApplication.Presentation.ViewModels
{
    public class CreateFolderMapViewModel : EventViewModelBase, ICreateFolderMapViewModel, IInstanceCountVM
    {
        // TODO(crhodes)
        // Decide if need this or just define in Properties
        public class RegEx
        {
            // SharePoint Folder/File/Document Libraries may not contain any of the following characters
            //   / \ : * ? " < > | <TAB> { } % ~ &
            // nor may they end in periods or contain embedded double periods.
            // The following regular expressions capture these rules.  

            internal const string cIllegalFileCharacters = "[/\\\\:\\*\\?\"<>\\|#\\{}%~&]|\\.\\.";

            internal const string cIllegalFolderCharacters = "[:\\*\\?\"<>\\|#\\{}%~&]";
        }

        #region Fields and Constants

        private const int _INDENT_LEVEL = 1;
        private const int _COL_WIDTH = 3;
        private const int _NOTE_WIDTH = 20;
        private const int _FILE_FONT_SIZE = 6;

        private const int _FOLDER_FONT_SIZE = 8;
        private const int _HEADING_ROW = 2;

        private const int _INITIAL_ROW = _HEADING_ROW + 1;

        private const string _ERROR_EMPTY_CELL = "Cell is empty.  Must select a populated starting cell first.";
        // Folder level info starts here
        private const int _FOLDER_INFO_COL = 1;
        private const int _FOLDER_INFO_LEN = 10;
        // File Info starts here
        private const int _FILE_INFO_COL = 5;
        private const int _FILE_INFO_LEN = 4;
        private const int _NOTE_COL = 10;
        // Map Info starts here
        private const int _INITIAL_COL = 11;

        private const bool _MAKE_BOLD = true;


        FolderBrowserDialog _folderBrowserDialog = new FolderBrowserDialog();

        private int _Row;

        private int _Column;

        private int _TotalFolders;
        private int _TotalFiles;
        private long _TotalSize;

        #endregion

        #region Constructors, Initialization, and Load

        public CreateFolderMapViewModel(
            IEventAggregator eventAggregator,
            IDialogService dialogService) : base(eventAggregator, dialogService)
        {
            Int64 startTicks = 0;
            if (Common.VNCLogLevel.Constructor) startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            InstanceCountVM++;

            // TODO(crhodes)
            // Save constructor parameters here

            InitializeViewModel();

            if (Common.VNCLogLevel.Constructor) Log.CONSTRUCTOR($"Exit VM:{InstanceCountVM}", Common.LOG_CATEGORY, startTicks);
        }

        private void InitializeViewModel()
        {
            Int64 startTicks = 0;
            if (Common.VNCLogLevel.ViewModelLow) startTicks = Log.VIEWMODEL_LOW("Enter", Common.LOG_CATEGORY);

            // NOTE(crhodes)
            // Put things here that initialize the ViewModel
            // Initialize EventHandlers, Commands, etc.

            SayHelloCommand = new DelegateCommand(
                SayHello, SayHelloCanExecute);

            CreateFolderMapCommand = new DelegateCommand(CreateFolderMap, CreateFolderMapCanExecute);

            // If using CommandParameter, figure out TYPE here and below
            // and remove above declaration
            //CreateFolderMapCommand = new DelegateCommand<TYPE>(CreateFolderMap, CreateFolderMapCanExecute);

            // Start Cut Two - Put this in InitializeViewModel or Constructor

            PickStartingFolderCommand = new DelegateCommand(PickStartingFolder, PickStartingFolderCanExecute);

            // If using CommandParameter, figure out TYPE here and below
            // and remove above declaration
            //PickStartingFolderCommand = new DelegateCommand<TYPE>(PickStartingFolder, PickStartingFolderCanExecute);

            // End Cut Two

            Message = "CreateFolderMapViewModel says hello";

            if (Common.VNCLogLevel.ViewModelLow) Log.VIEWMODEL_LOW("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region Enums

        public enum _DateType : int
        {
            LastCreate = 1,
            LastWrite = 2,
            LastAccess = 3
        }

        #endregion

        #region Structures (none)



        #endregion

        #region Properties

        #region Folder Selection

        private string _startingFolder = "";
        public string StartingFolder
        {
            get => _startingFolder;
            set
            {
                if (_startingFolder == value)
                    return;
                _startingFolder = value;
                OnPropertyChanged();
            }
        }

        public string StartingFolderToolTip { get; set; } = "Starting Folder ToolTip";

        #endregion

        #region Folder Map Content

        private Boolean _showFolders = true;
        public Boolean ShowFolders
        {
            get => _showFolders;
            set
            {
                if (_showFolders == value)
                    return;
                _showFolders = value;
                OnPropertyChanged();
            }
        }

        private Boolean _skipEmptyFolders = false;
        public Boolean SkipEmptyFolders
        {
            get => _skipEmptyFolders;
            set
            {
                if (_skipEmptyFolders == value)
                    return;
                _skipEmptyFolders = value;
                OnPropertyChanged();
            }
        }

        private Boolean _limitLevels = false;
        public Boolean LimitLevels
        {
            get => _limitLevels;
            set
            {
                if (_limitLevels == value)
                    return;
                _limitLevels = value;
                OnPropertyChanged();
            }
        }

        private Boolean _showFiles = false;
        public Boolean ShowFiles
        {
            get => _showFiles;
            set
            {
                if (_showFiles == value)
                    return;
                _showFiles = value;
                OnPropertyChanged();
            }
        }

        //public Int32 LimitLevels { get; set; }
        private Int32 _folderLimitLevel = 1;
        public Int32 FolderLimitLevel
        {
            get => _folderLimitLevel;
            set
            {
                if (_folderLimitLevel == value) return;
                _folderLimitLevel = value;
                OnPropertyChanged();
            }
        }

        #endregion

        #region Inclusion Matching Rules

        private Boolean _folderMatchUsingRegEx = false;
        public Boolean FolderMatchUsingRegEx
        {
            get => _folderMatchUsingRegEx;
            set
            {
                if (_folderMatchUsingRegEx == value)
                    return;
                _folderMatchUsingRegEx = value;
                OnPropertyChanged();
            }
        }

        //public string FolderRegExPattern { get; set; }
        private string _FolderRegExPattern = "";
        public string FolderRegExPattern
        {
            get => _FolderRegExPattern;
            set
            {
                if (_FolderRegExPattern == value) return;
                _FolderRegExPattern = value;
                OnPropertyChanged();
            }
        }

        public string FolderRegExPatternToolTip { get; set; } = "Folder RegEx Pattern ToolTip";

        private Boolean _fileMatchUsingRegEx = false;
        public Boolean FileMatchUsingRegEx
        {
            get => _fileMatchUsingRegEx;
            set
            {
                if (_fileMatchUsingRegEx == value)
                    return;
                _fileMatchUsingRegEx = value;
                OnPropertyChanged();
            }
        }

        //public string FileRegExPattern { get; set; }
        private string _FileRegExPattern = "";
        public string FileRegExPattern
        {
            get => _FileRegExPattern;
            set
            {
                if (_FileRegExPattern == value) return;
                _FileRegExPattern = value;
                OnPropertyChanged();
            }
        }

        public string FileRegExPatternToolTip { get; set; } = "File RegEx Pattern ToolTip";

        // NOTE(crhodes)
        // Excel uses System.Drawing.Color, but we want to use System.Windows.Media.Color for WPF binding.
        // So store as ARGB int for use in Excel font.Color.

        private Int32 FolderHighlightColorARGB;
        private Color _folderHighlightColor;
        public Color FolderHighlightColor
        {
            get => _folderHighlightColor;
            set
            {
                if (_folderHighlightColor == value)
                    return;
                _folderHighlightColor = value;
                FolderHighlightColorARGB = value.A << 24 | value.R << 16 | value.G << 8 | value.B;
                OnPropertyChanged();
            }
        }

        private Int32 FileHighlightColorARGB;
        private Color _fileHighlightColor;
        public Color FileHighlightColor
        {
            get => _fileHighlightColor;
            set
            {
                if (_fileHighlightColor == value)
                    return;
                _fileHighlightColor = value;
                FileHighlightColorARGB = value.A << 24 | value.R << 16 | value.G << 8 | value.B;
                OnPropertyChanged();
            }
        }

        #endregion

        #region Fonts

        private Int32 _defaultFontSize = 10;
        public Int32 DefaultFontSize
        {
            get => _defaultFontSize;
            set
            {
                if (_defaultFontSize == value) return;
                _defaultFontSize = value;
                OnPropertyChanged();
            }
        }
        private Int32 _folderFontSize = 10;
        public Int32 FolderFontSize
        {
            get => _folderFontSize;
            set
            {
                if (_folderFontSize == value) return;
                _folderFontSize = value;
                OnPropertyChanged();
            }
        }

        private Int32 _fileFontSize = 10;
        public Int32 FileFontSize
        {
            get => _fileFontSize;
            set
            {
                if (_fileFontSize == value) return;
                _fileFontSize = value;
                OnPropertyChanged();
            }
        }

        #endregion

        #region Date Information

        private Boolean _colorCodeDates = false;
        public Boolean ColorCodeDates
        {
            get => _colorCodeDates;
            set
            {
                if (_colorCodeDates == value)
                    return;
                _colorCodeDates = value;
                OnPropertyChanged();
            }
        }

        private Int32 DefaultColorARGB;
        private Color _defaultColor = Colors.Black;
        public Color DefaultColor
        {
            get => _defaultColor;
            set
            {
                if (_defaultColor == value)
                    return;
                _defaultColor = value;
                DefaultColorARGB = value.A << 24 | value.R << 16 | value.G << 8 | value.B;
                OnPropertyChanged();
            }
        }

        private Int32 _monthsSinceCreated = 24;
        public Int32 MonthsSinceCreated
        {
            get => _monthsSinceCreated;
            set
            {
                if (_monthsSinceCreated == value) return;
                _monthsSinceCreated = value;
                OnPropertyChanged();
            }
        }

        private Int32 CreatedColorARGB;
        private Color _createdColor = Colors.Red;
        public Color CreatedColor
        {
            get => _createdColor;
            set
            {
                if (_createdColor == value)
                    return;
                _createdColor = value;
                CreatedColorARGB = value.A << 24 | value.R << 16 | value.G << 8 | value.B;
                OnPropertyChanged();
            }
        }

        private Int32 _monthsSinceWritten = 24;
        public Int32 MonthsSinceWritten
        {
            get => _monthsSinceWritten;
            set
            {
                if (_monthsSinceWritten == value) return;
                _monthsSinceWritten = value;
                OnPropertyChanged();
            }
        }

        private Int32 WrittenColorARGB;
        private Color _writtenColor = Colors.Green;
        public Color WrittenColor
        {
            get => _writtenColor;
            set
            {
                if (_writtenColor == value)
                    return;
                _writtenColor = value;
                WrittenColorARGB = value.A << 24 | value.R << 16 | value.G << 8 | value.B;
                OnPropertyChanged();
            }
        }

        private Int32 _monthsSinceAccessed = 24;
        public Int32 MonthsSinceAccessed
        {
            get => _monthsSinceAccessed;
            set
            {
                if (_monthsSinceAccessed == value) return;
                _monthsSinceAccessed = value;
                OnPropertyChanged();
            }
        }

        private Int32 AccessedColorARGB;
        private Color _accessedColor = Colors.Blue;
        public Color AccessedColor
        {
            get => _accessedColor;
            set
            {
                if (_accessedColor == value)
                    return;
                _accessedColor = value;
                AccessedColorARGB = value.A << 24 | value.R << 16 | value.G << 8 | value.B;
                OnPropertyChanged();
            }
        }

        #endregion

        #region SharePoint Information

        private Boolean _checkForIllegalCharacters = false;
        public Boolean CheckForIllegalCharacters
        {
            get => _checkForIllegalCharacters;
            set
            {
                if (_checkForIllegalCharacters == value)
                    return;
                _checkForIllegalCharacters = value;
                OnPropertyChanged();
            }
        }

        private Int32 IllegalCharactersColorARGB;
        private Color _illegalCharactersColor = Colors.Orange;
        public Color IllegalCharactersColor
        {
            get => _illegalCharactersColor;
            set
            {
                if (_illegalCharactersColor == value)
                    return;
                _illegalCharactersColor = value;
                IllegalCharactersColorARGB = value.A << 24 | value.R << 16 | value.G << 8 | value.B;
                OnPropertyChanged();
            }
        }

        private string _illegalFolderCharacters = "[:\\*\\?\"<>\\|#\\{}%~&]";
        public string IllegalFolderCharacters
        {
            get => _illegalFolderCharacters;
            set
            {
                if (_illegalFolderCharacters == value) return;
                _illegalFolderCharacters = value;
                OnPropertyChanged();
            }
        }

        public string IllegalFolderCharactersToolTip { get; set; } = "Illegal Folder Characters ToolTip";

        private string _illegalFileCharacters = "[/\\\\:\\*\\?\"<>\\|#\\{}%~&]|\\.\\.";
        public string IllegalFileCharacters
        {
            get => _illegalFileCharacters;
            set
            {
                if (_illegalFileCharacters == value) return;
                _illegalFileCharacters = value;
                OnPropertyChanged();
            }
        }

        public string IllegalFileCharactersToolTip { get; set; } = "Illegal File Characters ToolTip";

        private Int32 IllegalFileNameLengthColorARGB;
        private Color _illegalFileNameLengthColor = Colors.Red;
        public Color IllegalFileNameLengthColor
        {
            get => _illegalFileNameLengthColor;
            set
            {
                if (_illegalFileNameLengthColor == value)
                    return;
                _illegalFileNameLengthColor = value;
                IllegalFileNameLengthColorARGB = value.A << 24 | value.R << 16 | value.G << 8 | value.B;
                OnPropertyChanged();
            }
        }

        private Boolean _checkFileNameLength = false;
        public Boolean CheckFileNameLength
        {
            get => _checkFileNameLength;
            set
            {
                if (_checkFileNameLength == value)
                    return;
                _checkFileNameLength = value;
                OnPropertyChanged();
            }
        }

        private int _maxFileNameLength = 250;
        public int MaxFileNameLength
        {
            get => _maxFileNameLength;
            set
            {
                if (_maxFileNameLength == value)
                    return;
                _maxFileNameLength = value;
                OnPropertyChanged();
            }
        }

        #endregion

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
            if (Common.VNCLogLevel.EventHandler) startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Message = "Hello";

            if (Common.VNCLogLevel.EventHandler) Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private bool SayHelloCanExecute()
        {
            return true;
        }

        #endregion

        #region PickStartingFolder Command

        // If displaying UserControl
        // public static WindowHost _PickStartingFolderHost = null;

        //public DelegateCommand<TYPE>? PickStartingFolderCommand { get; set; }
        // If using CommandParameter, figure out TYPE here and above
        public DelegateCommand? PickStartingFolderCommand { get; set; }

        public string PickStartingFolderContent { get; set; } = "Pick Starting Folder";
        public string PickStartingFolderToolTip { get; set; } = "Pick Starting Folder ToolTip";

        // Can get fancy and use Resources
        //public string PickStartingFolderContent { get; set; } = "ViewName_PickStartingFolderContent";
        //public string PickStartingFolderToolTip { get; set; } = "ViewName_PickStartingFolderContentToolTip";

        // Put these in Resource File
        //    <system:String x:Key="ViewName_PickStartingFolderContent">PickStartingFolder</system:String>
        //    <system:String x:Key="ViewName_PickStartingFolderContentToolTip">PickStartingFolder ToolTip</system:String>  

        // If using CommandParameter, figure out TYPE here
        //public void PickStartingFolder(TYPE value)
        public void PickStartingFolder()
        {
            Int64 startTicks = 0;
            if (Common.VNCLogLevel.EventHandler) startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);
            // TODO(crhodes)
            // Do something amazing.

            Message = "Cool, you called PickStartingFolder";

            PublishStatusMessage(Message);

            _folderBrowserDialog.ShowNewFolderButton = false;

            if (_startingFolder?.Length > 0)
            {
                _folderBrowserDialog.SelectedPath = _startingFolder;
            }

            if (_folderBrowserDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                StartingFolder = _folderBrowserDialog.SelectedPath;
            }

            // If launching a UserControl

            // if (_PickStartingFolderHost is null) _PickStartingFolderHost = new WindowHost();
            // var userControl = new USERCONTROL();

            // _PickStartingFolderHost.DisplayUserControlInHost(
            //     "TITLE GOES HERE",
            //     //Common.DEFAULT_WINDOW_WIDTH,
            //     //Common.DEFAULT_WINDOW_HEIGHT,
            //     (Int32)userControl.Width + Common.WINDOW_HOSTING_USER_CONTROL_WIDTH_PAD,
            //     (Int32)userControl.Height + Common.WINDOW_HOSTING_USER_CONTROL_HEIGHT_PAD,
            //     ShowWindowMode.Modeless_Show,
            //     userControl);

            // Uncomment this if you are telling someone else to handle this

            // Common.EventAggregator.GetEvent<PickStartingFolderEvent>().Publish();

            // May want EventArgs

            //  EventAggregator.GetEvent<PickStartingFolderEvent>().Publish(
            //      new PickStartingFolderEventArgs()
            //      {
            //            Organization = _collectionMainViewModel.SelectedCollection.Organization,
            //            Process = _contextMainViewModel.Context.SelectedProcess
            //      });

            // Start Cut Four - Put this in PrismEvents

            // public class PickStartingFolderEvent : PubSubEvent { }

            // End Cut Four

            // Start Cut Five - Put this in places that listen for event

            //Common.EventAggregator.GetEvent<PickStartingFolderEvent>().Subscribe(PickStartingFolder);

            // End Cut Five

            if (Common.VNCLogLevel.EventHandler) Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // If using CommandParameter, figure out TYPE and fix above
        //public bool PickStartingFolderCanExecute(TYPE value)
        public bool PickStartingFolderCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion

        #region CreateFolderMap Command

        // If displaying UserControl
        // public static WindowHost _CreateFolderMapHost = null;

        public DelegateCommand? CreateFolderMapCommand { get; set; }
        // If using CommandParameter, figure out TYPE here and above
        // and remove above declaration
        //public DelegateCommand<TYPE>? CreateFolderMapCommand { get; set; }

        // If using CommandParameter, figure out TYPE here
        //public TYPE CreateFolderMapCommandParameter;

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
            if (Common.VNCLogLevel.EventHandler) startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);
            // TODO(crhodes)
            // Do something amazing.

            Message = "Cool, you called CreateFolderMap";

            PublishStatusMessage(Message);

            PopulateFolderMap(StartingFolder);
            SaveFolderMap(StartingFolder);

            // If launching a UserControl

            // if (_CreateFolderMapHost is null) _CreateFolderMapHost = new WindowHost();
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

            if (Common.VNCLogLevel.EventHandler) Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
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

        #endregion

        #region Protected Methods (none)



        #endregion

        #region Private Methods (none)

        public void PopulateFolderMap(string startingFolder)
        {
            Workbook workbook = Common.ExcelApplication.ActiveWorkbook;
            Worksheet workSheet;

            try
            {
                // Set starting point for folder output

                _Row = _INITIAL_ROW;
                _Column = _INITIAL_COL;
                _TotalFolders = 0;
                _TotalFiles = 0;
                _TotalSize = 0;

                // Always add a new sheet so can accumulate results in Workbook.
                // Need to handle case when no Workbook exists.  Tried a variety of 
                // ways to determine empty workbook.  PERSONAL.XLS may be open if 
                // macros have been created so cannot rely on Workbooks.count.
                // Worksheets.Count throws an exception.
                //
                //Dim wb as Microsoft.Office.Interop.Excel.Workbook
                //
                //For Each wb In .Workbooks
                //    Debug.Print(wb.Name)
                //Next

                //Dim ws As Microsoft.Office.Interop.Excel.Worksheet

                //For Each ws In .Worksheets
                //    Debug.Print(ws.Name)
                //Next

                // ActiveWorkbook seems to work reliably.  Maybe Interop issue, who knows.

                if (workbook == null)
                {
                    workbook = Common.ExcelApplication.Workbooks.Add();
                    // Get a new WorkSheet (or more :)) for free.
                    workSheet = (Worksheet)workbook.ActiveSheet;
                }
                else
                {
                    // TODO: Prompt to use existing sheet if found.
                    workSheet = (Worksheet)workbook.Worksheets.Add();
                }

                ((Range)workSheet.Cells[_HEADING_ROW, _FOLDER_INFO_COL]).Value = "Cumulative Folder Count";
                ((Range)workSheet.Cells[_HEADING_ROW, _FOLDER_INFO_COL + 1]).Value = "Cummulative File Count";
                ((Range)workSheet.Cells[_HEADING_ROW, _FOLDER_INFO_COL + 2]).Value = "Cummulative Size";

                ((Range)workSheet.Cells[_HEADING_ROW, _FILE_INFO_COL]).Value = "Count";
                ((Range)workSheet.Cells[_HEADING_ROW, _FILE_INFO_COL + 1]).Value = "Size";
                ((Range)workSheet.Cells[_HEADING_ROW, _FILE_INFO_COL + 2]).Value = "Last Create";
                ((Range)workSheet.Cells[_HEADING_ROW, _FILE_INFO_COL + 3]).Value = "Last Write";
                ((Range)workSheet.Cells[_HEADING_ROW, _FILE_INFO_COL + 4]).Value = "Last Access";

                ((Range)workSheet.Cells[_HEADING_ROW, _INITIAL_COL]).Value = startingFolder;

                int startingRow = _Row;
                int startingColumn = _Column;

                //int numberFoldersLocal = 0;
                //int numberFilesLocal = 0;
                //long sizeFilesLocal = 0;

                FileInfo dirInfo = new FileInfo(startingFolder);

                DateTime maxLastCreate = DateTime.MinValue;
                DateTime maxLastWrite = DateTime.MinValue;
                DateTime maxLastAccess = DateTime.MinValue;

                //int column = 1;
                int fontSize = 10;

                // Note: maxLastDate is passed in, even though we don't use the updated value.
                // it is used during the recursion process when ListFolders calls itself.  Do
                // not be mislead and rework this logic!

                ListFolders(startingFolder, _Column, fontSize, 
                    ref _TotalFolders, ref _TotalFiles, ref _TotalSize, 
                    ref maxLastCreate, ref maxLastWrite, ref maxLastAccess, ShowFiles, ShowFolders);

                XlPageOrientation orientation = XlPageOrientation.xlPortrait;

                FormatFolderMapSheet(startingFolder, orientation);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                throw ex;
            }
        }

        public void SaveFolderMap(string startingFolder)
        {
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                string strOutputFile = null;
                try
                {
                    saveFileDialog.FileName = "ExcelFolderMaps.xlsx";
                    saveFileDialog.InitialDirectory = startingFolder;

                    if (System.Windows.Forms.DialogResult.Cancel == saveFileDialog.ShowDialog())
                    {
                        return;
                    }
                    else
                    {
                        strOutputFile = saveFileDialog.FileName;
                    }
                    if (string.IsNullOrEmpty(strOutputFile))
                    {
                        return;
                    }

                    Common.ExcelApplication?.ActiveWorkbook.SaveAs(strOutputFile);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        //public int GetEndOfSectionDown(int startRow, int startCol, int lastPopulatedRow, int initialColumn)
        //{
        //    Worksheet activeSheet = (Worksheet)Common.ExcelApplication.ActiveSheet;

        //    int functionReturnValue = 0;
        //    int matchingRow = 0;

        //    // Search down for a matching cell
        //    matchingRow = ((Range)activeSheet.Cells[startRow, startCol]).End[XlDirection.xlDown].Row;

        //    if (startCol == initialColumn)
        //    //if (startCol == _INITIAL_COL)
        //    {
        //        // We have back'd all the way back to the first column.
        //        // Return either the next matching cell down or the last populated row on the sheet.

        //        if (matchingRow < lastPopulatedRow)
        //        {
        //            // Section ends on the row prior to the match.
        //            functionReturnValue = matchingRow - 1;
        //        }
        //        else
        //        {
        //            // Return end of populated section
        //            functionReturnValue = lastPopulatedRow;
        //        }
        //    }
        //    else
        //    {
        //        if (matchingRow <= lastPopulatedRow)
        //        {
        //            // Back up one column and search down for a populated cell.
        //            // Treat row prior to matching row as new end.
        //            functionReturnValue = GetEndOfSectionDown(startRow, startCol - 1, matchingRow - 1, initialColumn);
        //        }
        //        else
        //        {
        //            // Back up one column and search down for a populated cell.
        //            // Treat end of worksheet as end.
        //            functionReturnValue = GetEndOfSectionDown(startRow, startCol - 1, lastPopulatedRow, initialColumn);
        //        }
        //    }

        //    return functionReturnValue;
        //}

        // Why does this use ref parameters
        private void ListFiles(
            ref string folder, int column, ref int numberFiles, ref long sizeFiles,
            ref DateTime maxLastCreateDate, ref DateTime maxLastWriteDate, ref DateTime maxLastAccessDate,
            bool showFiles)
        {
            string file = null;
            int defaultFileColor = DefaultColorARGB;

            string[] files = Directory.GetFiles(folder);
            // Full path names

            numberFiles = 0;
            sizeFiles = 0;

            bool fileAdded = false;

            foreach (string file_loopVariable in files.OrderBy(n => n))
            {
                file = file_loopVariable;
                // Is there a better way to use FileInfo?  Initialize just one and reuse?
                try
                {
                    fileAdded = false;

                    FileInfo fileInfo = new FileInfo(file);

                    if (FileMatchUsingRegEx)
                    {
                        if (!Regex.Match(fileInfo.Name, FileRegExPattern).Success)
                        {
                            continue;
                        }
                        else
                        {
                            defaultFileColor = FileHighlightColorARGB;
                        }
                    }

                    numberFiles += 1;
                    sizeFiles += fileInfo.Length;

                    // If there is a local file with a more current date, alert our caller.

                    if (maxLastCreateDate < fileInfo.CreationTime)
                    {
                        maxLastCreateDate = fileInfo.CreationTime;
                    }

                    if (maxLastWriteDate < fileInfo.LastWriteTime)
                    {
                        maxLastWriteDate = fileInfo.LastWriteTime;
                    }

                    if (maxLastAccessDate < fileInfo.LastAccessTime)
                    {
                        maxLastAccessDate = fileInfo.LastAccessTime;
                    }

                    // Note, we add file rows if checking for illegal characters or file name length
                    // even if not listing files.

                    if (CheckForIllegalCharacters)
                    {
                        if (HasIllegalFileNameCharacters(fileInfo.Name))
                        {
                            AddFileRow(fileInfo, column, _FILE_FONT_SIZE, false, IllegalCharactersColorARGB);
                            fileAdded = true;
                        }
                    }

                    // TODO: This overrides an IllegalCharacters color.  Maybe bold or something else.
                    if (CheckFileNameLength)
                    {
                        if (fileInfo.Name.Length > MaxFileNameLength)
                        {
                            AddFileRow(fileInfo, column, _FILE_FONT_SIZE, false, IllegalFileNameLengthColorARGB);
                            fileAdded = true;
                        }
                    }

                    if (showFiles & !fileAdded)
                    {
                        AddFileRow(fileInfo, column, _FILE_FONT_SIZE, false, DefaultColorARGB);
                    }
                }
                catch (PathTooLongException)
                {
                    // TODO: We now know that the current file exists at a path+filename that exceeds
                    // the allowable lengths, however, we don't have the size and date info.
                    // Need to do some extra work to get this.  Not sure it is worth it.  For now
                    // just log that the path is too long.

                    string baseFileName = null;

                    if (ShowFolders)
                    {
                        // There will be folder information to show where the file is located so just
                        // display the filename.

                        baseFileName = file.Substring(file.LastIndexOf("\\") + 1);
                    }
                    else
                    {
                        // Need to display the full path to the file since no folders are being displayed.

                        baseFileName = file;
                    }

                    //System.IO.Directory.SetCurrentDirectory(folder)
                    //' Get the base filename skipping the last "\"

                    //' This still blows up!
                    //Dim fileInfo As FileInfo = New FileInfo(baseFileName)
                    //sizeFiles += fileInfo.Length

                    string errorInfo = string.Format("{0} - path({1}) + filename({2}) too long({3})", baseFileName, folder.Length, baseFileName.Length, file.Length);
                    AddErrorRow(errorInfo, column, _FILE_FONT_SIZE, false, IllegalFileNameLengthColorARGB);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                    throw ex;
                }
            }
        }

        //   This routine calls ListFiles() and then calls
        //   itself recursively to descend the folder hierarchy.
        // TODO(crhodes)
        // Why does this use ref parameters

        private void ListFolders(
            string startingFolder, int column, int fontSize,
            ref int numberFoldersCummulative, ref int numberFilesCummulative, ref long sizeFilesCummulative,
            ref DateTime maxLastCreateDate, ref DateTime maxLastWriteDate, ref DateTime maxLastAccessDate,
            bool showFiles = false, bool showFolders = true)
        {
            //Dim objFld As Scripting.Folder
            int intColMax = 0;
            //Dim innerDir As String

            int numberFoldersCummulativeLocal = 0;
            int numberFilesCummulativeLocal = 0;
            long sizeFilesCummulativeLocal = 0;

            int numberFoldersLocal = 0;
            int numberFilesLocal = 0;
            long sizeFilesLocal = 0;

            FileInfo dirInfo = new FileInfo(startingFolder);
            bool folderAdded = false;

            // Start off with Date.MinValue for dates.
            // The ListFiles methods will further update this with file information.

            DateTime localDirCreateDate = DateTime.MinValue;
            DateTime localDirLastWriteDate = DateTime.MinValue;
            DateTime localDirLastAccessDate = DateTime.MinValue;

            if (CheckForIllegalCharacters)
            {
                if (HasIllegalFolderNameCharacters(dirInfo.Name))
                {
                    AddFolderRow(dirInfo, column, fontSize, _MAKE_BOLD, IllegalCharactersColorARGB);
                    folderAdded = true;
                }
            }

            if (showFolders & !folderAdded)
            {
                AddFolderRow(dirInfo, column, fontSize, _MAKE_BOLD);
            }

            // Save the current location as we need to come back and add the totals to this row
            // We already bumped _Row (in AddFolderRow) when we added the Folder we are on
            int currentRow = _Row - 1;
            int currentColumn = column;

            try
            {
                // First list the files in the current folder.  
                // Note: We call ListFiles even if not showing files (blnShowFiles is False) 
                // so we can get information about the files to include in the totals.

                ListFiles(ref startingFolder, column + _INDENT_LEVEL, 
                    ref numberFilesLocal, ref sizeFilesLocal, 
                    ref localDirCreateDate, ref localDirLastWriteDate, ref localDirLastAccessDate, 
                    showFiles);

                // Update the dates with the information from the files that were found.
                // The dates will not have changed if there were no local files.  If
                // that is the case use the directory dates.

                if (localDirCreateDate > maxLastCreateDate)
                {
                    maxLastCreateDate = localDirCreateDate;
                }
                else
                {
                    maxLastCreateDate = dirInfo.CreationTime;
                }

                if (localDirLastWriteDate > maxLastWriteDate)
                {
                    maxLastWriteDate = localDirLastWriteDate;
                }
                else
                {
                    maxLastWriteDate = dirInfo.LastWriteTime;
                }

                if (localDirLastAccessDate > maxLastAccessDate)
                {
                    maxLastAccessDate = localDirLastAccessDate;
                }
                else
                {
                    maxLastAccessDate = dirInfo.LastAccessTime;
                }
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow("ListFolders (ListFiles): " + ex.ToString());
                Log.ERROR(ex, Common.LOG_CATEGORY);
                throw (ex);
                // So we can add code to catch later.
            }

            string[] dirs = Directory.GetDirectories(startingFolder);

            // Then explore each sub folder

            foreach (string innerDir in dirs.OrderBy(n => n))
            {
                try
                {
                    FileInfo innerDirInfo = new FileInfo(innerDir);

                    numberFoldersLocal += 1;

                    int numberFoldersI = 0;
                    int numberFilesI = 0;
                    long sizeFilesI = 0;

                    // Stop if limiting the depth of the exploration

                    if (false == LimitLevels | (true == LimitLevels & FolderLimitLevel > column))
                    {
                        // Call ourselves recursively to display sub folders.

                        localDirCreateDate = DateTime.MinValue;
                        localDirLastWriteDate = DateTime.MinValue;
                        localDirLastAccessDate = DateTime.MinValue;

                        ListFolders(innerDir, column + _INDENT_LEVEL, _FOLDER_FONT_SIZE, ref numberFoldersI, ref numberFilesI, ref sizeFilesI, ref localDirCreateDate, ref localDirLastWriteDate, ref localDirLastAccessDate, showFiles, showFolders);

                        if (localDirCreateDate > maxLastCreateDate)
                        {
                            maxLastCreateDate = localDirCreateDate;
                        }

                        if (localDirLastWriteDate > maxLastWriteDate)
                        {
                            maxLastWriteDate = localDirLastWriteDate;
                        }

                        if (localDirLastAccessDate > maxLastAccessDate)
                        {
                            maxLastAccessDate = localDirLastAccessDate;
                        }

                        numberFoldersCummulativeLocal += numberFoldersI;
                        numberFilesCummulativeLocal += numberFilesI;
                        sizeFilesCummulativeLocal += sizeFilesI;

                        intColMax = column + _INDENT_LEVEL;
                    }
                }
                catch (UnauthorizedAccessException)
                {
                    // Add a message to indicate there is no access to the item previously added.
                    // Move over +1 so grouping isn't impacted.
                    // TODO(crhodes)
                    // Need new color
                    AddErrorRow("<No Access>", column + _INDENT_LEVEL + 1, _FOLDER_FONT_SIZE, false, AccessedColorARGB);
                }
                catch (System.IO.PathTooLongException)
                {
                    //Dim indexS As String = File.LastIndexOf("\")
                    //Dim length As Integer = File.Length
                    //Dim errorInfo As String = String.Format("{0} - Path too long ({1})", File.Substring(File.LastIndexOf("\")), File.Length)
                    //AddErrorRow(errorInfo, column, _FILE_FONT_SIZE)
                }
                catch (Exception ex)
                {
                    Common.WriteToDebugWindow("ListFolders (fe Dir): " + ex.ToString());
                    throw (ex);
                    // TODO: PLException.PLApplicationException.Publish(ex)
                }
            }

            // Add the numberFolders, numberFiles, and sizeFiles to the folder row 
            // now that we know stuff about the files in the current folder and below

            numberFoldersCummulativeLocal += numberFoldersLocal;
            numberFilesCummulativeLocal += numberFilesLocal;
            sizeFilesCummulativeLocal += sizeFilesLocal;

            try
            {
                UpdateFolderRow(currentRow, currentColumn, numberFoldersCummulativeLocal, numberFilesCummulativeLocal, sizeFilesCummulativeLocal, numberFoldersLocal, numberFilesLocal, sizeFilesLocal, maxLastCreateDate, maxLastWriteDate,
                maxLastAccessDate, fontSize);
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow("ListFolders (UpdateFolderRow): " + ex.ToString());
                throw (ex);
                // So we can add code to catch later
            }

            numberFoldersCummulative += numberFoldersCummulativeLocal;
            numberFilesCummulative += numberFilesCummulativeLocal;
            sizeFilesCummulative += sizeFilesCummulativeLocal;

            if (SkipEmptyFolders)
            {
                if (0 == numberFilesLocal & 0 == numberFilesCummulative)
                {
                    ((Range)((Worksheet)Common.ExcelApplication.ActiveSheet).Rows[currentRow]).Delete();
                    _Row -= 1;
                }
            }

            // Save highest column used
            if (intColMax > _Column)
            {
                _Column = intColMax;
            }
        }

        // Formats the Folder Map Sheet and Page.  Can this be changed to
        // call the common worksheet format thing?

        private void FormatFolderMapSheet(string strSheetName, XlPageOrientation enumOrientation = XlPageOrientation.xlPortrait)
        {
            int i = 0;
            Worksheet activeSheet = (Worksheet)Common.ExcelApplication.ActiveSheet;
            Range formatRange;

            //var _with4 = Globals.ThisAddIn.Application;


            for (i = _INITIAL_COL; i <= _Column; i++)
            {
                ((Range)activeSheet.Columns[i]).ColumnWidth = _COL_WIDTH;
            }

            ((Range)activeSheet.Columns[_Column + 1]).ColumnWidth = _NOTE_WIDTH;

            //activeSheet.Range("A2:I2").Select();

            formatRange = activeSheet.Range["A2:I2"];

            formatRange.HorizontalAlignment = Constants.xlGeneral;
            formatRange.VerticalAlignment = Constants.xlBottom;
            formatRange.WrapText = true;
            formatRange.Orientation = 0;
            formatRange.AddIndent = false;
            formatRange.IndentLevel = 0;
            formatRange.ShrinkToFit = false;
            //_with5.ReadingOrder = Constants.xlContext;
            formatRange.MergeCells = false;
            formatRange.Font.Bold = true;

            ((Range)activeSheet.Columns["A:A"]).ColumnWidth = 13.57;
            ((Range)activeSheet.Columns["B:B"]).ColumnWidth = 13.57;
            ((Range)activeSheet.Columns["C:C"]).ColumnWidth = 13.43;

            ((Range)activeSheet.Columns["E:E"]).ColumnWidth = 6.43;
            ((Range)activeSheet.Columns["F:F"]).ColumnWidth = 7.71;
            ((Range)activeSheet.Columns["G:G"]).ColumnWidth = 13.57;
            ((Range)activeSheet.Columns["H:H"]).ColumnWidth = 13.57;
            ((Range)activeSheet.Columns["I:I"]).ColumnWidth = 13.43;

            //_with4.Range["K2"].Select();

            formatRange = activeSheet.Range["K2"];

            formatRange.Font.Name = "Arial";
            formatRange.Font.Size = 12;
            formatRange.Font.Strikethrough = false;
            formatRange.Font.Superscript = false;
            formatRange.Font.Subscript = false;
            formatRange.Font.OutlineFont = false;
            formatRange.Font.Shadow = false;
            formatRange.Font.Underline = XlUnderlineStyle.xlUnderlineStyleNone;
            formatRange.Font.ColorIndex = Constants.xlAutomatic;
            formatRange.Font.Bold = true;

            ((Range)activeSheet.Columns["A:J"]).Group();
            ((Range)activeSheet.Columns["A:D"]).Group();
            ((Range)activeSheet.Columns["G:I"]).Group();

            activeSheet.Range["A1"].Select();

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="dirInfo"></param>
        /// <param name="column"></param>
        /// <param name="fontSize"></param>
        /// <param name="makeBold"></param>
        /// <param name="fontColor"></param>
        private void AddFolderRow(FileInfo dirInfo, int column = 1, int fontSize = 10, 
            bool makeBold = false, Int32 fontColor = (int)ConsoleColor.Black)
        {
            string strS = null;

            // HACK(crhodes)
            // Not sure what this was all about

            //if (true == _GroupResults)
            //{
            //    if (_GroupLevel == column)
            //    {
            //        // Mark starting point.  We may need it later for grouping.  Don't reset.
            //        if (0 == _GroupStartRow)
            //        {
            //            _GroupStartRow = _Row;
            //        }
            //    }
            //}

            if (FolderMatchUsingRegEx)
            {
                if ((Regex.Match(dirInfo.Name, FolderRegExPattern).Success))
                {
                    fontColor = FolderHighlightColorARGB;
                }
            }

            Worksheet activeSheet = (Worksheet)Common.ExcelApplication.ActiveSheet;

            Range folderRow = (Range)activeSheet.Cells[_Row, column];
            // If we are not displaying folders but for some reason
            // are calling AddFolderRow(), then use full path.

            if (ShowFolders)
            {
                folderRow.Value = string.Format("{0}\\", dirInfo.Name);
            }
            else
            {
                //folderRow.Value = string.Format("{0}\\", dirInfo.FullName);
                folderRow.Value = string.Format("{0}\\", dirInfo.DirectoryName);
                folderRow.Offset[0, 1].Value = string.Format("{0}\\", dirInfo.FullName);
            }

            folderRow.Font.Bold = makeBold;
            folderRow.Font.Size = fontSize;
            folderRow.Font.Color = fontColor;

            // Now check if we need to do any grouping.

            // HACK(crhodes)
            // Not sure what this is all about
            //if (_GroupStartRow > 0)
            //{
            //    // We have been adding rows to a possible grouping set.
            //    if (_GroupLevel > column)
            //    {
            //        // We have transitioned up the chain above the grouping set.
            //        strS = string.Format("{0}:{1}", _GroupStartRow, _Row - 1);
            //        //strS = _GroupStartRow + ":" + _Row - 1;
            //        ((Range)((Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Rows[strS]).Group();
            //        // Reset the grouping counter.
            //        _GroupStartRow = 0;
            //    }
            //}

            // Next row to add content
            _Row += 1;
        }

        //   Adds a line to the spreadsheet indented at the appropriate level.

        private void AddFileRow(FileInfo fileInfo, 
            int column = 1, int fontSize = 10, bool makeBold = false, 
            Int32 fontColor = (int)ConsoleColor.Black)
        {
            Worksheet activeSheet = (Worksheet)Common.ExcelApplication.ActiveSheet;
            Range formatRange;

            string strS = null;

            // HACK(crhodes)
            // Not sure what this was all about.  Think it was to do automatic grouping as the rows were added,
            // but it seems to cause more problems than it solves.
            // Maybe we can add grouping after the fact if needed.  For now just ignore.

            //if (true == _GroupResults)
            //{
            //    if (_GroupLevel == column)
            //    {
            //        // Mark starting point.  We may need it later for grouping.  Don't reset.
            //        if (0 == _GroupStartRow)
            //        {
            //            _GroupStartRow = _Row;
            //        }
            //    }
            //}

            //if (_TableSyleOuput)
            //{
            //    formatRange = (Range)activeSheet.Cells[_Row, _INITIAL_COL];
            //}
            //else
            //{
                formatRange = (Range)activeSheet.Cells[_Row, column];
            //}

            // If we are not displaying folders but for some reason 
            // are calling AddFileRow(), then use full path.

            if (ShowFolders)
            {
                formatRange.Value = string.Format(" - {0}", fileInfo.Name);
            }
            else
            {
                //formatRange.Value = string.Format(" - {0}", fileInfo.FullName);
                // HACK(crhodes)
                // This let's us come back and wrap a table around everything to see if there are duplicate files.
                formatRange.Value = string.Format("{0}", fileInfo.DirectoryName);
                formatRange.Offset[0, 1].Value = string.Format("{0}", fileInfo.Name);
                formatRange.Offset[0, 1].Font.Size = fontSize;
                formatRange.Offset[0, 1].Font.Color = fontColor;

                // TODO(crhodes)
                // Decide if want to implement this
                //if (_CalculateCRC)
                //{

                //    try
                //    {
                //        using (var fileStream = fileInfo.OpenRead())
                //        {
                //            byte[] fileBytes = File.ReadAllBytes(fileInfo.FullName);

                //            formatRange.Offset[0, 2].Value = Crc32CAlgorithm.Compute(fileBytes).ToString();
                //        }
                //    }
                //    catch (System.IO.IOException ioex)
                //    {
                //        formatRange.Offset[0, 2].Value = ioex.ToString();
                //    }
                //    catch (Exception ex)
                //    {
                //        var et = ex.GetType();
                //        var be = ex.GetBaseException();
                //        MessageBox.Show(ex.ToString());
                //    }
                //}
            }

            formatRange.Font.Bold = makeBold;
            formatRange.Font.Size = fontSize;
            //.ColorIndex = fontColor
            formatRange.Font.Color = fontColor;

            // Start at _FILE_INFO_COL + 1 as we don't display file count on file row.

            formatRange = (Range)activeSheet.Cells[_Row, _FILE_INFO_COL + 1];
            formatRange.Value = fileInfo.Length;
            formatRange.NumberFormat = "#,##0_);(#,##0)";

            formatRange.Font.Bold = makeBold;
            formatRange.Font.Size = fontSize;

            Range rng = default(Range);
            System.DateTime dateD = default(System.DateTime);

            rng = (Range)activeSheet.Cells[_Row, _FILE_INFO_COL + 2];
            dateD = fileInfo.CreationTime;
            rng.Value = dateD;

            ColorCodeDate(rng, dateD, false, fontSize, _DateType.LastCreate);

            rng = (Range)activeSheet.Cells[_Row, _FILE_INFO_COL + 3];
            dateD = fileInfo.LastWriteTime;
            rng.Value = dateD;

            ColorCodeDate(rng, dateD, false, fontSize, _DateType.LastWrite);

            rng = (Range)activeSheet.Cells[_Row, _FILE_INFO_COL + 4];
            dateD = fileInfo.LastAccessTime;
            rng.Value = dateD;

            ColorCodeDate(rng, dateD, false, fontSize, _DateType.LastAccess);

            // Now check if we need to do any grouping.

            //if (_GroupStartRow > 0)
            //{
            //    // We have been adding rows to a possible grouping set.
            //    if (_GroupLevel > column)
            //    {
            //        // We have transitioned up the chain above the grouping set.
            //        //strS = _GroupStartRow + ":" + _Row - 1;
            //        strS = string.Format("{0}:{1}", _GroupStartRow, _Row - 1);
            //        ((Range)activeSheet.Rows[strS]).Group();
            //        // Reset the grouping counter.
            //        _GroupStartRow = 0;
            //    }
            //}

            // Next row to add content.
            _Row = _Row + 1;
        }

        private void AddErrorRow(string errorInfo, 
            int column = 1, int fontSize = 10, bool makeBold = false, 
            Int32 fontColor = (int)ConsoleColor.Black)
        {
            Worksheet activeSheet = (Worksheet)Common.ExcelApplication.ActiveSheet;
            Range formatRange;

            string strS = null;

            //    Debug.Print m_lngRow & " " & intCol & ":" & m_intGroupStartRow & " " & strText

            //if (true == _GroupResults)
            //{
            //    if (_GroupLevel == column)
            //    {
            //        // Mark starting point.  We may need it later for grouping.  Don't reset.
            //        if (0 == _GroupStartRow)
            //        {
            //            _GroupStartRow = _Row;
            //        }
            //    }
            //}

            var _with15 = Common.ExcelApplication;
            formatRange = (Range)activeSheet.Cells[_Row, column];
            formatRange.Value = string.Format(" - {0}", errorInfo);

            formatRange.Font.Bold = makeBold;
            formatRange.Font.Size = fontSize;
            formatRange.Font.Color = fontColor;

            // Now check if we need to do any grouping.

            //if (_GroupStartRow > 0)
            //{
            //    // We have been adding rows to a possible grouping set.
            //    if (_GroupLevel > column)
            //    {
            //        // We have transitioned up the chain above the grouping set.
            //        //strS = _GroupStartRow + ":" + _Row - 1;
            //        strS = string.Format("{0}:{1}", _GroupStartRow, _Row - 1);
            //        ((Range)activeSheet.Rows[strS]).Group();
            //        // Reset the grouping counter.
            //        _GroupStartRow = 0;
            //    }
            //}

            // Next row to add content.
            _Row = _Row + 1;
        }

        private void UpdateFolderRow(int row, int column, 
            int numberFoldersCummulative, int numberFilesCummulative, long sizeFilesCummulative, 
            int numberFolders, int numberFiles, long sizeFiles, 
            DateTime maxLastCreateDate, DateTime maxLastWriteDate, DateTime maxLastAccessDate, 
            int fontSize = 10)
        {
            Worksheet activeSheet = (Worksheet)Common.ExcelApplication.ActiveSheet;
            Range formatRange;

            var excelApp = Common.ExcelApplication;

            formatRange = (Range)activeSheet.Cells[row, _FOLDER_INFO_COL];
            formatRange.Value = numberFoldersCummulative;
            formatRange.NumberFormat = "#,##0_);(#,##0)";
            formatRange.Font.Size = fontSize;
            // .Bold = blnBold

            formatRange = (Range)activeSheet.Cells[row, _FOLDER_INFO_COL + 1];
            formatRange.Value = numberFilesCummulative;
            formatRange.NumberFormat = "#,##0_);(#,##0)";
            formatRange.Font.Size = fontSize;

            formatRange = (Range)activeSheet.Cells[row, _FOLDER_INFO_COL + 2];
            formatRange.Value = sizeFilesCummulative;
            formatRange.NumberFormat = "#,##0_);(#,##0)";
            formatRange.Font.Size = fontSize;

            formatRange = (Range)activeSheet.Cells[row, _FILE_INFO_COL];
            formatRange.Value = numberFiles;
            formatRange.NumberFormat = "#,##0_);(#,##0)";
            formatRange.Font.Size = fontSize;

            formatRange = (Range)activeSheet.Cells[row, _FILE_INFO_COL + 1];
            formatRange.Value = sizeFiles;
            formatRange.NumberFormat = "#,##0_);(#,##0)";
            formatRange.Font.Size = fontSize;

            DateTime dateD = default(DateTime);

            formatRange = (Range)activeSheet.Cells[row, _FILE_INFO_COL + 2];
            dateD = maxLastCreateDate;
            formatRange.Value = dateD;

            ColorCodeDate(formatRange, dateD, false, fontSize, _DateType.LastCreate);

            formatRange = (Range)activeSheet.Cells[row, _FILE_INFO_COL + 3];
            dateD = maxLastWriteDate;
            formatRange.Value = dateD;

            ColorCodeDate(formatRange, dateD, false, fontSize, _DateType.LastWrite);

            formatRange = (Range)activeSheet.Cells[row, _FILE_INFO_COL + 4];
            dateD = maxLastAccessDate;
            formatRange.Value = dateD;

            ColorCodeDate(formatRange, dateD, false, fontSize, _DateType.LastAccess);
        }

        private void ColorCodeDate(Range formatRange, DateTime dt, bool makeBold, int fontSize, _DateType dateType)
        {
            formatRange.Font.Bold = makeBold;
            formatRange.Font.Size = fontSize;

            if (ColorCodeDates)
            {
                switch (dateType)
                {
                    case _DateType.LastCreate:
                        if ((DateTime.Compare(DateTime.Now, dt.AddMonths(MonthsSinceCreated)) > 0))
                        {
                            formatRange.Font.Color = CreatedColorARGB;
                        }
                        else
                        {
                            formatRange.Font.Color = DefaultColorARGB;
                        }

                        break;

                    case _DateType.LastWrite:
                        if ((DateTime.Compare(DateTime.Now, dt.AddMonths(MonthsSinceWritten)) > 0))
                        {
                            formatRange.Font.Color = WrittenColorARGB;
                        }
                        else
                        {
                            formatRange.Font.Color = DefaultColorARGB;
                        }

                        break;

                    case _DateType.LastAccess:
                        if ((DateTime.Compare(DateTime.Now, dt.AddMonths(MonthsSinceAccessed)) > 0))
                        {
                            formatRange.Font.Color = AccessedColorARGB; 
                        }
                        else
                        {
                            formatRange.Font.Color = DefaultColorARGB;
                        }

                        break;
                }
            }
        }

        private bool HasIllegalFileNameCharacters(string name)
        {
            Match illegalCharactersMatch = Regex.Match(name, IllegalFileCharacters);

            return illegalCharactersMatch.Success;
        }

        private bool HasIllegalFolderNameCharacters(string name)
        {
            Match illegalCharactersMatch = Regex.Match(name, IllegalFolderCharacters);

            return illegalCharactersMatch.Success;
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
