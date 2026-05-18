using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Input;

using System.Windows.Media;

using DevExpress.Utils.CommonDialogs;

using Prism.Commands;
using Prism.Dialogs;
using Prism.Events;

using VNC;
using VNC.Core.Events;
using VNC.Core.Mvvm;
using VNC.Core.Services;

using VNCExcelToolsApplication.Presentation.Views;

namespace VNCExcelToolsApplication.Presentation.ViewModels
{
    public class CreateFolderMapViewModel : EventViewModelBase, ICreateFolderMapViewModel, IInstanceCountVM
    {
        public class RegEx
        {
            // SharePoint Folder/File/Document Libraries may not contain any of the following characters
            //   / \ : * ? " < > | <TAB> { } % ~ &
            // nor may they end in periods or contain embedded double periods.
            // The following regular expressions capture these rules.  

            internal const string cIllegalFileCharacters = "[/\\\\:\\*\\?\"<>\\|#\\{}%~&]|\\.\\.";

            internal const string cIllegalFolderCharacters = "[:\\*\\?\"<>\\|#\\{}%~&]";

        }

        #region Fields (none)

        FolderBrowserDialog _folderBrowserDialog = new FolderBrowserDialog();

        #endregion

        #region Constructors, Initialization, and Load

        public CreateFolderMapViewModel(
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

            if (Common.VNCLogging.ViewModelLow) Log.VIEWMODEL_LOW("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region Enums (none)



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
        private Int32 _folderLimitLevels = 1;
        public Int32 FolderLimitLevels
        {
            get => _folderLimitLevels;
            set
            {
                if (_folderLimitLevels == value) return;
                _folderLimitLevels = value;
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

        public string FolderRegExPatternToolTip { get; set; }

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

        public string FileRegExPatternToolTip { get; set; }

        private Color _folderHighlightColor;
        public Color FolderHighlightColor
        {
            get => _folderHighlightColor;
            set
            {
                if (_folderHighlightColor == value)
                    return;
                _folderHighlightColor = value;
                OnPropertyChanged();
            }
        }

        private Color _fileHighlightColor;
        public Color FileHighlightColor
        {
            get => _fileHighlightColor;
            set
            {
                if (_fileHighlightColor == value)
                    return;
                _fileHighlightColor = value;
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

        private Color _defaultColor = Colors.Black;
        public Color DefaultColor
        {
            get => _defaultColor;
            set
            {
                if (_defaultColor == value)
                    return;
                _defaultColor = value;
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

        private Color _createdColor = Colors.Red;
        public Color CreatedColor
        {
            get => _createdColor;
            set
            {
                if (_createdColor == value)
                    return;
                _createdColor = value;
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

        private Color _writtenColor = Colors.Green;
        public Color WrittenColor
        {
            get => _writtenColor;
            set
            {
                if (_writtenColor == value)
                    return;
                _writtenColor = value;
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

        private Color _accessedColor = Colors.Blue;
        public Color AccessedColor
        {
            get => _accessedColor;
            set
            {
                if (_accessedColor == value)
                    return;
                _accessedColor = value;
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

        private Color _illegalCharactersColor = Colors.Orange;
        public Color IllegalCharactersColor
        {
            get => _illegalCharactersColor;
            set
            {
                if (_illegalCharactersColor == value)
                    return;
                _illegalCharactersColor = value;
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

        public string IllegalFolderCharactersToolTip { get; set; }

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

        public string IllegalFileCharactersToolTip { get; set; }

        private Color _illegalPathLengthColor = Colors.Red;
        public Color IllegalPathLengthColor
        {
            get => _illegalPathLengthColor;
            set
            {
                if (_illegalPathLengthColor == value)
                    return;
                _illegalPathLengthColor = value;
                OnPropertyChanged();
            }
        }

        private int _maxPathLength = 250;
        public int MaxPathLength
        {
            get => _maxPathLength;
            set
            {
                if (_maxPathLength == value)
                    return;
                _maxPathLength = value;
                OnPropertyChanged();
            }
        }

        #endregion

        #endregion

        #region Event Handlers (none)

        // FIXME(crhodes)
        // This was from the code behind in SupportTools_Excel.User_Interface.Forms.frmExcel_FolderMaps.cs
        // Remove as we add things to View and ViewModel
        //private void btnSelectFolder_Click(object sender, EventArgs e)
        //{
        //    FolderBrowserDialog1.ShowNewFolderButton = false;

        //    if (txtStartingFolder.Text.Length > 0)
        //    {
        //        FolderBrowserDialog1.SelectedPath = txtStartingFolder.Text;
        //    }

        //    if (FolderBrowserDialog1.ShowDialog() == DialogResult.OK)
        //    {
        //        txtStartingFolder.Text = FolderBrowserDialog1.SelectedPath;
        //    }
        //}

        //private void chkColorCodeDates_CheckedChanged(object sender, EventArgs e)
        //{
        //    if (chkColorCodeDates.Checked)
        //    {
        //        pnlDefaultColor.Enabled = true;
        //        pnlMonthCreatedColor.Enabled = true;
        //        pnlMonthAccessedColor.Enabled = true;
        //        pnlMonthWrittenColor.Enabled = true;

        //        txtMonthsSinceCreated.Enabled = true;
        //        txtMonthsSinceWritten.Enabled = true;
        //        txtMonthsSinceAccessed.Enabled = true;

        //        spnMonthsSinceCreated.Enabled = true;
        //        spnMonthsSinceWritten.Enabled = true;
        //        spnMonthsSinceAccessed.Enabled = true;
        //    }
        //    else
        //    {
        //        pnlDefaultColor.Enabled = false;
        //        pnlMonthCreatedColor.Enabled = false;
        //        pnlMonthAccessedColor.Enabled = false;
        //        pnlMonthWrittenColor.Enabled = false;

        //        txtMonthsSinceCreated.Enabled = false;
        //        txtMonthsSinceWritten.Enabled = false;
        //        txtMonthsSinceAccessed.Enabled = false;

        //        spnMonthsSinceCreated.Enabled = false;
        //        spnMonthsSinceWritten.Enabled = false;
        //        spnMonthsSinceAccessed.Enabled = false;
        //    }
        //}

        //private void chkFileNameLength_CheckedChanged(object sender, EventArgs e)
        //{
        //    if (chkFileNameLength.Checked)
        //    {
        //        txtMaxFileNameLength.Enabled = true;
        //        pnlIllegalFileNameLengthColor.Enabled = true;
        //    }
        //    else
        //    {
        //        txtMaxFileNameLength.Enabled = false;
        //        pnlIllegalFileNameLengthColor.Enabled = false;
        //    }
        //}

        //private void chkGroupResults_CheckedChanged(object sender, EventArgs e)
        //{
        //    if (chkGroupResults.Checked)
        //    {
        //        txtGroupLevel.Enabled = true;
        //        spnGroupLevel.Enabled = true;
        //    }
        //    else
        //    {
        //        txtGroupLevel.Enabled = false;
        //        spnGroupLevel.Enabled = false;
        //    }
        //}

        //private void chkIllegalCharacters_CheckedChanged(object sender, EventArgs e)
        //{
        //    if (chkIllegalCharacters.Checked)
        //    {
        //        pnlIllegalCharactersColor.Enabled = true;
        //        txtIllegalFileCharacters.Enabled = true;
        //        txtIllegalFolderCharacters.Enabled = true;
        //    }
        //    else
        //    {
        //        pnlIllegalCharactersColor.Enabled = false;
        //        txtIllegalFileCharacters.Enabled = false;
        //        txtIllegalFolderCharacters.Enabled = false;
        //    }
        //}

        //private void chkLimitLevels_CheckedChanged(object sender, EventArgs e)
        //{
        //    if (chkLimitLevels.Checked)
        //    {
        //        txtLimitLevel.Enabled = true;
        //        spnLimitLevel.Enabled = true;
        //    }
        //    else
        //    {
        //        txtLimitLevel.Enabled = false;
        //        spnLimitLevel.Enabled = false;
        //    }
        //}

        //private void chkPatternMatchFileOutput_CheckedChanged(object sender, EventArgs e)
        //{

        //}

        //private void chkPatternMatchFolderHighlight_CheckedChanged(object sender, EventArgs e)
        //{

        //}

        //private void chkShowFiles_CheckedChanged(object sender, EventArgs e)
        //{
        //    if (chkShowFiles.Checked)
        //    {
        //        chkPatternMatchFileOutput.Enabled = true;
        //        txtFileMatchPattern.Enabled = true;
        //        chkSkipFoldersWithNoFiles.Enabled = true;
        //    }
        //    else
        //    {
        //        chkPatternMatchFileOutput.Enabled = false;
        //        chkPatternMatchFileOutput.CheckState = CheckState.Unchecked;
        //        txtFileMatchPattern.Enabled = false;
        //        chkSkipFoldersWithNoFiles.Enabled = false;
        //    }
        //}

        //private void chkShowFolders_CheckedChanged(object sender, EventArgs e)
        //{

        //}

        //private void chkSkipFoldersWithNoFiles_CheckedChanged(object sender, EventArgs e)
        //{

        //}

        //private void cmdCreateFolderMap_Click(object sender, EventArgs e)
        //{
        //    if (txtStartingFolder.Text.Length > 0)
        //    {
        //        DialogResult = System.Windows.Forms.DialogResult.OK;
        //        Close();
        //    }
        //    else
        //    {
        //        MessageBox.Show("Must select starting folder");
        //        txtStartingFolder.Focus();
        //    }
        //}

        //private void frmExcel_FolderMaps_Load(object sender, EventArgs e)
        //{
        //    txtLimitLevel.Text = "1";
        //    txtLimitLevel.Enabled = false;
        //    spnLimitLevel.Minimum = 1;
        //    spnLimitLevel.Value = 1;
        //    spnLimitLevel.Enabled = false;

        //    txtGroupLevel.Text = "3";
        //    txtGroupLevel.Enabled = false;
        //    spnGroupLevel.Minimum = 3;
        //    spnGroupLevel.Value = 3;
        //    spnGroupLevel.Enabled = false;

        //    txtMonthsSinceCreated.Text = "24";
        //    txtMonthsSinceCreated.Enabled = false;
        //    spnMonthsSinceCreated.Minimum = 1;
        //    spnMonthsSinceCreated.Value = 24;
        //    spnMonthsSinceCreated.Enabled = false;

        //    txtMonthsSinceAccessed.Text = "24";
        //    txtMonthsSinceAccessed.Enabled = false;
        //    spnMonthsSinceAccessed.Minimum = 1;
        //    spnMonthsSinceAccessed.Value = 24;
        //    spnMonthsSinceAccessed.Enabled = false;

        //    txtMonthsSinceWritten.Text = "24";
        //    txtMonthsSinceWritten.Enabled = false;
        //    spnMonthsSinceWritten.Minimum = 1;
        //    spnMonthsSinceWritten.Value = 24;
        //    spnMonthsSinceWritten.Enabled = false;

        //    txtIllegalFileCharacters.Text = RegEx.cIllegalFileCharacters;
        //    txtIllegalFileCharacters.Enabled = false;

        //    txtIllegalFolderCharacters.Text = RegEx.cIllegalFolderCharacters;
        //    txtIllegalFolderCharacters.Enabled = false;

        //    txtMaxFileNameLength.Text = Common.cMaxFileNameLength.ToString();
        //    txtMaxFileNameLength.Enabled = false;
        //}

        //private void ColorBox_DoubleClick(object sender, EventArgs e)
        //{
        //    ColorDialog1.Color = ((Panel)sender).BackColor;

        //    if (ColorDialog1.ShowDialog() != DialogResult.Cancel)
        //    {
        //        ((Panel)sender).BackColor = ColorDialog1.Color;
        //    }
        //}

        //private void spnGroupLevel_ValueChanged(object sender, EventArgs e)
        //{
        //    txtGroupLevel.Text = spnGroupLevel.Value.ToString();
        //}

        //private void spnLimitLevel_ValueChanged(object sender, EventArgs e)
        //{
        //    txtLimitLevel.Text = spnLimitLevel.Value.ToString();
        //}

        //private void spnMonthsSinceAccessed_ValueChanged(object sender, EventArgs e)
        //{
        //    txtMonthsSinceAccessed.Text = spnMonthsSinceAccessed.Value.ToString();
        //}

        //private void spnMonthsSinceCreated_ValueChanged(object sender, EventArgs e)
        //{
        //    txtMonthsSinceCreated.Text = spnMonthsSinceCreated.Value.ToString();
        //}

        //private void spnMonthsSinceWritten_ValueChanged(object sender, EventArgs e)
        //{
        //    txtMonthsSinceWritten.Text = spnMonthsSinceWritten.Value.ToString();
        //}

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
            if (Common.VNCLogging.EventHandler) startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);
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

            if (Common.VNCLogging.EventHandler) Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
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
            if (Common.VNCLogging.EventHandler) startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);
            // TODO(crhodes)
            // Do something amazing.

            Message = "Cool, you called CreateFolderMap";

            PublishStatusMessage(Message);

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

        #endregion

        #region Protected Methods (none)



        #endregion

        #region Private Methods (none)



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
