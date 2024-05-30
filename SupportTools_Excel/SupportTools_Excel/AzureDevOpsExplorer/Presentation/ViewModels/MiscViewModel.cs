using Prism.Commands;
using SupportTools_Excel.AzureDevOpsExplorer.Presentation.Views;
using SupportTools_Excel.Core.Presentation.ViewModels;

using VNC;
using VNC.Core.Mvvm;

namespace SupportTools_Excel.AzureDevOpsExplorer.Presentation.ViewModels
{
    public class MiscViewModel : ViewModelBase, IAZDOMiscViewModel
    {
        #region Constructors and Load

        // View First

        public MiscViewModel()
        {
            long startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            CodeChurnCommand = new DelegateCommand(OnCodeChurnExecute, OnCodeChurnCanExecute);
            SearchForFilesCommand = new DelegateCommand(OnSearchForFilesExecute, OnSearchForFilesCanExecute);
            UnMergedChangesCommand = new DelegateCommand(OnUnMergedChangesExecute, OnUnMergedChangesCanExecute);

            TeamProjectCollection_DoubleClick_Command = new DelegateCommand(TeamProjectCollection_DoubleClick);
            TeamProjectPath_DoubleClick_Command = new DelegateCommand(TeamProjectPath_DoubleClick);

            FilePattern_DoubleClick_Command = new DelegateCommand(FilePattern_DoubleClick);

            // TODO(crhodes)
            // Decide if we want defaults
            //XXX = new XXXWrapper(new Domain.XXX());

            //InitializeRows();

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // ViewModel First
        // Calling base(view) wires this ViewModel into the View

        public MiscViewModel(Misc view) : base(view)
        {
            long startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            CodeChurnCommand = new DelegateCommand(OnCodeChurnExecute, OnCodeChurnCanExecute);
            SearchForFilesCommand = new DelegateCommand(OnSearchForFilesExecute, OnSearchForFilesCanExecute);
            UnMergedChangesCommand = new DelegateCommand(OnUnMergedChangesExecute, OnUnMergedChangesCanExecute);

            TeamProjectCollection_DoubleClick_Command = new DelegateCommand(TeamProjectCollection_DoubleClick);
            TeamProjectPath_DoubleClick_Command = new DelegateCommand(TeamProjectPath_DoubleClick);

            FilePattern_DoubleClick_Command = new DelegateCommand(FilePattern_DoubleClick);

            //InitializeRows();

            //View = view;

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region Fields and Properties

        public string FilePatternText { get; set; }
        public string FilePatternToolTip { get; set; }

        public string TeamProjectPathText { get; set; }
        public string TeamProjectPathToolTip { get; set; }

        public string TeamProjectCollectionText { get; set; }
        public string TeamProjectCollectionToolTip { get; set; }

        public string CodeChurnContent { get; set; }
        public string CodeChurnToolTip { get; set; }

        public string SearchForFilesContent { get; set; }
        public string SearchForFilesToolTip { get; set; }

        public string UnMergedChangesContent { get; set; }
        public string UnMergedChangesToolTip { get; set; }

        #endregion

        #region Commands

        #region SearchForFiles Command

        public DelegateCommand SearchForFilesCommand { get; set; }

        public void OnSearchForFilesExecute()
        {
            // TODO(crhodes)
            // Do something amazing.

            Message = "Search For Files";
        }

        public bool OnSearchForFilesCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.

            return true;
        }

        #endregion

        #region CodeChurn Command

        public DelegateCommand CodeChurnCommand { get; set; }

        public void OnCodeChurnExecute()
        {
            // TODO(crhodes)
            // Do something amazing.
            Message = "Code Churn";
        }

        public bool OnCodeChurnCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion

        #region UnmergedChanges Command

        public DelegateCommand UnMergedChangesCommand { get; set; }

        public DelegateCommand TeamProjectCollection_DoubleClick_Command  { get; set; }
        public DelegateCommand TeamProjectPath_DoubleClick_Command { get; set; }
        public DelegateCommand FilePattern_DoubleClick_Command { get; set; }

        public void OnUnMergedChangesExecute()
        {
            // TODO(crhodes)
            // Do something amazing.

            Message = "Unmerged Changes";
        }

        public bool OnUnMergedChangesCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.

            return true;
        }

        #endregion

        #endregion Commands

        string _message = "Double Click Text Boxes to test events";
        public string Message
        {
            get
            {
                return _message;
            }
            set
            {
                _message = value;
                OnPropertyChanged();
            }
        }

        public void FilePattern_DoubleClick()
        {
            Message = "TeamProjectPath_DoubleClick";
        }

        public void TeamProjectCollection_DoubleClick()
        {
            Message = "TeamProjectCollection_DoubleClick";
        }

        public void TeamProjectPath_DoubleClick()
        {
            Message = "TeamProjectPath_DoubleClick";
        }
    }
}
