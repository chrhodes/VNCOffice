using DevExpress.Pdf.Native;

using Prism.Commands;

using VNCVisioToolsApplication.Core;
using VNCVisioToolsApplication.Presentation.Views;

using System;

using VNC;
using VNC.Core.Mvvm;

using MSVisio = Microsoft.Office.Interop.Visio;

namespace VNCVisioToolsApplication.Presentation.ViewModels
{
    public class DuplicatePageViewModel : ViewModelBase, IDuplicatePageViewModel
    {
        #region Constructors and Load

        // View First
        // View creates new ViewModel in code or Xaml
        // or ViewModel passed into View constructor

        public DuplicatePageViewModel()
        {
            long startTicks = Log.TRACE("Enter", Common.LOG_CATEGORY);

            // TODO(crhodes)
            // Decide if we want defaults
            //DuplicatePage = new DuplicatePageWrapper(new Domain.DuplicatePage());

            InitializeViewModel();

            Log.TRACE("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // ViewModel First
        // Calling base(view) wires this ViewModel into the View

        public DuplicatePageViewModel(DuplicatePage view) : base(view)
        {
            long startTicks = Log.TRACE("Enter", Common.LOG_CATEGORY);

            InitializeViewModel();

            Log.TRACE("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void InitializeViewModel()
        {
            // TODO(crhodes)
            // Initialize any controls and/or properties that need to be

            Message_DoubleClick_Command = new DelegateCommand(Message_DoubleClick);

            LoadPageCommand = new DelegateCommand(OnLoadPageExecute, OnLoadPageCanExecute);

            SavePageCommand = new DelegateCommand(OnSavePageExecute, OnSavePageCanExecute);

            UpdateCurrentPage();

            //InitializeRows();
        }

        private void UpdateCurrentPage()
        {
            MSVisio.Application app = Common.VisioApplication;
            CurrentPageName = app.ActivePage.NameU;
            CurrentPageIndex = app.ActivePage.Index;
            NewPageName = $"{CurrentPageName} X";
        }

        #endregion

        #region Fields

        #endregion

        #region Properties



        string _message = "Click Button to do something";
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

        private string _currentPageName;
        public string CurrentPageName
        {
            get => _currentPageName;
            set
            {
                if (_currentPageName == value)
                    return;
                _currentPageName = value;
                OnPropertyChanged();
            }
        }

        private short _currentPageIndex;
        public short CurrentPageIndex
        {
            get => _currentPageIndex;
            set
            {
                if (_currentPageIndex == value)
                    return;
                _currentPageIndex = value;
                OnPropertyChanged();
            }
        }

        private string _newPageName;
        public string NewPageName
        {
            get => _newPageName;
            set
            {
                if (_newPageName == value)
                    return;
                _newPageName = value;
                OnPropertyChanged();
            }
        }

        // TODO(crhodes)
        // This is for a Grid or List

        // public System.Collections.ObjectModel.ObservableCollection<string> SelectedFruits { get; set; }

        // public System.Collections.ObjectModel.ObservableCollection<DuplicatePageWrapper> Rows { get; set; }

        // // and the SelectedItem in the Grid or List

        // DuplicatePageWrapper _selectedItem;
        // public DuplicatePageWrapper SelectedItem
        // {
        // get
        // {
        // return _selectedItem;
        // }
        // set
        // {
        // _selectedItem = value;
        // OnPropertyChanged();
        // }
        // }

        // Don't forget to uncomment InitializeRows in InitializeViewModel()

        // void InitializeRows()
        // {
        // Rows = new System.Collections.ObjectModel.ObservableCollection<DuplicatePageWrapper>();
        // Rows.Add(new DuplicatePageWrapper(new Domain.DuplicatePage(){ StringProperty ="Red", IntProperty = 1}));
        // Rows.Add(new DuplicatePageWrapper(new Domain.DuplicatePage(){ StringProperty = "Green", IntProperty = 2 }));
        // Rows.Add(new DuplicatePageWrapper(new Domain.DuplicatePage(){ StringProperty = "Blue", IntProperty = 3 }));

        // OnPropertyChanged("Rows");
        // }		

        #endregion

        #region Commands

        #region Control Commands (Not Buttons)

        public DelegateCommand Message_DoubleClick_Command { get; set; }

        public void Message_DoubleClick()
        {
            Message = "Message DoubleClicked!";
        }

        #endregion

        #region LoadPage Command

        public DelegateCommand LoadPageCommand { get; set; }
        public string LoadPageContent { get; set; } = "LoadPage";
        public string LoadPageToolTip { get; set; } = "LoadPage ToolTip";

        // Can get fancy and use Resources
        //public string LoadPageContent { get; set; } = "ViewName_LoadPageContent";
        //public string LoadPageToolTip { get; set; } = "ViewName_LoadPageContentToolTip";

        // Put these in Resource File
        //    <system:String x:Key="ViewName_LoadPageContent">LoadPage</system:String>
        //    <system:String x:Key="ViewName_LoadPageContentToolTip">LoadPage ToolTip</system:String>  

        public void OnLoadPageExecute()
        {
            // TODO(crhodes)
            // Do something amazing.
            Message = "Cool, you called LoadPage";
            Common.EventAggregator.GetEvent<LoadPageEvent>().Publish();

            UpdateCurrentPage();

            // Start Cut Four

            // Put this in places that listen for event
            //Common.EventAggregator.GetEvent<LoadPageEvent>().Subscribe(LoadPage);

            // End Cut Four
        }

        public bool OnLoadPageCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion

        #region SavePage Command

        public DelegateCommand SavePageCommand { get; set; }
        public string SavePageContent { get; set; } = "SavePage";
        public string SavePageToolTip { get; set; } = "SavePage ToolTip";

        // Can get fancy and use Resources
        //public string SavePageContent { get; set; } = "ViewName_SavePageContent";
        //public string SavePageToolTip { get; set; } = "ViewName_SavePageContentToolTip";

        // Put these in Resource File
        //    <system:String x:Key="ViewName_SavePageContent">SavePage</system:String>
        //    <system:String x:Key="ViewName_SavePageContentToolTip">SavePage ToolTip</system:String>  

        public void OnSavePageExecute()
        {
            Log.TRACE("Enter", Common.LOG_CATEGORY);
            // Wrap a big, OMG, what have I done ???, undo around the whole thing !!!

            int undoScope = Common.VisioApplication.BeginUndoScope("OnSavePageExecute");
            Message = "Cool, you called SavePage";
            //Common.EventAggregator.GetEvent<SavePageEvent>().Publish();

            MSVisio.Application app = Common.VisioApplication;

            MSVisio.Page newPage = app.ActiveDocument.Pages.Add();
            newPage.NameU = NewPageName;
            newPage.Index = (short)(CurrentPageIndex + 1);

            // Start Cut Four

            // Put this in places that listen for event
            // Common.EventAggregator.GetEvent<SavePageEvent>().Subscribe(SavePage);

            // End Cut Four

            Common.VisioApplication.EndUndoScope(undoScope, true);
            Log.TRACE("Exit", Common.LOG_CATEGORY);
        }

        public bool OnSavePageCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion

        #endregion Commands

    }
}
