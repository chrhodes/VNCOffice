using System;
using System.Windows;
using System.Windows.Controls;
using SupportTools_Excel.ActiveDirectoryExplorer.Presentation.ViewModels;
using SupportTools_Excel.Core.Presentation.ViewModels;
using SupportTools_Excel.Presentation.ViewModels;

using VNC;
using VNC.Core.Mvvm;

namespace SupportTools_Excel.ActiveDirectoryExplorer.Presentation.Views
{
    public partial class ActiveDirectoryExplorer : UserControl, IView
    {
        #region Constructors and Load

        // View First.  
        // View is passed ViewModel through Injection
        // or can declare ViewModel in Xaml or Code

        // ViewModel First.  
        // ViewModel creates View 
        // and sets DataContext by setting ViewModel property

        public ActiveDirectoryExplorer()
        {
            long startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            InitializeComponent();

            // If View First with ViewModel in Xaml
            // Expose ViewModel
            // ViewModel = (I$customTYPE$ViewModel)DataContext;

            // Can create directly
            ViewModel = new ActiveDirectoryExplorerViewModel();
            ViewModel.View = this;

            InitializeView();

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public ActiveDirectoryExplorer(IADMainViewModel viewModel)
        {
            long startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            InitializeComponent();
            ViewModel = viewModel;

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void InitializeView()
        {
            long startTicks = Log.VIEW("Enter", Common.LOG_CATEGORY);

            // TODO(crhodes)
            // Perform any initialization or configuration of View
            LoadControlContents();
            //lgMain.IsCollapsed = true;

            Log.VIEW("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region Properties

        private IViewModel _viewModel;

        public IViewModel ViewModel
        {
            get { return _viewModel; }

            set
            {
                _viewModel = value;
                DataContext = _viewModel;
            }
        }

        #endregion

        private void LoadControlContents()
        {
            long startTicks = Log.VIEW("Enter", Common.LOG_CATEGORY);

            try
            {
                wucActiveDirectory_Picker.PopulateControlFromFile(Common.cCONFIG_FILE);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            Log.VIEW("Exit", Common.LOG_CATEGORY, startTicks);
        }
    }
}
