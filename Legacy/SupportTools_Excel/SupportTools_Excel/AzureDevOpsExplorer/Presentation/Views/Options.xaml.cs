using System.Windows.Controls;

using SupportTools_Excel.Core.Presentation.ViewModels;

using VNC;
using VNC.Core.Mvvm;

namespace SupportTools_Excel.AzureDevOpsExplorer.Presentation.Views
{
    public partial class Options : UserControl, IView
    {
        #region Constructors and Load

        // ViewModel First.  ViewModel creates View 
        // and sets DataContext by setting ViewModel property

        public Options()
        {
            long startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            InitializeComponent();

            // If View First with ViewModel in Xaml
            // Expose ViewModel
            // ViewModel = (ICatViewModel)DataContext;

            // Can create directly
            ViewModel = (IAZDOOptionsViewModel)DataContext;
            // If ViewModel needs access to view (or view's ViewModel)
            // wire them up.
            ViewModel.View = this;

            InitializeView();

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // View First.  View is passed ViewModel through Injection
        // or can declare ViewModel in Xaml

        public Options(IAZDOOptionsViewModel viewModel)
        {
            long startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            InitializeComponent();

            ViewModel = viewModel;
            ViewModel.View = this;

            InitializeView();

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void InitializeView()
        {
            long startTicks = Log.VIEW("Enter", Common.LOG_CATEGORY);
            // TODO(crhodes)
            // Perform any initialization or configuration of View

            //lgMain.IsCollapsed = true;

            debugOptions.IsCollapsed = true;
            dateRange.IsCollapsed = true;
            workItemOptions.IsCollapsed = true;
            loopingDelays.IsCollapsed = true;
            miscOptions.IsCollapsed = true;
            excelOutputOptions.IsCollapsed = true;

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

    }
}
