using System.Windows.Controls;

using SupportTools_Excel.Core.Presentation.ViewModels;

using VNC;
using VNC.Core.Mvvm;

namespace SupportTools_Excel.AzureDevOpsExplorer.Presentation.Views
{
    public partial class WorkItems : UserControl, IView
    {

        #region Constructors and Load

        // ViewModel First

        public WorkItems()
        {
            long startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            InitializeComponent();

            lgMain.IsCollapsed = true;

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }


        // View First
        // Can also declare this in Xaml

        public WorkItems(IAZDOWorkItemsViewModel viewModel)
        {
            long startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            InitializeComponent();

            ViewModel = viewModel;

            InitializeView();

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void InitializeView()
        {
            long startTicks = Log.VIEW("Enter", Common.LOG_CATEGORY);

            // TODO(crhodes)
            // Perform any initialization or configuration of View

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
