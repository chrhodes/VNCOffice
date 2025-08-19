using System.Windows.Controls;

using VNC;
using VNC.Core.Mvvm;
using VNCVisioToolsApplication.Core;
using VNCVisioToolsApplication.Presentation.ViewModels;

namespace VNCVisioToolsApplication.Presentation.Views
{
    public partial class DuplicatePage : UserControl, IView
    {
        #region Constructors and Load

        // View First.  
        // View is passed ViewModel through Injection
        // or can declare ViewModel in Xaml or Code

        // ViewModel First.  
        // ViewModel creates View 
        // and sets DataContext by setting ViewModel property

        public DuplicatePage()
        {
            long startTicks = Log.TRACE("Enter", Common.LOG_CATEGORY);

            InitializeComponent();

            // If View First with ViewModel in Xaml
            // Expose ViewModel
            // ViewModel = (IDuplicatePageViewModel)DataContext;

            // Can create directly
            // ViewModel = new DuplicatePageViewModel();

            InitializeView();

            Log.TRACE("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public DuplicatePage(IDuplicatePageViewModel viewModel)
        {
            long startTicks = Log.TRACE("Enter", Common.LOG_CATEGORY);

            InitializeComponent();

            ViewModel = viewModel;

            InitializeView();

            Log.TRACE("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void InitializeView()
        {
            // TODO(crhodes)
            // Perform any initialization or configuration of View

            //lgMain.IsCollapsed = true;
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
