using System.Windows.Controls;
using SupportTools_Visio.Presentation.ViewModels;
using VNC;

namespace SupportTools_Visio.Presentation.Views
{
    public partial class Character : UserControl
    {
        #region Constructors and Load

        public Character()
        {
            Log.Trace("Enter", Common.LOG_CATEGORY);
            InitializeComponent();
            //_viewModel = viewModel;
            //DataContext = _viewModel;
            Log.Trace("Exit", Common.LOG_CATEGORY);
        }

        #endregion
    }
}
