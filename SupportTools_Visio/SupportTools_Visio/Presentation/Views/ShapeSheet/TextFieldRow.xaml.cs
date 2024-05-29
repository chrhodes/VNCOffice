using System.Windows.Controls;
using SupportTools_Visio.Presentation.ViewModels;
using VNC;

namespace SupportTools_Visio.Presentation.Views
{
    public partial class TextFieldRow : UserControl
    {
        private readonly TextFieldRowViewModel _viewModel;

        #region Constructors and Load

        public TextFieldRow(TextFieldRowViewModel viewModel)
        {
            Log.Trace("Enter", Common.LOG_CATEGORY);
            InitializeComponent();
            _viewModel = viewModel;
            DataContext = _viewModel;
            Log.Trace("Exit", Common.LOG_CATEGORY);
        }

        #endregion
    }
}
