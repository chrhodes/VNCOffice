using System;
using System.Windows.Controls;

using SupportTools_Visio.Presentation.ViewModels;

using VNC;

namespace SupportTools_Visio.Presentation.Views
{
    public partial class ShapeSheetSection : UserControl
    {
        private readonly ShapeSheetSectionBase _viewModel;

        public ShapeSheetSection(ShapeSheetSectionBase viewModel, ContentControl ssUserControl)
        {
            Int64 startTicks = Log.CONSTRUCTOR($"Enter viewModel:({viewModel.GetType()} userControl:({ssUserControl.GetType()}))", Common.LOG_CATEGORY);

            InitializeComponent();
            _viewModel = viewModel;
            DataContext = _viewModel;
            ssSectionUserControl.Content = ssUserControl;

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }
    }
}
