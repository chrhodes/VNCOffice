using System;
using System.Windows.Controls;

using VNCShapeSheetApplication.Presentation.ViewModels;

using VNC;

namespace VNCShapeSheetApplication.Presentation.Views
{
    public partial class ShapeSheetObjectSection : UserControl
    {
        private readonly ShapeSheetSectionBase _viewModel;

        public ShapeSheetObjectSection(ShapeSheetSectionBase viewModel, ContentControl ssUserControl)
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
