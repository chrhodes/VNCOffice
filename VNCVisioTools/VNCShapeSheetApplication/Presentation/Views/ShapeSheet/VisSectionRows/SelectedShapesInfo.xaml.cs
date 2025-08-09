using System;
using System.Windows.Controls;

using VNC;

namespace VNCShapeSheetApplication.Presentation.Views
{
    /// <summary>
    /// Interaction logic for SelectedShapesInfo.xaml
    /// </summary>
    public partial class SelectedShapesInfo : UserControl
    {
        public SelectedShapesInfo()
        {
            Int64 startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            InitializeComponent();

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }
    }
}
