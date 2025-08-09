using System;
using System.Windows.Controls;
using VNCShapeSheetApplication.Presentation.ViewModels;
using VNC;

namespace VNCShapeSheetApplication.Presentation.Views
{
    public partial class GeometryRow : UserControl
    {
        #region Constructors and Load

        public GeometryRow()
        {
            Int64 startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            InitializeComponent();

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion
    }
}
