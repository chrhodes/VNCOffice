using System;
using System.Windows.Controls;

using VNC;

namespace VNCShapeSheetApplication.Presentation.Views
{
    public partial class PageLayout : UserControl
    {
        #region Constructors and Load

        public PageLayout()
        {
            Int64 startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            InitializeComponent();

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion
    }
}
