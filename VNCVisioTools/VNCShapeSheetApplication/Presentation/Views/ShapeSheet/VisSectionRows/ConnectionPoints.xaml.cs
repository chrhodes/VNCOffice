using System;
using System.Windows.Controls;
using DevExpress.XtraRichEdit.Model;
using VNCShapeSheetApplication.Presentation.ViewModels;
using VNC;

namespace VNCShapeSheetApplication.Presentation.Views
{
    public partial class ConnectionPoints : UserControl
    {
        #region Constructors and Load

        public ConnectionPoints()
        {
            Int64 startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            InitializeComponent();

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion
    }
}
