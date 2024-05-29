using System.Windows.Controls;
using DevExpress.XtraRichEdit.Model;
using SupportTools_Visio.Presentation.ViewModels;
using VNC;

namespace SupportTools_Visio.Presentation.Views
{
    public partial class ConnectionPoints : UserControl
    {
        #region Constructors and Load

        public ConnectionPoints()
        {
            Log.Trace("Enter", Common.LOG_CATEGORY);
            InitializeComponent();
            Log.Trace("Exit", Common.LOG_CATEGORY);
        }

        #endregion
    }
}
