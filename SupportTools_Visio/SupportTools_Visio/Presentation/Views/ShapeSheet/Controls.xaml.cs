using System.Windows.Controls;
using SupportTools_Visio.Presentation.ViewModels;
using VNC;

namespace SupportTools_Visio.Presentation.Views
{
    public partial class Controls : UserControl
    {
        #region Constructors and Load

        public Controls()
        {
            Log.Trace("Enter", Common.LOG_CATEGORY);
            InitializeComponent();
            Log.Trace("Exit", Common.LOG_CATEGORY);
        }

        #endregion
    }
}
