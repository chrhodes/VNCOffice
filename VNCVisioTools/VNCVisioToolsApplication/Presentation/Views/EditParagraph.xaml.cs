using System.Windows.Controls;
using DevExpress.Xpf.Editors;
using VNCVisioToolsApplication.Presentation.ViewModels;

using VNC;

namespace VNCVisioToolsApplication.Presentation.Views
{
    public partial class EditParagraph : UserControl
    {
        private readonly EditParagraphViewModel _viewModel;

        #region Constructors and Load

        public EditParagraph(EditParagraphViewModel viewModel)
        {
            Log.Trace("Enter", Common.LOG_CATEGORY);
            InitializeComponent();
            _viewModel = viewModel;
            DataContext = _viewModel;
            Log.Trace("Exit", Common.LOG_CATEGORY);
        }

        #endregion

        private void cbeHorizonatlAlignment_SelectedIndexChanged(object sender, System.Windows.RoutedEventArgs e)
        {
            ComboBoxEdit cbe = (ComboBoxEdit)sender;

            var selectedItem = cbe.SelectedItem;
            var selectedText = cbe.SelectedText;
            var selectedIndex = cbe.SelectedIndex;

            ItemInfo itemInfo = (ItemInfo)selectedItem;

            _viewModel.Paragraph.HAlign = itemInfo.Value;
        }

        // TODO(crhodes)
        // How can we make this go away.  Put events raised from Controls (not buttons which support Commands)

        //private void teHAlign2_EditValueChanged(object sender, EditValueChangedEventArgs e)
        //{
        //    var eThing = e;

        //    TextEdit te = (TextEdit)sender;

        //    _viewModel.Paragraph.HAlign = te.Text;

        //    _viewModel.Paragraph.HAlign = (string)e.NewValue;
        //}
    }
}
