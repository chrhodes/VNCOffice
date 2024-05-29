using System;
using System.Windows;

using dxe = DevExpress.Xpf.Editors;

using VNC;
using VNC.Core.Mvvm;

namespace SupportTools_Visio.Presentation.Views
{
    public partial class RenamePage : ViewBase, IInstanceCountV
    {
        #region Constructors and Load

        public RenamePage()
        {
            Int64 startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            InstanceCountV++;
            InitializeComponent();

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public RenamePage(ViewModels.IRenamePageViewModel viewModel)
        {
            Int64 startTicks = Log.CONSTRUCTOR($"Enter viewModel({viewModel.GetType()}", Common.LOG_CATEGORY);

            InstanceCountV++;
            InitializeComponent();

            ViewModel = viewModel;

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #region IInstanceCount

        private static int _instanceCountV;

        public int InstanceCountV
        {
            get => _instanceCountV;
            set => _instanceCountV = value;
        }

        #endregion

#endregion

        #region Event Handlers

        private void btnExecuteCommand_Click(object sender, RoutedEventArgs e)
        {
            VNC.Log.Trace("", Common.LOG_CATEGORY, 0);
            // Wrap a big, OMG, what have I done ???, undo around the whole thing !!!

            int undoScope = Globals.ThisAddIn.Application.BeginUndoScope("ParseCommand");

            SupportTools_Visio.Actions.Visio_Document.RenamePages(teSearchExpression.Text, teReplacementExpression.Text);

            Globals.ThisAddIn.Application.EndUndoScope(undoScope, true);
        }

        private void cbeDefaultPatterns_SelectedIndexChanged(object sender, RoutedEventArgs e)
        {
            dxe.ComboBoxEdit control = (dxe.ComboBoxEdit)sender;

            dxe.ComboBoxEditItem item = (dxe.ComboBoxEditItem)control.SelectedItem;

            switch (item.Tag)
            {
                case "Front":
                    teSearchExpression.Text = "XXX(.*$)";
                    teReplacementExpression.Text = "YYY$1";
                    break;

                case "Middle":
                    teSearchExpression.Text = "(^.*)XXX(.*$)";
                    teReplacementExpression.Text = "$1YYY$2";
                    break;

                case "End":
                    teSearchExpression.Text = "(^.*)XXX$";
                    teReplacementExpression.Text = "$1YYY";
                    break;

                case "Empty":
                    teSearchExpression.Text = "";
                    teReplacementExpression.Text = "";
                    break;

                default:
                    teSearchExpression.Text = "";
                    teReplacementExpression.Text = "";
                    break;
            }
        }

        #endregion


        #region Private Methods

        private void LoadControlContents()
        {
            try
            {
                //visioCommand_Picker.PopulateControlFromFile(Common.cCONFIG_FILE);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        #endregion
    }
}
