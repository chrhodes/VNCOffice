using System;
using System.Windows;

using dxe = DevExpress.Xpf.Editors;

using VNC;
using VNC.Core.Mvvm;

namespace SupportTools_Visio.Presentation.Views
{
    public partial class MovePage : ViewBase, IInstanceCountV
    {

        public MovePage()
        {
            Int64 startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            InstanceCountV++;
            InitializeComponent();

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public MovePage(ViewModels.IMovePageViewModel viewModel)
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

        //private void btnExecuteCommand_Click(object sender, RoutedEventArgs e)
        //{
        //    Log.Trace("Enter", Common.LOG_CATEGORY, 0);
        //    // Wrap a big, OMG, what have I done ???, undo around the whole thing !!!

        //    int undoScope = Globals.ThisAddIn.Application.BeginUndoScope("ParseCommand");

        //    // TODO(crhodes)
        //    // Get this from UI

        //    string targetDocument = (string)cbeOpenDocuments.SelectedItemValue;
        //    SupportTools_Visio.Actions.Visio_Document.MovePages(targetDocument);

        //    Globals.ThisAddIn.Application.EndUndoScope(undoScope, true);

        //    Log.Trace("Exit", Common.LOG_CATEGORY);
        //}

        //private void cbeDefaultPatterns_SelectedIndexChanged(object sender, RoutedEventArgs e)
        //{
        //    dxe.ComboBoxEdit control = (dxe.ComboBoxEdit)sender;

        //    dxe.ComboBoxEditItem item = (dxe.ComboBoxEditItem)control.SelectedItem;
        //}

        //private void cbeOpenDocuments_SelectedIndexChanged(object sender, RoutedEventArgs e)
        //{

        //}
    }
}
