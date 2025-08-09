using System;
using System.Windows;
using System.Windows.Controls;

using Prism.Commands;
using DevExpress.Mvvm.POCO;
using SupportTools_Visio.Presentation.ModelWrappers;
using SupportTools_Visio.Presentation.ViewModels;

using VNC;
//using System.Windows.Forms;

namespace SupportTools_Visio.Presentation.Views
{
    public partial class EditText : UserControl
    {
        private readonly EditTextViewModel _viewModel;

        #region Constructors and Load

        public EditText(EditTextViewModel viewModel)
        {
            Log.Trace("Enter", Common.LOG_CATEGORY);
            InitializeComponent();
            _viewModel = viewModel;
            DataContext = _viewModel;
            //LoadControlContents();
            Log.Trace("Exit", Common.LOG_CATEGORY);
        }

        //private void UserControl_Loaded(object sender, RoutedEventArgs e)
        //{
        //    Log.Trace("Enter", Common.LOG_CATEGORY);
        //    //VNC.Log.Trace("", Common.LOG_APPNAME, 0);
        //    //VNCVisioAddIn.Common.DisplayInDebugWindow(string.Format("{0}()",
        //    //    System.Reflection.MethodInfo.GetCurrentMethod().Name));
        //    Log.Trace("Exit", Common.LOG_CATEGORY);
        //}

        //private void UserControl_Unloaded(object sender, RoutedEventArgs e)
        //{
        //    Log.Trace("Enter", Common.LOG_CATEGORY);
        //    //VNC.Log.Trace("", Common.LOG_APPNAME, 0);
        //    //VNCVisioAddIn.Common.DisplayInDebugWindow(string.Format("{0}()",
        //    //    System.Reflection.MethodInfo.GetCurrentMethod().Name));
        //    Log.Trace("Exit", Common.LOG_CATEGORY);
        //}

        #endregion

        //private TextBlockFormat textBlockFormat = new TextBlockFormat();
        //private TextBlockFormatWrapper textBlockFormat; 

        #region Event Handlers


        private void btnInitializeMargins_Click(object sender, RoutedEventArgs e)
        {
            //if (textBlockFormat is null)
            //{
            //    textBlockFormat = new TextBlockFormat();
            //}

            string tag = ((Button)sender).Tag.ToString();

            switch (tag)
            {
                case "0 pt":
                    SetMargins(tag);
                    break;

                case "1 pt":
                    SetMargins(tag);
                    break;

                case "2 pt":
                    SetMargins(tag);
                    break;

                default:
                    MessageBox.Show($"Unknown tag: {tag}");
                    break;
            }
        }

        #endregion

        #region Private Methods

        //private void LoadControlContents()
        //{
        //    try
        //    {
        //        //textBlockFormat = new TextBlockFormatWrapper(new Domain.TextBlockFormat());
        //        //DataContext = textBlockFormat;

        //        //layoutControl.DataContext = textBlockFormat;
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.ToString());
        //    }
        //}

        void SetMargins(string tag)
        {
            _viewModel.TextBlockFormat.TopMargin = tag;
            _viewModel.TextBlockFormat.BottomMargin = tag;
            _viewModel.TextBlockFormat.LeftMargin = tag;
            _viewModel.TextBlockFormat.RightMargin = tag;
        }

        #endregion

        //private void btnDoSomething_Click(object sender, RoutedEventArgs e)
        //{
        //    Log.Trace("Enter", Common.LOG_CATEGORY);
        //    //VNC.Log.Trace("", Common.LOG_APPNAME, 0);
        //    // Wrap a big, OMG, what have I done ???, undo around the whole thing !!!

        //    int undoScope = Globals.ThisAddIn.Application.BeginUndoScope("AddConnectionPoints");

        //    // Just need to pass in the model.

        //    Actions.Visio_Shape.UpdateTextSections(_viewModel.TextBlockFormat.Model);

        //    Globals.ThisAddIn.Application.EndUndoScope(undoScope, true);
        //}
    }
}
