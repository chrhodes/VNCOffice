using System;
using System.Windows;
using System.Windows.Input;
using Prism.Commands;
using VNCVisioToolsApplication.Presentation.ModelWrappers;
using VNC;
using VNC.Core.Mvvm;

using VNCVisioAddIn = VNC.Visio.VSTOAddIn;

namespace VNCVisioToolsApplication.Presentation.ViewModels
{
    public class EditTextViewModel : ViewModelBase, IEditTextViewModel, IInstanceCountVM
    {
        #region Constructors, Initialization, and Load

        public EditTextViewModel()
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.Constructor) startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            InstanceCountVM++;

            // TODO(crhodes)
            // Save constructor parameters here

            InitializeViewModel();

            if (Common.VNCLogging.Constructor) Log.CONSTRUCTOR($"Exit VM:{InstanceCountVM}", Common.LOG_CATEGORY, startTicks);
        }

        private void InitializeViewModel()
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.ViewModelLow) startTicks = Log.VIEWMODEL_LOW("Enter", Common.LOG_CATEGORY);

            // NOTE(crhodes)
            // Put things here that initialize the ViewModel
            // Initialize EventHandlers, Commands, etc.

            DoSomething = new DelegateCommand(OnDoSomethingExecute, OnDoSomethingCanExecute);
            DoSomethingElse = new DelegateCommand<string>(OnDoSomethingElseExecute, OnDoSomethingElseCanExecute);
            //DoSomething = new DelegateCommand(OnDoSomethingExecute);
            TextBlockFormat = new VNCVisioAddIn.Presentation.ModelWrappers.TextBlockFormatWrapper(new VNCVisioAddIn.Domain.TextBlockFormatRow());


            if (Common.VNCLogging.ViewModelLow) Log.VIEWMODEL_LOW("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region Enums (None)


        #endregion

        #region Structures (None)


        #endregion

        #region Fields and Properties

        public VNCVisioAddIn.Presentation.ModelWrappers.TextBlockFormatWrapper TextBlockFormat
        {
            get => textBlockFormat;
            set => textBlockFormat = value;
        }

        public string Message { get; set; } = "Dr. Natalie Rene Rhodes";

        private VNCVisioAddIn.Presentation.ModelWrappers.TextBlockFormatWrapper textBlockFormat;

        #endregion

        #region Event Handlers (None)


        #endregion

        #region Commands (None)

        #region DoSomething

        public DelegateCommand DoSomething { get; private set; }

        public void OnDoSomethingExecute()
        {
            Log.TRACE("Enter", Common.LOG_CATEGORY);
            // Wrap a big, OMG, what have I done ???, undo around the whole thing !!!

            int undoScope = Common.VisioApplication.BeginUndoScope("AddConnectionPoints");

            // Just need to pass in the model.

            Actions.Visio_Shape.UpdateTextSections(TextBlockFormat.Model);

            Common.VisioApplication.EndUndoScope(undoScope, true);
            Log.TRACE("Exit", Common.LOG_CATEGORY);
        }

        public Boolean OnDoSomethingCanExecute()
        {
            return true;
        }

        #endregion

        #region DoSomethingElse

        public DelegateCommand<string> DoSomethingElse { get; private set; }

        void OnDoSomethingElseExecute(string value)
        {
            switch (value)
            {
                case "0 pt":
                    SetMargins(value);
                    break;

                case "1 pt":
                    SetMargins(value);
                    break;

                case "2 pt":
                    SetMargins(value);
                    break;

                default:
                    MessageBox.Show($"Unknown tag: {value}");
                    break;
            }
        }

        bool OnDoSomethingElseCanExecute(string arg)
        {
            return true;
        }

        #endregion

        #endregion

        #region Public Methods (None)


        #endregion

        #region Protected Methods (None)


        #endregion

        #region Private Methods

        private void SetMargins(string tag)
        {
            TextBlockFormat.TopMargin = tag;
            TextBlockFormat.BottomMargin = tag;
            TextBlockFormat.LeftMargin = tag;
            TextBlockFormat.RightMargin = tag;
        }

        #endregion

        #region IInstanceCountVM

        private static int _instanceCountVM;

        public int InstanceCountVM
        {
            get => _instanceCountVM;
            set => _instanceCountVM = value;
        }

        #endregion
    }
}
