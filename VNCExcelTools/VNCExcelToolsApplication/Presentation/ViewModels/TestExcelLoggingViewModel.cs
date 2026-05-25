using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Input;

using Prism.Commands;
using Prism.Dialogs;
using Prism.Events;

using VNC;
using VNC.Core.Events;
using VNC.Core.Mvvm;
using VNC.Core.Services;

namespace VNCExcelToolsApplication.Presentation.ViewModels
{
    public class TestExcelLoggingViewModel : EventViewModelBase, ITestExcelLoggingViewModel, IInstanceCountVM
    {
        #region Fields (none)



        #endregion

        #region Constructors, Initialization, and Load

        public TestExcelLoggingViewModel(
            IEventAggregator eventAggregator,
            IDialogService dialogService) : base(eventAggregator, dialogService)
        {
            Int64 startTicks = 0;
            if (Common.VNCLogLevel.Constructor) startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            InstanceCountVM++;

            // TODO(crhodes)
            // Save constructor parameters here

            InitializeViewModel();

            if (Common.VNCLogLevel.Constructor) Log.CONSTRUCTOR($"Exit VM:{InstanceCountVM}", Common.LOG_CATEGORY, startTicks);
        }

        private void InitializeViewModel()
        {
            Int64 startTicks = 0;
            if (Common.VNCLogLevel.ViewModelLow) startTicks = Log.VIEWMODEL_LOW("Enter", Common.LOG_CATEGORY);

            // NOTE(crhodes)
            // Put things here that initialize the ViewModel
            // Initialize EventHandlers, Commands, etc.

            SayHelloCommand = new DelegateCommand(
                SayHello, SayHelloCanExecute);

            Message = "TestExcelLoggingViewModel says hello";

            if (Common.VNCLogLevel.ViewModelLow) Log.VIEWMODEL_LOW("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region Enums (none)



        #endregion

        #region Structures (none)



        #endregion

        #region Properties (none)



        #endregion

        #region Event Handlers (none)



        #endregion

        #region Public Methods (none)



        #endregion

        #region Commands

        #region SayHello Command

        public ICommand SayHelloCommand { get; private set; }

        private void SayHello()
        {
            Int64 startTicks = 0;
            if (Common.VNCLogLevel.EventHandler) startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Message = "Hello";

            if (Common.VNCLogLevel.EventHandler) Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private bool SayHelloCanExecute()
        {
            return true;
        }

        #endregion

        #endregion

        #region Protected Methods (none)



        #endregion

        #region Private Methods (none)



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
