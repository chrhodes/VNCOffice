using System;

using Prism.Events;
using Prism.Services.Dialogs;

using VNC;
using VNC.Core.Mvvm;

namespace DemoAndTestLogging2.Presentation.ViewModels
{
    public class ViewABCDViewModel : EventViewModelBase, IViewABCDViewModel, IInstanceCountVM
    {
        #region Constructors, Initialization, and Load

        public ViewABCDViewModel(
            IEventAggregator eventAggregator,
            IDialogService dialogService) : base(eventAggregator, dialogService)
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.Constructor) startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            // TODO(crhodes)
            // Save constructor parameters here

            InstanceCountVM++;

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

            Message = "ViewABCDViewModel says hello";
            MessageA = "ViewABCDViewModel says hello A";
            MessageB = "ViewABCDViewModel says hello B";
            MessageC = "ViewABCDViewModel says hello C";
            MessageD = "ViewABCDViewModel says hello D";

            if (Common.VNCLogging.ViewModelLow) Log.VIEWMODEL_LOW("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region Enums (none)


        #endregion

        #region Structures (none)


        #endregion

        #region Fields and Properties

        private string _messageA;

        public string MessageA
        {
            get => _messageA;
            set
            {
                if (_messageA == value)
                    return;
                _messageA = value;
                OnPropertyChanged();
            }
        }

        private string _messageB;

        public string MessageB
        {
            get => _messageB;
            set
            {
                if (_messageB == value)
                    return;
                _messageB = value;
                OnPropertyChanged();
            }
        }

        private string _messageC;

        public string MessageC
        {
            get => _messageC;
            set
            {
                if (_messageC == value)
                    return;
                _messageC = value;
                OnPropertyChanged();
            }
        }

        private string _messageD;

        public string MessageD
        {
            get => _messageD;
            set
            {
                if (_messageD == value)
                    return;
                _messageD = value;
                OnPropertyChanged();
            }
        }

        #endregion

        #region Event Handlers (none)



        #endregion

        #region Commands (none)



        #endregion

        #region Public Methods (none)



        #endregion

        #region Protected Methods (none)



        #endregion

        #region Private Methods (none)



        #endregion

        #region IInstanceCount

        private static int _instanceCountVM;

        public int InstanceCountVM
        {
            get => _instanceCountVM;
            set => _instanceCountVM = value;
        }

        #endregion
    }
}
