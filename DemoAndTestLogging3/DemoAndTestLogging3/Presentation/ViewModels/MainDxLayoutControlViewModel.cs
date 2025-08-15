using System;

using Prism.Commands;
using Prism.Events;
using Prism.Services.Dialogs;

using VNC;
using VNC.Core.Mvvm;

namespace DemoAndTestLogging3.Presentation.ViewModels
{
    public class MainDxLayoutControlViewModel : EventViewModelBase, IInstanceCountVM
    {
        #region Constructors, Initialization, and Load

        public MainDxLayoutControlViewModel(
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

            Title = "DemoAndTestLogging3 - MainDxLayout";

            DeveloperModeCommand = new DelegateCommand(DeveloperMode, DeveloperModeCanExecute);

            // If using CommandParameter, figure out TYPE here and below
            // and remove above declaration
            //DeveloperModeCommand = new DelegateCommand<TYPE>(DeveloperMode, DeveloperModeCanExecute);

            if (Common.VNCLogging.ViewModelLow) Log.VIEWMODEL_LOW("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region Enums (none)



        #endregion

        #region Structures (none)



        #endregion

        #region Fields and Properties

        // private string _title = "DemoAndTestLogging3 - MainDxLayout";

        // public string Title
        // {
        // get => _title;
        // set
        // {
        // if (_title == value)
        // return;
        // _title = value;
        // OnPropertyChanged();
        // }
        // }



        #endregion

        #region Event Handlers (none)


        #endregion

        #region Commands

        #region DeveloperMode Command

        public DelegateCommand DeveloperModeCommand { get; set; }
        // If using CommandParameter, figure out TYPE here and above
        // and remove above declaration
        //public DelegateCommand<TYPE> DeveloperModeCommand { get; set; }

        //public TYPE DeveloperModeCommandParameter;
        public string DeveloperModeContent { get; set; } = "DeveloperMode";
        public string? DeveloperModeToolTip { get; set; } = "DeveloperMode ToolTip";

        // Can get fancy and use Resources
        //public string DeveloperModeContent { get; set; } = "ViewName_DeveloperModeContent";
        //public string DeveloperModeToolTip { get; set; } = "ViewName_DeveloperModeContentToolTip";

        // Put these in Resource File
        //    <system:String x:Key="ViewName_DeveloperModeContent">DeveloperMode</system:String>
        //    <system:String x:Key="ViewName_DeveloperModeContentToolTip">DeveloperMode ToolTip</system:String>

        // If using CommandParameter, figure out TYPE and fix above
        //public void DeveloperMode(TYPE value)
        public void DeveloperMode()
        {
            Int64 startTicks = Log.EVENT("Enter", Common.LOG_CATEGORY);
            // TODO(crhodes)
            // Do something amazing.
            Message = "Cool, you called DeveloperMode";

            PublishStatusMessage(Message);

            if (Common.DeveloperMode)
            {
                //if (Common.CurrentRibbonShell is not null) Common.CurrentRibbonShell.DeveloperUIMode = System.Windows.Visibility.Collapsed;
                if (Common.CurrentShell is not null) Common.CurrentShell.DeveloperUIMode = System.Windows.Visibility.Collapsed;
            }
            else
            {
                //if (Common.CurrentRibbonShell is not null) Common.CurrentRibbonShell.DeveloperUIMode = System.Windows.Visibility.Visible;
                if (Common.CurrentShell is not null) Common.CurrentShell.DeveloperUIMode = System.Windows.Visibility.Visible;
            }

            Common.DeveloperMode = !Common.DeveloperMode;

            PublishDeveloperMode(Common.DeveloperMode);

            // Uncomment this if you are telling someone else to handle this

            // Common.EventAggregator.GetEvent<DeveloperModeEvent>().Publish();

            // May want EventArgs

            //  EventAggregator.GetEvent<DeveloperModeEvent>().Publish(
            //      new DeveloperModeEventArgs()
            //      {
            //            Organization = _collectionMainViewModel.SelectedCollection.Organization,
            //            Process = _contextMainViewModel.Context.SelectedProcess
            //      });

            // Start Cut Four - Put this in PrismEvents

            // public class DeveloperModeEvent : PubSubEvent { }

            // End Cut Four

            // Start Cut Five - Put this in places that listen for event

            //Common.EventAggregator.GetEvent<DeveloperModeEvent>().Subscribe(DeveloperMode);

            // End Cut Five

            Log.EVENT("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // If using CommandParameter, figure out TYPE and fix above
        //public bool DeveloperModeCanExecute(TYPE value)
        public bool DeveloperModeCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion

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
