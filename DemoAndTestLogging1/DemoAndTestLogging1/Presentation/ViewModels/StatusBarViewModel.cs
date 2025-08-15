using System;

using Prism.Commands;
using Prism.Events;
using Prism.Services.Dialogs;

using VNC;
using VNC.Core.Events;
using VNC.Core.Mvvm;
using VNC.Core.Presentation;
using VNC.WPF.Presentation.Dx.Views;
using VNC.WPF.Presentation.Views;

namespace DemoAndTestLogging1.Presentation.ViewModels
{
    public class StatusBarViewModel : EventViewModelBase, IStatusBarViewModel, IInstanceCountVM
    {
        #region Constructors, Initialization, and Load

        public StatusBarViewModel(
            IEventAggregator eventAggregator,
            DialogService dialogService) : base(eventAggregator, dialogService)
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

            LoggingConfigurationCommand = new DelegateCommand(LoggingConfiguration, LoggingConfigurationCanExecute);

            EventAggregator.GetEvent<StatusMessageEvent>().Subscribe(UpdateStatusMessage);

            if (Common.VNCLogging.ViewModelLow) Log.VIEWMODEL_LOW("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region Enums (none)


        #endregion

        #region Structures (none)


        #endregion

        #region Fields and Properties (none)



        #endregion

        #region Event Handlers

        private void UpdateStatusMessage(string message)
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.EventHandler) startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Message = message;

            if (Common.VNCLogging.EventHandler) Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region Commands

        #region LoggingConfiguration Command

        public DelegateCommand LoggingConfigurationCommand { get; set; }

        public static WindowHost _loggingConfigurationHost = null;

        //public TYPE LoggingConfigurationCommandParameter;
        public string LoggingConfigurationContent { get; set; } = "LoggingConfiguration";
        public string? LoggingConfigurationToolTip { get; set; } = "LoggingConfiguration ToolTip";

        // Can get fancy and use Resources
        //public string LoggingConfigurationContent { get; set; } = "ViewName_LoggingConfigurationContent";
        //public string LoggingConfigurationToolTip { get; set; } = "ViewName_LoggingConfigurationContentToolTip";

        // Put these in Resource File
        //    <system:String x:Key="ViewName_LoggingConfigurationContent">LoggingConfiguration</system:String>
        //    <system:String x:Key="ViewName_LoggingConfigurationContentToolTip">LoggingConfiguration ToolTip</system:String>

        // If using CommandParameter, figure out TYPE and fix above
        //public void LoggingConfiguration(TYPE value)
        public void LoggingConfiguration()
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.EventHandler) startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Message = "Cool, you called LoggingConfiguration";

            if (_loggingConfigurationHost is null) _loggingConfigurationHost = new WindowHost(EventAggregator);

            var userControl = new VNCLoggingConfigMain();

            _loggingConfigurationHost.DisplayUserControlInHost(
                "VNC Logging Configuration",
                //Common.DEFAULT_WINDOW_WIDTH,
                //Common.DEFAULT_WINDOW_HEIGHT,
                (Int32)userControl.Width + Common.WINDOW_HOSTING_USER_CONTROL_WIDTH_PAD,
                (Int32)userControl.Height + Common.WINDOW_HOSTING_USER_CONTROL_HEIGHT_PAD,
                ShowWindowMode.Modeless_Show,
                userControl);

            // Uncomment this if you are telling someone else to handle this

            // Common.EventAggregator.GetEvent<LoggingConfigurationEvent>().Publish();

            // May want EventArgs

            //  EventAggregator.GetEvent<LoggingConfigurationEvent>().Publish(
            //      new LoggingConfigurationEventArgs()
            //      {
            //            Organization = _collectionMainViewModel.SelectedCollection.Organization,
            //            Process = _contextMainViewModel.Context.SelectedProcess
            //      });

            // Start Cut Four - Put this in PrismEvents

            // public class LoggingConfigurationEvent : PubSubEvent { }

            // End Cut Four

            // Start Cut Five - Put this in places that listen for event

            //Common.EventAggregator.GetEvent<LoggingConfigurationEvent>().Subscribe(LoggingConfiguration);

            // End Cut Five

            if (Common.VNCLogging.EventHandler) Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // If using CommandParameter, figure out TYPE and fix above
        //public bool LoggingConfigurationCanExecute(TYPE value)
        public bool LoggingConfigurationCanExecute()
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
