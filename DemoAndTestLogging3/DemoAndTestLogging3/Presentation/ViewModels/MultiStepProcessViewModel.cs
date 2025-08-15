using System;
using System.Windows.Input;

using Prism.Commands;
using Prism.Events;
using Prism.Regions;
using Prism.Services.Dialogs;

using VNC;
using VNC.Core.Mvvm;

namespace DemoAndTestLogging3.Presentation.ViewModels
{
    public class MultiStepProcessViewModel : EventViewModelBase, IMultiStepProcessViewModel, INavigationAware, IInstanceCountVM
    {
        #region Constructors, Initialization, and Load

        private readonly IRegionManager _regionManager;

        public MultiStepProcessViewModel(
            IRegionManager regionManager,
            IEventAggregator eventAggregator,
            IDialogService dialogService) : base(eventAggregator, dialogService)
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.Constructor) startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            // TODO(crhodes)
            // Save constructor parameters here

            _regionManager = regionManager;

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

            SayHelloCommand = new DelegateCommand(
                SayHello, SayHelloCanExecute);

            NavigateStepACommand = new DelegateCommand(NavigateStepA, NavigateStepACanExecute);

            // If using CommandParameter, figure out TYPE here and below
            // and remove above declaration
            //NavigateStepACommand = new DelegateCommand<TYPE>(NavigateStepA, NavigateStepACanExecute)

            Message = "MultiStepProcessViewModel says hello";

            if (Common.VNCLogging.ViewModelLow) Log.VIEWMODEL_LOW("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region Enums (none)



        #endregion

        #region Structures (none)



        #endregion

        #region Fields and Properties (none)



        #endregion

        #region Event Handlers (none)



        #endregion

        #region Public Methods (none)


        #endregion

        #region Commands

        #region NavigateStepA Command

        public DelegateCommand NavigateStepACommand { get; set; }
        // If using CommandParameter, figure out TYPE here and above
        // and remove above declaration
        //public DelegateCommand<TYPE> NavigateStepACommand { get; set; }

        // If displaying UserControl
        // public static WindowHost _NavigateStepAHost = null;

        // If using CommandParameter, figure out TYPE here
        //public TYPE NavigateStepACommandParameter;

        public string NavigateStepAContent { get; set; } = "NavigateStepA";
        public string NavigateStepAToolTip { get; set; } = "NavigateStepA ToolTip";

        // Can get fancy and use Resources
        //public string NavigateStepAContent { get; set; } = "ViewName_NavigateStepAContent";
        //public string NavigateStepAToolTip { get; set; } = "ViewName_NavigateStepAContentToolTip";

        // Put these in Resource File
        //    <system:String x:Key="ViewName_NavigateStepAContent">NavigateStepA</system:String>
        //    <system:String x:Key="ViewName_NavigateStepAContentToolTip">NavigateStepA ToolTip</system:String>

        // If using CommandParameter, figure out TYPE here
        //public void NavigateStepA(TYPE value)
        public void NavigateStepA()
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.EventHandler) startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);
            // TODO(crhodes)
            // Do something amazing.

            Message = "Cool, you called NavigateStepA";

            PublishStatusMessage(Message);

            // TODO(crhodes)
            // Need error handling or some way of checking if things exist
            _regionManager.RequestNavigate("MultiStepProcessView", "uistepa");

            // If launching a UserControl

            // if (_NavigateStepAHost is null) _NavigateStepAnHost = new WindowHost();
            // var userControl = new USERCONTROL();

            // _loggingConfigurationHost.DisplayUserControlInHost(
            //     "TITLE GOES HERE",
            //     //Common.DEFAULT_WINDOW_WIDTH,
            //     //Common.DEFAULT_WINDOW_HEIGHT,
            //     (Int32)userControl.Width + Common.WINDOW_HOSTING_USER_CONTROL_WIDTH_PAD,
            //     (Int32)userControl.Height + Common.WINDOW_HOSTING_USER_CONTROL_HEIGHT_PAD,
            //     ShowWindowMode.Modeless_Show,
            //     userControl);

            // Uncomment this if you are telling someone else to handle this

            // Common.EventAggregator.GetEvent<NavigateStepAEvent>().Publish();

            // May want EventArgs

            //  EventAggregator.GetEvent<NavigateStepAEvent>().Publish(
            //      new NavigateStepAEventArgs()
            //      {
            //            Organization = _collectionMainViewModel.SelectedCollection.Organization,
            //            Process = _contextMainViewModel.Context.SelectedProcess
            //      });

            // Start Cut Four - Put this in PrismEvents

            // public class NavigateStepAEvent : PubSubEvent { }

            // End Cut Four

            // Start Cut Five - Put this in places that listen for event

            //Common.EventAggregator.GetEvent<NavigateStepAEvent>().Subscribe(NavigateStepA);

            // End Cut Five

            if (Common.VNCLogging.EventHandler) Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // If using CommandParameter, figure out TYPE and fix above
        //public bool NavigateStepACanExecute(TYPE value)
        public bool NavigateStepACanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion

        #region SayHello Command

        public ICommand SayHelloCommand { get; private set; }

        private bool SayHelloCanExecute()
        {
            return true;
        }

        private void SayHello()
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.EventHandler) startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Message = "Hello";

            PublishStatusMessage(Message);

            if (Common.VNCLogging.EventHandler) Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion
        #endregion

        #region Protected Methods (none)


        #endregion

        #region Private Methods



        #region INavigationAware

        public bool IsNavigationTarget(NavigationContext navigationContext)
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.EventHandler) startTicks = Log.EVENT_HANDLER($"Enter {navigationContext.Uri}", Common.LOG_CATEGORY);


            if (Common.VNCLogging.EventHandler) Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);

            // Create New Instance
            return false;
            // Use Existing Instance
            // return true;
        }

        public void OnNavigatedTo(NavigationContext navigationContext)
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.EventHandler) startTicks = Log.EVENT_HANDLER($"Enter {navigationContext.Uri}", Common.LOG_CATEGORY);

            var n = navigationContext;

            if (Common.VNCLogging.EventHandler) Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public void OnNavigatedFrom(NavigationContext navigationContext)
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.EventHandler) startTicks = Log.EVENT_HANDLER($"Enter {navigationContext.Uri}", Common.LOG_CATEGORY);

            var n = navigationContext;

            if (Common.VNCLogging.EventHandler) Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

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
