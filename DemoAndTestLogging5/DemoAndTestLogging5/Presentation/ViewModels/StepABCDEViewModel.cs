using System;
using System.Windows.Input;

using Prism.Commands;
using Prism.Events;
using Prism.Regions;
using Prism.Services.Dialogs;

using VNC;
using VNC.Core.Mvvm;

namespace DemoAndTestLogging5.Presentation.ViewModels
{
    public class StepABCDEViewModel : EventViewModelBase, IStepABCDEViewModel, IInstanceCountVM
    {
        #region Constructors, Initialization, and Load

        IRegionManager _regionManager;

        public StepABCDEViewModel(
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

            BackCommand = new DelegateCommand<string>(Back, BackCanExecute);
            NextCommand = new DelegateCommand<string>(Next, NextCanExecute);

            FirstName_DoubleClick_Command = new DelegateCommand(FirstName_DoubleClick);

            Message = "StepABCDE ViewModel says hello";

            if (Common.VNCLogging.ViewModelLow) Log.VIEWMODEL_LOW("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region Enums (none)


        #endregion

        #region Structures (none)


        #endregion

        #region Fields and Properties

        private bool _stepAComplete;
        public bool StepAComplete
        {
            get => _stepAComplete;
            set
            {
                if (_stepAComplete == value)
                    return;
                _stepAComplete = value;
                OnPropertyChanged();
            }
        }

        private bool _stepBComplete;
        public bool StepBComplete
        {
            get => _stepBComplete;
            set
            {
                if (_stepBComplete == value)
                    return;
                _stepBComplete = value;
                OnPropertyChanged();
            }
        }

        private bool _stepCComplete;
        public bool StepCComplete
        {
            get => _stepCComplete;
            set
            {
                if (_stepCComplete == value)
                    return;
                _stepCComplete = value;
                OnPropertyChanged();
            }
        }

        private bool _stepDComplete;
        public bool StepDComplete
        {
            get => _stepDComplete;
            set
            {
                if (_stepDComplete == value)
                    return;
                _stepDComplete = value;
                OnPropertyChanged();
            }
        }

        private bool _stepEComplete;
        public bool StepEComplete
        {
            get => _stepEComplete;
            set
            {
                if (_stepEComplete == value)
                    return;
                _stepEComplete = value;
                OnPropertyChanged();
            }
        }

        private string _FirstName;
        public string FirstName
        {
            get => _FirstName;
            set
            {
                if (_FirstName == value) return;
                _FirstName = value;
                OnPropertyChanged();
            }
        }

        public string FirstNameToolTip { get; set; }

        #endregion

        #region Event Handlers (none)



        #endregion

        #region Commands

        public ICommand SayHelloCommand { get; private set; }


        #region SayHello Command

        private void SayHello()
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.EventHandler) startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Message = "Hello";
            PublishStatusMessage(Message);

            if (Common.VNCLogging.EventHandler) Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private bool SayHelloCanExecute()
        {
            return true;
        }

        #endregion

        #region Command FirstName DoubleClick

        public DelegateCommand FirstName_DoubleClick_Command { get; set; }

        public void FirstName_DoubleClick()
        {
            Message = "FirstName_DoubleClick";
        }

        #endregion

        #region Back Command

        public DelegateCommand<string> BackCommand { get; set; }

        // If displaying UserControl
        // public static WindowHost _BackHost = null;

        // If using CommandParameter, figure out TYPE here
        //public TYPE BackCommandParameter;

        public string BackContent { get; set; } = "Back";
        public string BackToolTip { get; set; } = "Back ToolTip";

        // Can get fancy and use Resources
        //public string BackContent { get; set; } = "ViewName_BackContent";
        //public string BackToolTip { get; set; } = "ViewName_BackContentToolTip";

        // Put these in Resource File
        //    <system:String x:Key="ViewName_BackContent">Back</system:String>
        //    <system:String x:Key="ViewName_BackContentToolTip">Back ToolTip</system:String>

        // If using CommandParameter, figure out TYPE here
        public void Back(string viewNavigationName)
        //public void Back()
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.EventHandler) startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);
            // TODO(crhodes)
            // Do something amazing.

            Message = $"Cool, you called Back {viewNavigationName}";

            PublishStatusMessage(Message);

            _regionManager.RequestNavigate("MultiStepProcessView", viewNavigationName);

            // If launching a UserControl

            // if (_BackHost is null) _BackHost = new WindowHost();
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

            // Common.EventAggregator.GetEvent<BackEvent>().Publish();

            // May want EventArgs

            //  EventAggregator.GetEvent<BackEvent>().Publish(
            //      new BackEventArgs()
            //      {
            //            Organization = _collectionMainViewModel.SelectedCollection.Organization,
            //            Process = _contextMainViewModel.Context.SelectedProcess
            //      });

            // Start Cut Four - Put this in PrismEvents

            // public class BackEvent : PubSubEvent { }

            // End Cut Four

            // Start Cut Five - Put this in places that listen for event

            //Common.EventAggregator.GetEvent<BackEvent>().Subscribe(Back);

            // End Cut Five

            if (Common.VNCLogging.EventHandler) Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // If using CommandParameter, figure out TYPE and fix above
        public bool BackCanExecute(string viewNavigationName)
        //public bool BackCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion

        #region Next Command

        public DelegateCommand<string> NextCommand { get; set; }

        // If displaying UserControl
        // public static WindowHost _NextHost = null;

        // If using CommandParameter, figure out TYPE here
        //public TYPE NextCommandParameter;

        public string NextContent { get; set; } = "Next";
        public string NextToolTip { get; set; } = "Next ToolTip";

        // Can get fancy and use Resources
        //public string NextContent { get; set; } = "ViewName_NextContent";
        //public string NextToolTip { get; set; } = "ViewName_NextContentToolTip";

        // Put these in Resource File
        //    <system:String x:Key="ViewName_NextContent">Next</system:String>
        //    <system:String x:Key="ViewName_NextContentToolTip">Next ToolTip</system:String>

        // If using CommandParameter, figure out TYPE here
        public void Next(string viewNavigationName)
        //public void Next()
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.EventHandler) startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);
            // TODO(crhodes)
            // Do something amazing.

            Message = $"Cool, you called Next {viewNavigationName}";

            PublishStatusMessage(Message);

            _regionManager.RequestNavigate("MultiStepProcessView", viewNavigationName);

            // If launching a UserControl

            // if (_NextHost is null) _NextHost = new WindowHost();
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

            // Common.EventAggregator.GetEvent<NextEvent>().Publish();

            // May want EventArgs

            //  EventAggregator.GetEvent<NextEvent>().Publish(
            //      new NextEventArgs()
            //      {
            //            Organization = _collectionMainViewModel.SelectedCollection.Organization,
            //            Process = _contextMainViewModel.Context.SelectedProcess
            //      });

            // Start Cut Four - Put this in PrismEvents

            // public class NextEvent : PubSubEvent { }

            // End Cut Four

            // Start Cut Five - Put this in places that listen for event

            //Common.EventAggregator.GetEvent<NextEvent>().Subscribe(Next);

            // End Cut Five

            if (Common.VNCLogging.EventHandler) Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // If using CommandParameter, figure out TYPE and fix above
        public bool NextCanExecute(string value)
        //public bool NextCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion

        #endregion

        #region Public Methods

        #region INavigationAware

        // TODO(crhodes)
        // Override these methods as necessary
        // Default implementation logs and returns false from IsNavigationTarget

        public bool IsNavigationTarget(NavigationContext navigationContext)
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.EventHandler) startTicks = Log.EVENT_HANDLER($"Enter {navigationContext.Uri}", Common.LOG_CATEGORY);

            var n = navigationContext;

            if (Common.VNCLogging.EventHandler) Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);

            // Create New Instance
            //return false;
            // Use Existing Instance
            return true;
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
