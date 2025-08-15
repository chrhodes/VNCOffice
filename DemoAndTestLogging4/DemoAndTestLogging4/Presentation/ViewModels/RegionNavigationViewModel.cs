using System;

using Prism.Commands;
using Prism.Events;
using Prism.Regions;
using Prism.Services.Dialogs;

using VNC;
using VNC.Core.Mvvm;

namespace DemoAndTestLogging4.Presentation.ViewModels
{
    public class RegionNavigationViewModel : EventViewModelBase, IRegionNavigationViewModel, IInstanceCountVM
    {
        #region Constructors, Initialization, and Load

        private readonly IRegionManager _regionManager;

        public RegionNavigationViewModel(
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

            NavigateUIOneCommand = new DelegateCommand<string>(NavigateUIOne, NavigateUIOneCanExecute);
            NavigateUITwoCommand = new DelegateCommand<string>(NavigateUITwo, NavigateUITwoCanExecute);
            NavigateUIThreeCommand = new DelegateCommand<string>(NavigateUIThree, NavigateUIThreeCanExecute);
            NavigateUIFourCommand = new DelegateCommand<string>(NavigateUIFour, NavigateUIFourCanExecute);
            NavigateUIFiveCommand = new DelegateCommand<string>(NavigateUIFive, NavigateUIFiveCanExecute);

            Message = "RegionNavigationViewModel says hello";

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

        #region NavigateUIOne Command

        public DelegateCommand<string> NavigateUIOneCommand { get; set; }

        // If displaying UserControl
        // public static WindowHost _NavigateUIOneHost = null;

        // If using CommandParameter, figure out TYPE here
        public string NavigateUIOneCommandParameter;

        public string NavigateUIOneContent { get; set; } = "NavigateUIOne";
        public string NavigateUIOneToolTip { get; set; } = "NavigateUIOne Add ToolTip";
        public string NavigateUIOneBetaContent { get; set; } = "NavigateUIOne Beta";
        public string NavigateUIOneBetaToolTip { get; set; } = "NavigateUIOne Beta ToolTip";

        // Can get fancy and use Resources
        //public string NavigateUIOneContent { get; set; } = "ViewName_NavigateUIOneContent";
        //public string NavigateUIOneToolTip { get; set; } = "ViewName_NavigateUIOneContentToolTip";

        // Put these in Resource File
        //    <system:String x:Key="ViewName_NavigateUIOneContent">NavigateUIOne</system:String>
        //    <system:String x:Key="ViewName_NavigateUIOneContentToolTip">NavigateUIOne ToolTip</system:String>

        // If using CommandParameter, figure out TYPE here
        //public void NavigateUIOne(TYPE value)
        public void NavigateUIOne(string action)
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.EventHandler) startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);
            // TODO(crhodes)
            // Do something amazing.

            Message = $"Cool, you called NavigateUIOne {action}";

            PublishStatusMessage(Message);

            switch (action)
            {
                case "beta":
                    _regionManager.RequestNavigate("RegionNavigationView", "uionebeta");
                    break;

                default:
                    _regionManager.RequestNavigate("RegionNavigationView", "uione");

                    // Or this way

                    //IRegion region = _regionManager.Regions["RegionNavigationView"];

                    //region.RequestNavigate("uione");
                    break;
            }

            //IRegion region = _regionManager.Regions[RegionNames.ViewNavigationView];

            //switch (action)
            //{
            //    case "add":
            //        uiOne = new UIOne();
            //        region.Add(uiOne);
            //        NavigateUIOneCommand.RaiseCanExecuteChanged();
            //        ModifyUIOneCommand.RaiseCanExecuteChanged();
            //        RemoveUIOneCommand.RaiseCanExecuteChanged();
            //        break;
            //}

            // If launching a UserControl

            // if (_NavigateUIOneHost is null) _NavigateUIOnenHost = new WindowHost();
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

            // Common.EventAggregator.GetEvent<NavigateUIOneEvent>().Publish();

            // May want EventArgs

            //  EventAggregator.GetEvent<NavigateUIOneEvent>().Publish(
            //      new NavigateUIOneEventArgs()
            //      {
            //            Organization = _collectionMainViewModel.SelectedCollection.Organization,
            //            Process = _contextMainViewModel.Context.SelectedProcess
            //      });

            // Start Cut Four - Put this in PrismEvents

            // public class NavigateUIOneEvent : PubSubEvent { }

            // End Cut Four

            // Start Cut Five - Put this in places that listen for event

            //Common.EventAggregator.GetEvent<NavigateUIOneEvent>().Subscribe(NavigateUIOne);

            // End Cut Five

            if (Common.VNCLogging.EventHandler) Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // If using CommandParameter, figure out TYPE and fix above
        //public bool NavigateUIOneCanExecute(TYPE value)
        public bool NavigateUIOneCanExecute(string action)
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.

            return true;
        }

        #endregion

        #region NavigateUITwo Command

        public DelegateCommand<string> NavigateUITwoCommand { get; set; }

        // If displaying UserControl
        // public static WindowHost _NavigateUITwoHost = null;

        // If using CommandParameter, figure out TYPE here
        public string NavigateUITwoCommandParameter;

        public string NavigateUITwoContent { get; set; } = "NavigateUITwo";
        public string NavigateUITwoToolTip { get; set; } = "NavigateUITwo Add ToolTip";
        public string NavigateUITwoBetaContent { get; set; } = "NavigateUITwo Beta";
        public string NavigateUITwoBetaToolTip { get; set; } = "NavigateUITwo Beta ToolTip";

        // Can get fancy and use Resources
        //public string NavigateUITwoContent { get; set; } = "ViewName_NavigateUITwoContent";
        //public string NavigateUITwoToolTip { get; set; } = "ViewName_NavigateUITwoContentToolTip";

        // Put these in Resource File
        //    <system:String x:Key="ViewName_NavigateUITwoContent">NavigateUITwo</system:String>
        //    <system:String x:Key="ViewName_NavigateUITwoContentToolTip">NavigateUITwo ToolTip</system:String>

        // If using CommandParameter, figure out TYPE here
        //public void NavigateUITwo(TYPE value)
        public void NavigateUITwo(string action)
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.EventHandler) startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);
            // TODO(crhodes)
            // Do something amazing.

            Message = $"Cool, you called NavigateUITwo {action}";

            PublishStatusMessage(Message);

            switch (action)
            {
                case "beta":
                    _regionManager.RequestNavigate("RegionNavigationView", "uitwobeta");
                    break;

                default:
                    //_regionManager.RequestNavigate("RegionNavigationView", "UITwo");

                    // Or this way

                    IRegion region = _regionManager.Regions["RegionNavigationView"];

                    region.RequestNavigate("uitwo");
                    break;
            }

            // If launching a UserControl

            // if (_NavigateUITwoHost is null) _NavigateUITwonHost = new WindowHost();
            // var userControl = new USERCONTROL();

            // _loggingConfigurationHost.DisplayUserControlInHost(
            //     "TITLE GOES HERE",
            //     //Common.DEFAULT_WINDOW_WIDTH,
            //     //Common.DEFAULT_WINDOW_HEIGHT,
            //     (Int32)userControl.Width + Common.WINDOW_HOSTING_USER_CONTROL_WIDTH_PAD,
            //     (Int32)userControl.Height + Common.WINDOW_HOSTING_USER_CONTROL_HEIGHT_PAD,
            //     ShowWindowMode.Modeless_Show,
            //     userControl);

            // Uncomment this if you are telling someTwo else to handle this

            // Common.EventAggregator.GetEvent<NavigateUITwoEvent>().Publish();

            // May want EventArgs

            //  EventAggregator.GetEvent<NavigateUITwoEvent>().Publish(
            //      new NavigateUITwoEventArgs()
            //      {
            //            Organization = _collectionMainViewModel.SelectedCollection.Organization,
            //            Process = _contextMainViewModel.Context.SelectedProcess
            //      });

            // Start Cut Four - Put this in PrismEvents

            // public class NavigateUITwoEvent : PubSubEvent { }

            // End Cut Four

            // Start Cut Five - Put this in places that listen for event

            //Common.EventAggregator.GetEvent<NavigateUITwoEvent>().Subscribe(NavigateUITwo);

            // End Cut Five

            if (Common.VNCLogging.EventHandler) Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // If using CommandParameter, figure out TYPE and fix above
        //public bool NavigateUITwoCanExecute(TYPE value)
        public bool NavigateUITwoCanExecute(string action)
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion

        #region NavigateUIThree Command

        public DelegateCommand<string> NavigateUIThreeCommand { get; set; }

        // If displaying UserControl
        // public static WindowHost _NavigateUIThreeHost = null;

        // If using CommandParameter, figure out TYPE here
        public string NavigateUIThreeCommandParameter;

        public string NavigateUIThreeContent { get; set; } = "NavigateUIThree";
        public string NavigateUIThreeToolTip { get; set; } = "NavigateUIThree Add ToolTip";
        public string NavigateUIThreeBetaContent { get; set; } = "NavigateUIThree Beta";
        public string NavigateUIThreeBetaToolTip { get; set; } = "NavigateUIThree Beta ToolTip";

        // Can get fancy and use Resources
        //public string NavigateUIThreeContent { get; set; } = "ViewName_NavigateUIThreeContent";
        //public string NavigateUIThreeToolTip { get; set; } = "ViewName_NavigateUIThreeContentToolTip";

        // Put these in Resource File
        //    <system:String x:Key="ViewName_NavigateUIThreeContent">NavigateUIThree</system:String>
        //    <system:String x:Key="ViewName_NavigateUIThreeContentToolTip">NavigateUIThree ToolTip</system:String>

        // If using CommandParameter, figure out TYPE here
        //public void NavigateUIThree(TYPE value)
        public void NavigateUIThree(string action)
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.EventHandler) startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);
            // TODO(crhodes)
            // Do something amazing.

            Message = $"Cool, you called NavigateUIThree {action}";

            PublishStatusMessage(Message);

            switch (action)
            {
                case "beta":
                    _regionManager.RequestNavigate("RegionNavigationView", "uithreebeta");
                    break;

                default:
                    _regionManager.RequestNavigate("RegionNavigationView", "uithree");

                    // Or this way

                    //IRegion region = _regionManager.Regions["RegionNavigationView"];

                    //region.RequestNavigate("uiThree");
                    break;
            }

            // If launching a UserControl

            // if (_NavigateUIThreeHost is null) _NavigateUIThreenHost = new WindowHost();
            // var userControl = new USERCONTROL();

            // _loggingConfigurationHost.DisplayUserControlInHost(
            //     "TITLE GOES HERE",
            //     //Common.DEFAULT_WINDOW_WIDTH,
            //     //Common.DEFAULT_WINDOW_HEIGHT,
            //     (Int32)userControl.Width + Common.WINDOW_HOSTING_USER_CONTROL_WIDTH_PAD,
            //     (Int32)userControl.Height + Common.WINDOW_HOSTING_USER_CONTROL_HEIGHT_PAD,
            //     ShowWindowMode.Modeless_Show,
            //     userControl);

            // Uncomment this if you are telling someThree else to handle this

            // Common.EventAggregator.GetEvent<NavigateUIThreeEvent>().Publish();

            // May want EventArgs

            //  EventAggregator.GetEvent<NavigateUIThreeEvent>().Publish(
            //      new NavigateUIThreeEventArgs()
            //      {
            //            Organization = _collectionMainViewModel.SelectedCollection.Organization,
            //            Process = _contextMainViewModel.Context.SelectedProcess
            //      });

            // Start Cut Three - Put this in PrismEvents

            // public class NavigateUIThreeEvent : PubSubEvent { }

            // End Cut Three

            // Start Cut Five - Put this in places that listen for event

            //Common.EventAggregator.GetEvent<NavigateUIThreeEvent>().Subscribe(NavigateUIThree);

            // End Cut Five

            if (Common.VNCLogging.EventHandler) Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // If using CommandParameter, figure out TYPE and fix above
        //public bool NavigateUIThreeCanExecute(TYPE value)
        public bool NavigateUIThreeCanExecute(string action)
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion

        #region NavigateUIFour Command

        public DelegateCommand<string> NavigateUIFourCommand { get; set; }

        // If displaying UserControl
        // public static WindowHost _NavigateUIFourHost = null;

        // If using CommandParameter, figure out TYPE here
        public string NavigateUIFourCommandParameter;

        public string NavigateUIFourContent { get; set; } = "NavigateUIFour";
        public string NavigateUIFourToolTip { get; set; } = "NavigateUIFour Add ToolTip";
        public string NavigateUIFourBetaContent { get; set; } = "NavigateUIFour Beta";
        public string NavigateUIFourBetaToolTip { get; set; } = "NavigateUIFour Beta ToolTip";

        // Can get fancy and use Resources
        //public string NavigateUIFourContent { get; set; } = "ViewName_NavigateUIFourContent";
        //public string NavigateUIFourToolTip { get; set; } = "ViewName_NavigateUIFourContentToolTip";

        // Put these in Resource File
        //    <system:String x:Key="ViewName_NavigateUIFourContent">NavigateUIFour</system:String>
        //    <system:String x:Key="ViewName_NavigateUIFourContentToolTip">NavigateUIFour ToolTip</system:String>

        // If using CommandParameter, figure out TYPE here
        //public void NavigateUIFour(TYPE value)
        public void NavigateUIFour(string action)
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.EventHandler) startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);
            // TODO(crhodes)
            // Do something amazing.

            Message = $"Cool, you called NavigateUIFour {action}";

            PublishStatusMessage(Message);

            switch (action)
            {
                case "beta":
                    _regionManager.RequestNavigate("RegionNavigationView", "uifourbeta");
                    break;

                default:
                    _regionManager.RequestNavigate("RegionNavigationView", "uifour");

                    // Or this way

                    //IRegion region = _regionManager.Regions["RegionNavigationView"];

                    //region.RequestNavigate("uiFour");
                    break;
            }

            // If launching a UserControl

            // if (_NavigateUIFourHost is null) _NavigateUIFournHost = new WindowHost();
            // var userControl = new USERCONTROL();

            // _loggingConfigurationHost.DisplayUserControlInHost(
            //     "TITLE GOES HERE",
            //     //Common.DEFAULT_WINDOW_WIDTH,
            //     //Common.DEFAULT_WINDOW_HEIGHT,
            //     (Int32)userControl.Width + Common.WINDOW_HOSTING_USER_CONTROL_WIDTH_PAD,
            //     (Int32)userControl.Height + Common.WINDOW_HOSTING_USER_CONTROL_HEIGHT_PAD,
            //     ShowWindowMode.Modeless_Show,
            //     userControl);

            // Uncomment this if you are telling someFour else to handle this

            // Common.EventAggregator.GetEvent<NavigateUIFourEvent>().Publish();

            // May want EventArgs

            //  EventAggregator.GetEvent<NavigateUIFourEvent>().Publish(
            //      new NavigateUIFourEventArgs()
            //      {
            //            Organization = _collectionMainViewModel.SelectedCollection.Organization,
            //            Process = _contextMainViewModel.Context.SelectedProcess
            //      });

            // Start Cut Four - Put this in PrismEvents

            // public class NavigateUIFourEvent : PubSubEvent { }

            // End Cut Four

            // Start Cut Five - Put this in places that listen for event

            //Common.EventAggregator.GetEvent<NavigateUIFourEvent>().Subscribe(NavigateUIFour);

            // End Cut Five

            if (Common.VNCLogging.EventHandler) Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // If using CommandParameter, figure out TYPE and fix above
        //public bool NavigateUIFourCanExecute(TYPE value)
        public bool NavigateUIFourCanExecute(string action)
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion

        #region NavigateUIFive Command

        public DelegateCommand<string> NavigateUIFiveCommand { get; set; }

        // If displaying UserControl
        // public static WindowHost _NavigateUIFiveHost = null;

        // If using CommandParameter, figure out TYPE here
        public string NavigateUIFiveCommandParameter;

        public string NavigateUIFiveContent { get; set; } = "NavigateUIFive";
        public string NavigateUIFiveToolTip { get; set; } = "NavigateUIFive Add ToolTip";
        public string NavigateUIFiveBetaContent { get; set; } = "NavigateUIFive Beta";
        public string NavigateUIFiveBetaToolTip { get; set; } = "NavigateUIFive Beta ToolTip";

        // Can get fancy and use Resources
        //public string NavigateUIFiveContent { get; set; } = "ViewName_NavigateUIFiveContent";
        //public string NavigateUIFiveToolTip { get; set; } = "ViewName_NavigateUIFiveContentToolTip";

        // Put these in Resource File
        //    <system:String x:Key="ViewName_NavigateUIFiveContent">NavigateUIFive</system:String>
        //    <system:String x:Key="ViewName_NavigateUIFiveContentToolTip">NavigateUIFive ToolTip</system:String>

        // If using CommandParameter, figure out TYPE here
        //public void NavigateUIFive(TYPE value)
        public void NavigateUIFive(string action)
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.EventHandler) startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);
            // TODO(crhodes)
            // Do something amazing.

            Message = $"Cool, you called NavigateUIFive {action}";

            PublishStatusMessage(Message);

            switch (action)
            {
                case "beta":
                    _regionManager.RequestNavigate("RegionNavigationView", "uifivebeta");
                    break;

                default:
                    _regionManager.RequestNavigate("RegionNavigationView", "uifive");

                    // Or this way

                    //IRegion region = _regionManager.Regions["RegionNavigationView"];

                    //region.RequestNavigate("uiFive");
                    break;
            }

            // If launching a UserControl

            // if (_NavigateUIFiveHost is null) _NavigateUIFivenHost = new WindowHost();
            // var userControl = new USERCONTROL();

            // _loggingConfigurationHost.DisplayUserControlInHost(
            //     "TITLE GOES HERE",
            //     //Common.DEFAULT_WINDOW_WIDTH,
            //     //Common.DEFAULT_WINDOW_HEIGHT,
            //     (Int32)userControl.Width + Common.WINDOW_HOSTING_USER_CONTROL_WIDTH_PAD,
            //     (Int32)userControl.Height + Common.WINDOW_HOSTING_USER_CONTROL_HEIGHT_PAD,
            //     ShowWindowMode.Modeless_Show,
            //     userControl);

            // Uncomment this if you are telling someFive else to handle this

            // Common.EventAggregator.GetEvent<NavigateUIFiveEvent>().Publish();

            // May want EventArgs

            //  EventAggregator.GetEvent<NavigateUIFiveEvent>().Publish(
            //      new NavigateUIFiveEventArgs()
            //      {
            //            Organization = _collectionMainViewModel.SelectedCollection.Organization,
            //            Process = _contextMainViewModel.Context.SelectedProcess
            //      });

            // Start Cut Five - Put this in PrismEvents

            // public class NavigateUIFiveEvent : PubSubEvent { }

            // End Cut Five

            // Start Cut Five - Put this in places that listen for event

            //Common.EventAggregator.GetEvent<NavigateUIFiveEvent>().Subscribe(NavigateUIFive);

            // End Cut Five

            if (Common.VNCLogging.EventHandler) Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // If using CommandParameter, figure out TYPE and fix above
        //public bool NavigateUIFiveCanExecute(TYPE value)
        public bool NavigateUIFiveCanExecute(string action)
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
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
