using System;

using DemoAndTestLogging4.Core;

using Prism.Commands;
using Prism.Events;
using Prism.Regions;
using Prism.Services.Dialogs;

using VNC;
using VNC.Core.Mvvm;
using VNC.WPF.Presentation.Views;

namespace DemoAndTestLogging4.Presentation.ViewModels
{
    // NOTE(crhodes)
    // This ViewModel has knowledge of concrete views, UIOne, UITwo, UIThree
    // which is generally a bad practice
    // Could put Commands/EventHandlers in ViewDiscovery.xaml.cs to avoid
    // Generally better to use Uri based Region Navigation in ViewModel
    // See ViewNavigationViewModel

    public class ViewDiscoveryViewModel : EventViewModelBase, IViewDiscoveryViewModel, IInstanceCountVM
    {
        #region Constructors, Initialization, and Load

        private readonly IRegionManager _regionManager;

        public ViewDiscoveryViewModel(
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

            DiscoverViewsCommand = new DelegateCommand(DiscoverViews, DiscoverViewsCanExecute);

            Message = "ViewDiscoveryViewModel says hello";

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


        #region DiscoverViews Command

        public DelegateCommand DiscoverViewsCommand { get; set; }

        // If displaying UserControl
        // public static WindowHost _DiscoverViewsHost = null;

        // If using CommandParameter, figure out TYPE here
        public string DiscoverViewsCommandParameter;

        public string DiscoverViewsContent { get; set; } = "DiscoverViews";
        public string DiscoverViewsToolTip { get; set; } = "DiscoverViews ToolTip";

        // Can get fancy and use Resources
        //public string DiscoverViewsContent { get; set; } = "ViewName_DiscoverViewsContent";
        //public string DiscoverViewsToolTip { get; set; } = "ViewName_DiscoverViewsContentToolTip";

        // Put these in Resource File
        //    <system:String x:Key="ViewName_DiscoverViewsContent">DiscoverViews</system:String>
        //    <system:String x:Key="ViewName_DiscoverViewsContentToolTip">DiscoverViews ToolTip</system:String>

        // If using CommandParameter, figure out TYPE here
        //public void DiscoverViews(TYPE value)
        public void DiscoverViews()
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.EventHandler) startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);
            // TODO(crhodes)
            // Do something amazing.

            Message = "Cool, you called DiscoverViews";

            PublishStatusMessage(Message);

            _regionManager.RegisterViewWithRegion(RegionNames.ViewDiscoveryViewOne, typeof(UIOne));
            _regionManager.RegisterViewWithRegion(RegionNames.ViewDiscoveryViewTwo, typeof(UITwo));
            _regionManager.RegisterViewWithRegion(RegionNames.ViewDiscoveryViewThree, typeof(UIThree));
            _regionManager.RegisterViewWithRegion(RegionNames.ViewDiscoveryViewFour, typeof(UIFour));
            _regionManager.RegisterViewWithRegion(RegionNames.ViewDiscoveryViewFive, typeof(UIFive));

            _regionManager.RegisterViewWithRegion(RegionNames.ViewDiscoveryViewThreeBeta, typeof(UIThree_Beta));
            _regionManager.RegisterViewWithRegion(RegionNames.ViewDiscoveryViewFourBeta, typeof(UIFour_Beta));
            _regionManager.RegisterViewWithRegion(RegionNames.ViewDiscoveryViewFiveBeta, typeof(UIFive_Beta));

            // If launching a UserControl

            // if (_DiscoverViewsHost is null) _DiscoverViewsHost = new WindowHost();
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

            // Common.EventAggregator.GetEvent<DiscoverViewsEvent>().Publish();

            // May want EventArgs

            //  EventAggregator.GetEvent<DiscoverViewsEvent>().Publish(
            //      new DiscoverViewsEventArgs()
            //      {
            //            Organization = _collectionMainViewModel.SelectedCollection.Organization,
            //            Process = _contextMainViewModel.Context.SelectedProcess
            //      });

            // Start Cut Four - Put this in PrismEvents

            // public class DiscoverViewsEvent : PubSubEvent { }

            // End Cut Four

            // Start Cut Five - Put this in places that listen for event

            //Common.EventAggregator.GetEvent<DiscoverViewsEvent>().Subscribe(DiscoverViews);

            // End Cut Five

            if (Common.VNCLogging.EventHandler) Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // If using CommandParameter, figure out TYPE and fix above
        //public bool DiscoverViewsCanExecute(TYPE value)
        public bool DiscoverViewsCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.

            return true;
        }

        #endregion

        #endregion

        #region Protected Methods (none)


        #endregion

        #region Private Methods

        private bool SayHelloCanExecute()
        {
            return true;
        }

        private void SayHello()
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.EventHandler) startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Message = "Hello";

            if (Common.VNCLogging.EventHandler) Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

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
