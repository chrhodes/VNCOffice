using System;
using System.Windows.Input;

using DemoAndTestLogging5.Core;

using Prism.Commands;
using Prism.Events;
using Prism.Regions;
using Prism.Services.Dialogs;

using VNC;
using VNC.Core.Mvvm;
using VNC.WPF.Presentation.Views;

namespace DemoAndTestLogging5.Presentation.ViewModels
{
    // NOTE(crhodes)
    // This ViewModel has knowledge of concrete views, UIOne, UITwo, UIThree
    // which is generally a bad practice
    // Could put Commands/EventHandlers in ViewInjection.xaml.cs to avoid
    // Generally better to use Uri based Region Navigation in ViewModel
    // See ViewNavigationViewModel

    public class ViewInjectionViewModel : EventViewModelBase, IViewInjectionViewModel, IInstanceCountVM
    {
        #region Constructors, Initialization, and Load

        private readonly IRegionManager _regionManager;

        public ViewInjectionViewModel(
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

            //InjectUIOneCommand = new DelegateCommand(InjectUIOne, InjectUIOneCanExecute);

            // If using CommandParameter, figure out TYPE here and below
            // and remove above declaration
            InjectUIOneCommand = new DelegateCommand<string>(InjectUIOne, InjectUIOneCanExecute);
            InjectUITwoCommand = new DelegateCommand<string>(InjectUITwo, InjectUITwoCanExecute);
            InjectUIThreeCommand = new DelegateCommand<string>(InjectUIThree, InjectUIThreeCanExecute);

            ModifyUIOneCommand = new DelegateCommand<string>(ModifyUIOne, ModifyUIOneCanExecute);
            ModifyUITwoCommand = new DelegateCommand<string>(ModifyUITwo, ModifyUITwoCanExecute);
            ModifyUIThreeCommand = new DelegateCommand<string>(ModifyUIThree, ModifyUIThreeCanExecute);

            RemoveUIOneCommand = new DelegateCommand<string>(RemoveUIOne, RemoveUIOneCanExecute);
            RemoveUITwoCommand = new DelegateCommand<string>(RemoveUITwo, RemoveUITwoCanExecute);
            RemoveUIThreeCommand = new DelegateCommand<string>(RemoveUIThree, RemoveUIThreeCanExecute);


            Message = "ViewInjectionViewModel says hello";

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

        #region InjectUIOne Command

        public DelegateCommand<string> InjectUIOneCommand { get; set; }

        // If displaying UserControl
        // public static WindowHost _InjectUIOneHost = null;

        // If using CommandParameter, figure out TYPE here
        public string InjectUIOneCommandParameter;

        UIOne? uiOne = null;

        public string InjectUIOneAddContent { get; set; } = "InjectUIOne Add";
        public string InjectUIOneActivateContent { get; set; } = "InjectUIOne Activate";
        public string InjectUIOneDeactivateContent { get; set; } = "InjectUIOne Deactivate";
        public string InjectUIOneRemoveContent { get; set; } = "InjectUIOne Remove";
        public string InjectUIOneToolTip { get; set; } = "InjectUIOne Add ToolTip";

        // Can get fancy and use Resources
        //public string InjectUIOneContent { get; set; } = "ViewName_InjectUIOneContent";
        //public string InjectUIOneToolTip { get; set; } = "ViewName_InjectUIOneContentToolTip";

        // Put these in Resource File
        //    <system:String x:Key="ViewName_InjectUIOneContent">InjectUIOne</system:String>
        //    <system:String x:Key="ViewName_InjectUIOneContentToolTip">InjectUIOne ToolTip</system:String>

        // If using CommandParameter, figure out TYPE here
        //public void InjectUIOne(TYPE value)
        public void InjectUIOne(string action)
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.EventHandler) startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);
            // TODO(crhodes)
            // Do something amazing.

            Message = "Cool, you called InjectUIOne";

            PublishStatusMessage(Message);

            IRegion region = _regionManager.Regions[RegionNames.ViewInjectionView];

            switch (action)
            {
                case "add":
                    uiOne = new UIOne();
                    region.Add(uiOne);
                    InjectUIOneCommand.RaiseCanExecuteChanged();
                    ModifyUIOneCommand.RaiseCanExecuteChanged();
                    RemoveUIOneCommand.RaiseCanExecuteChanged();
                    break;
            }

            // If launching a UserControl

            // if (_InjectUIOneHost is null) _InjectUIOneHost = new WindowHost();
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

            // Common.EventAggregator.GetEvent<InjectUIOneEvent>().Publish();

            // May want EventArgs

            //  EventAggregator.GetEvent<InjectUIOneEvent>().Publish(
            //      new InjectUIOneEventArgs()
            //      {
            //            Organization = _collectionMainViewModel.SelectedCollection.Organization,
            //            Process = _contextMainViewModel.Context.SelectedProcess
            //      });

            // Start Cut Four - Put this in PrismEvents

            // public class InjectUIOneEvent : PubSubEvent { }

            // End Cut Four

            // Start Cut Five - Put this in places that listen for event

            //Common.EventAggregator.GetEvent<InjectUIOneEvent>().Subscribe(InjectUIOne);

            // End Cut Five

            if (Common.VNCLogging.EventHandler) Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // If using CommandParameter, figure out TYPE and fix above
        //public bool InjectUIOneCanExecute(TYPE value)
        public bool InjectUIOneCanExecute(string action)
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.

            return (uiOne is null) ? true : false;
        }

        #endregion

        #region InjectUITwo Command

        public DelegateCommand<string> InjectUITwoCommand { get; set; }

        // If displaying UserControl
        // public static WindowHost _InjectUITwoHost = null;

        // If using CommandParameter, figure out TYPE here
        public string InjectUITwoCommandParameter;

        UITwo? uiTwo = null;

        public string InjectUITwoAddContent { get; set; } = "InjectUITwo Add";
        public string InjectUITwoActivateContent { get; set; } = "InjectUITwo Activate";
        public string InjectUITwoDeactivateContent { get; set; } = "InjectUITwo Deactivate";
        public string InjectUITwoRemoveContent { get; set; } = "InjectUITwo Remove";
        public string InjectUITwoToolTip { get; set; } = "InjectUITwo Add ToolTip";

        // Can get fancy and use Resources
        //public string InjectUITwoContent { get; set; } = "ViewName_InjectUITwoContent";
        //public string InjectUITwoToolTip { get; set; } = "ViewName_InjectUITwoContentToolTip";

        // Put these in Resource File
        //    <system:String x:Key="ViewName_InjectUITwoContent">InjectUITwo</system:String>
        //    <system:String x:Key="ViewName_InjectUITwoContentToolTip">InjectUITwo ToolTip</system:String>

        // If using CommandParameter, figure out TYPE here
        //public void InjectUITwo(TYPE value)
        public void InjectUITwo(string action)
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.EventHandler) startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);
            // TODO(crhodes)
            // Do something amazing.

            Message = "Cool, you called InjectUITwo";

            PublishStatusMessage(Message);

            IRegion region = _regionManager.Regions[RegionNames.ViewInjectionView];

            switch (action)
            {
                case "add":
                    uiTwo = new UITwo();
                    region.Add(uiTwo);
                    InjectUITwoCommand.RaiseCanExecuteChanged();
                    ModifyUITwoCommand.RaiseCanExecuteChanged();
                    RemoveUITwoCommand.RaiseCanExecuteChanged();
                    break;
            }

            // If launching a UserControl

            // if (_InjectUITwoHost is null) _InjectUITwoHost = new WindowHost();
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

            // Common.EventAggregator.GetEvent<InjectUITwoEvent>().Publish();

            // May want EventArgs

            //  EventAggregator.GetEvent<InjectUITwoEvent>().Publish(
            //      new InjectUITwoEventArgs()
            //      {
            //            Organization = _collectionMainViewModel.SelectedCollection.Organization,
            //            Process = _contextMainViewModel.Context.SelectedProcess
            //      });

            // Start Cut Four - Put this in PrismEvents

            // public class InjectUITwoEvent : PubSubEvent { }

            // End Cut Four

            // Start Cut Five - Put this in places that listen for event

            //Common.EventAggregator.GetEvent<InjectUITwoEvent>().Subscribe(InjectUITwo);

            // End Cut Five

            if (Common.VNCLogging.EventHandler) Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // If using CommandParameter, figure out TYPE and fix above
        //public bool InjectUITwoCanExecute(TYPE value)
        public bool InjectUITwoCanExecute(string action)
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return (uiTwo is null) ? true : false;
        }

        #endregion

        #region InjectUIThree Command

        public DelegateCommand<string> InjectUIThreeCommand { get; set; }

        // If displaying UserControl
        // public static WindowHost _InjectUIThreeHost = null;

        // If using CommandParameter, figure out TYPE here
        public string InjectUIThreeCommandParameter;

        UIThree? uiThree = null;

        public string InjectUIThreeAddContent { get; set; } = "InjectUIThree Add";
        public string InjectUIThreeActivateContent { get; set; } = "InjectUIThree Activate";
        public string InjectUIThreeDeactivateContent { get; set; } = "InjectUIThree Deactivate";
        public string InjectUIThreeRemoveContent { get; set; } = "InjectUIThree Remove";
        public string InjectUIThreeToolTip { get; set; } = "InjectUIThree Add ToolTip";

        // Can get fancy and use Resources
        //public string InjectUIThreeContent { get; set; } = "ViewName_InjectUIThreeContent";
        //public string InjectUIThreeToolTip { get; set; } = "ViewName_InjectUIThreeContentToolTip";

        // Put these in Resource File
        //    <system:String x:Key="ViewName_InjectUIThreeContent">InjectUIThree</system:String>
        //    <system:String x:Key="ViewName_InjectUIThreeContentToolTip">InjectUIThree ToolTip</system:String>

        // If using CommandParameter, figure out TYPE here
        //public void InjectUIThree(TYPE value)
        public void InjectUIThree(string action)
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.EventHandler) startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);
            // TODO(crhodes)
            // Do something amazing.

            Message = "Cool, you called InjectUIThree";

            PublishStatusMessage(Message);

            IRegion region = _regionManager.Regions[RegionNames.ViewInjectionView];

            switch (action)
            {
                case "add":
                    uiThree = new UIThree();
                    region.Add(uiThree);
                    InjectUIThreeCommand.RaiseCanExecuteChanged();
                    ModifyUIThreeCommand.RaiseCanExecuteChanged();
                    RemoveUIThreeCommand.RaiseCanExecuteChanged();
                    break;
            }

            // If launching a UserControl

            // if (_InjectUIThreeHost is null) _InjectUIThreeHost = new WindowHost();
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

            // Common.EventAggregator.GetEvent<InjectUIThreeEvent>().Publish();

            // May want EventArgs

            //  EventAggregator.GetEvent<InjectUIThreeEvent>().Publish(
            //      new InjectUIThreeEventArgs()
            //      {
            //            Organization = _collectionMainViewModel.SelectedCollection.Organization,
            //            Process = _contextMainViewModel.Context.SelectedProcess
            //      });

            // Start Cut Four - Put this in PrismEvents

            // public class InjectUIThreeEvent : PubSubEvent { }

            // End Cut Four

            // Start Cut Five - Put this in places that listen for event

            //Common.EventAggregator.GetEvent<InjectUIThreeEvent>().Subscribe(InjectUIThree);

            // End Cut Five

            if (Common.VNCLogging.EventHandler) Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // If using CommandParameter, figure out TYPE and fix above
        //public bool InjectUIThreeCanExecute(TYPE value)
        public bool InjectUIThreeCanExecute(string action)
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return (uiThree is null) ? true : false;
        }

        #endregion

        #region ModifyUIOne Command

        public DelegateCommand<string> ModifyUIOneCommand { get; set; }

        // If displaying UserControl
        // public static WindowHost _ModifyUIOneHost = null;

        // If using CommandParameter, figure out TYPE here
        public string ModifyUIOneCommandParameter;

        public string ModifyUIOneAddContent { get; set; } = "ModifyUIOne Add";
        public string ModifyUIOneActivateContent { get; set; } = "ModifyUIOne Activate";
        public string ModifyUIOneDeactivateContent { get; set; } = "ModifyUIOne Deactivate";
        public string ModifyUIOneRemoveContent { get; set; } = "ModifyUIOne Remove";
        public string ModifyUIOneToolTip { get; set; } = "ModifyUIOne Add ToolTip";

        // Can get fancy and use Resources
        //public string ModifyUIOneContent { get; set; } = "ViewName_ModifyUIOneContent";
        //public string ModifyUIOneToolTip { get; set; } = "ViewName_ModifyUIOneContentToolTip";

        // Put these in Resource File
        //    <system:String x:Key="ViewName_ModifyUIOneContent">ModifyUIOne</system:String>
        //    <system:String x:Key="ViewName_ModifyUIOneContentToolTip">ModifyUIOne ToolTip</system:String>

        // If using CommandParameter, figure out TYPE here
        //public void ModifyUIOne(TYPE value)
        public void ModifyUIOne(string action)
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.EventHandler) startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);
            // TODO(crhodes)
            // Do something amazing.

            Message = "Cool, you called ModifyUIOne";

            PublishStatusMessage(Message);

            IRegion region = _regionManager.Regions[RegionNames.ViewInjectionView];

            switch (action)
            {
                case "activate":
                    region.Activate(uiOne);
                    break;

                case "deactivate":
                    region.Deactivate(uiOne);
                    break;
            }

            // If launching a UserControl

            // if (_ModifyUIOneHost is null) _ModifyUIOneHost = new WindowHost();
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

            // Common.EventAggregator.GetEvent<ModifyUIOneEvent>().Publish();

            // May want EventArgs

            //  EventAggregator.GetEvent<ModifyUIOneEvent>().Publish(
            //      new ModifyUIOneEventArgs()
            //      {
            //            Organization = _collectionMainViewModel.SelectedCollection.Organization,
            //            Process = _contextMainViewModel.Context.SelectedProcess
            //      });

            // Start Cut Four - Put this in PrismEvents

            // public class ModifyUIOneEvent : PubSubEvent { }

            // End Cut Four

            // Start Cut Five - Put this in places that listen for event

            //Common.EventAggregator.GetEvent<ModifyUIOneEvent>().Subscribe(ModifyUIOne);

            // End Cut Five

            if (Common.VNCLogging.EventHandler) Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // If using CommandParameter, figure out TYPE and fix above
        //public bool ModifyUIOneCanExecute(TYPE value)
        public bool ModifyUIOneCanExecute(string action)
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.

            return (uiOne is null) ? false : true;
        }

        #endregion

        #region ModifyUITwo Command

        public DelegateCommand<string> ModifyUITwoCommand { get; set; }

        // If displaying UserControl
        // public static WindowHost _ModifyUITwoHost = null;

        // If using CommandParameter, figure out TYPE here
        public string ModifyUITwoCommandParameter;

        public string ModifyUITwoAddContent { get; set; } = "ModifyUITwo Add";
        public string ModifyUITwoActivateContent { get; set; } = "ModifyUITwo Activate";
        public string ModifyUITwoDeactivateContent { get; set; } = "ModifyUITwo Deactivate";
        public string ModifyUITwoRemoveContent { get; set; } = "ModifyUITwo Remove";
        public string ModifyUITwoToolTip { get; set; } = "ModifyUITwo Add ToolTip";

        // Can get fancy and use Resources
        //public string ModifyUITwoContent { get; set; } = "ViewName_ModifyUITwoContent";
        //public string ModifyUITwoToolTip { get; set; } = "ViewName_ModifyUITwoContentToolTip";

        // Put these in Resource File
        //    <system:String x:Key="ViewName_ModifyUITwoContent">ModifyUITwo</system:String>
        //    <system:String x:Key="ViewName_ModifyUITwoContentToolTip">ModifyUITwo ToolTip</system:String>

        // If using CommandParameter, figure out TYPE here
        //public void ModifyUITwo(TYPE value)
        public void ModifyUITwo(string action)
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.EventHandler) startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);
            // TODO(crhodes)
            // Do something amazing.

            Message = "Cool, you called ModifyUITwo";

            PublishStatusMessage(Message);

            IRegion region = _regionManager.Regions[RegionNames.ViewInjectionView];

            switch (action)
            {
                case "activate":
                    region.Activate(uiTwo);
                    break;

                case "deactivate":
                    region.Deactivate(uiTwo);
                    break;
            }

            // If launching a UserControl

            // if (_ModifyUITwoHost is null) _ModifyUITwoHost = new WindowHost();
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

            // Common.EventAggregator.GetEvent<ModifyUITwoEvent>().Publish();

            // May want EventArgs

            //  EventAggregator.GetEvent<ModifyUITwoEvent>().Publish(
            //      new ModifyUITwoEventArgs()
            //      {
            //            Organization = _collectionMainViewModel.SelectedCollection.Organization,
            //            Process = _contextMainViewModel.Context.SelectedProcess
            //      });

            // Start Cut Four - Put this in PrismEvents

            // public class ModifyUITwoEvent : PubSubEvent { }

            // End Cut Four

            // Start Cut Five - Put this in places that listen for event

            //Common.EventAggregator.GetEvent<ModifyUITwoEvent>().Subscribe(ModifyUITwo);

            // End Cut Five

            if (Common.VNCLogging.EventHandler) Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // If using CommandParameter, figure out TYPE and fix above
        //public bool ModifyUITwoCanExecute(TYPE value)
        public bool ModifyUITwoCanExecute(string action)
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return (uiTwo is null) ? false : true;
        }

        #endregion

        #region ModifyUIThree Command

        public DelegateCommand<string> ModifyUIThreeCommand { get; set; }

        // If displaying UserControl
        // public static WindowHost _ModifyUIThreeHost = null;

        // If using CommandParameter, figure out TYPE here
        public string ModifyUIThreeCommandParameter;

        public string ModifyUIThreeAddContent { get; set; } = "ModifyUIThree Add";
        public string ModifyUIThreeActivateContent { get; set; } = "ModifyUIThree Activate";
        public string ModifyUIThreeDeactivateContent { get; set; } = "ModifyUIThree Deactivate";
        public string ModifyUIThreeRemoveContent { get; set; } = "ModifyUIThree Remove";
        public string ModifyUIThreeToolTip { get; set; } = "ModifyUIThree Add ToolTip";

        // Can get fancy and use Resources
        //public string ModifyUIThreeContent { get; set; } = "ViewName_ModifyUIThreeContent";
        //public string ModifyUIThreeToolTip { get; set; } = "ViewName_ModifyUIThreeContentToolTip";

        // Put these in Resource File
        //    <system:String x:Key="ViewName_ModifyUIThreeContent">ModifyUIThree</system:String>
        //    <system:String x:Key="ViewName_ModifyUIThreeContentToolTip">ModifyUIThree ToolTip</system:String>

        // If using CommandParameter, figure out TYPE here
        //public void ModifyUIThree(TYPE value)
        public void ModifyUIThree(string action)
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.EventHandler) startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);
            // TODO(crhodes)
            // Do something amazing.

            Message = "Cool, you called ModifyUIThree";

            PublishStatusMessage(Message);

            IRegion region = _regionManager.Regions[RegionNames.ViewInjectionView];

            switch (action)
            {
                case "activate":
                    region.Activate(uiThree);
                    break;

                case "deactivate":
                    region.Deactivate(uiThree);
                    break;
            }

            // If launching a UserControl

            // if (_ModifyUIThreeHost is null) _ModifyUIThreeHost = new WindowHost();
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

            // Common.EventAggregator.GetEvent<ModifyUIThreeEvent>().Publish();

            // May want EventArgs

            //  EventAggregator.GetEvent<ModifyUIThreeEvent>().Publish(
            //      new ModifyUIThreeEventArgs()
            //      {
            //            Organization = _collectionMainViewModel.SelectedCollection.Organization,
            //            Process = _contextMainViewModel.Context.SelectedProcess
            //      });

            // Start Cut Four - Put this in PrismEvents

            // public class ModifyUIThreeEvent : PubSubEvent { }

            // End Cut Four

            // Start Cut Five - Put this in places that listen for event

            //Common.EventAggregator.GetEvent<ModifyUIThreeEvent>().Subscribe(ModifyUIThree);

            // End Cut Five

            if (Common.VNCLogging.EventHandler) Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // If using CommandParameter, figure out TYPE and fix above
        //public bool ModifyUIThreeCanExecute(TYPE value)
        public bool ModifyUIThreeCanExecute(string action)
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return (uiThree is null) ? false : true;
        }

        #endregion

        #region RemoveUIOne Command

        public DelegateCommand<string> RemoveUIOneCommand { get; set; }

        // If displaying UserControl
        // public static WindowHost _RemoveUIOneHost = null;

        // If using CommandParameter, figure out TYPE here
        public string RemoveUIOneCommandParameter;

        public string RemoveUIOneAddContent { get; set; } = "RemoveUIOne Add";
        public string RemoveUIOneActivateContent { get; set; } = "RemoveUIOne Activate";
        public string RemoveUIOneDeactivateContent { get; set; } = "RemoveUIOne Deactivate";
        public string RemoveUIOneRemoveContent { get; set; } = "RemoveUIOne Remove";
        public string RemoveUIOneToolTip { get; set; } = "RemoveUIOne Add ToolTip";

        // Can get fancy and use Resources
        //public string RemoveUIOneContent { get; set; } = "ViewName_RemoveUIOneContent";
        //public string RemoveUIOneToolTip { get; set; } = "ViewName_RemoveUIOneContentToolTip";

        // Put these in Resource File
        //    <system:String x:Key="ViewName_RemoveUIOneContent">RemoveUIOne</system:String>
        //    <system:String x:Key="ViewName_RemoveUIOneContentToolTip">RemoveUIOne ToolTip</system:String>

        // If using CommandParameter, figure out TYPE here
        //public void RemoveUIOne(TYPE value)
        public void RemoveUIOne(string action)
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.EventHandler) startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);
            // TODO(crhodes)
            // Do something amazing.

            Message = "Cool, you called RemoveUIOne";

            PublishStatusMessage(Message);

            IRegion region = _regionManager.Regions[RegionNames.ViewInjectionView];

            switch (action)
            {
                case "remove":
                    region.Remove(uiOne);
                    uiOne = null;
                    RemoveUIOneCommand.RaiseCanExecuteChanged();
                    InjectUIOneCommand.RaiseCanExecuteChanged();
                    ModifyUIOneCommand.RaiseCanExecuteChanged();
                    break;
            }

            // If launching a UserControl

            // if (_RemoveUIOneHost is null) _RemoveUIOneHost = new WindowHost();
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

            // Common.EventAggregator.GetEvent<RemoveUIOneEvent>().Publish();

            // May want EventArgs

            //  EventAggregator.GetEvent<RemoveUIOneEvent>().Publish(
            //      new RemoveUIOneEventArgs()
            //      {
            //            Organization = _collectionMainViewModel.SelectedCollection.Organization,
            //            Process = _contextMainViewModel.Context.SelectedProcess
            //      });

            // Start Cut Four - Put this in PrismEvents

            // public class RemoveUIOneEvent : PubSubEvent { }

            // End Cut Four

            // Start Cut Five - Put this in places that listen for event

            //Common.EventAggregator.GetEvent<RemoveUIOneEvent>().Subscribe(RemoveUIOne);

            // End Cut Five

            if (Common.VNCLogging.EventHandler) Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // If using CommandParameter, figure out TYPE and fix above
        //public bool RemoveUIOneCanExecute(TYPE value)
        public bool RemoveUIOneCanExecute(string action)
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.

            return (uiOne is null) ? false : true;
        }

        #endregion

        #region RemoveUITwo Command

        public DelegateCommand<string> RemoveUITwoCommand { get; set; }

        // If displaying UserControl
        // public static WindowHost _RemoveUITwoHost = null;

        // If using CommandParameter, figure out TYPE here
        public string RemoveUITwoCommandParameter;

        public string RemoveUITwoAddContent { get; set; } = "RemoveUITwo Add";
        public string RemoveUITwoActivateContent { get; set; } = "RemoveUITwo Activate";
        public string RemoveUITwoDeactivateContent { get; set; } = "RemoveUITwo Deactivate";
        public string RemoveUITwoRemoveContent { get; set; } = "RemoveUITwo Remove";
        public string RemoveUITwoToolTip { get; set; } = "RemoveUITwo Add ToolTip";

        // Can get fancy and use Resources
        //public string RemoveUITwoContent { get; set; } = "ViewName_RemoveUITwoContent";
        //public string RemoveUITwoToolTip { get; set; } = "ViewName_RemoveUITwoContentToolTip";

        // Put these in Resource File
        //    <system:String x:Key="ViewName_RemoveUITwoContent">RemoveUITwo</system:String>
        //    <system:String x:Key="ViewName_RemoveUITwoContentToolTip">RemoveUITwo ToolTip</system:String>

        // If using CommandParameter, figure out TYPE here
        //public void RemoveUITwo(TYPE value)
        public void RemoveUITwo(string action)
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.EventHandler) startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);
            // TODO(crhodes)
            // Do something amazing.

            Message = "Cool, you called RemoveUITwo";

            PublishStatusMessage(Message);

            IRegion region = _regionManager.Regions[RegionNames.ViewInjectionView];

            switch (action)
            {
                case "remove":
                    region.Remove(uiTwo);
                    uiTwo = null;
                    RemoveUITwoCommand.RaiseCanExecuteChanged();
                    InjectUITwoCommand.RaiseCanExecuteChanged();
                    ModifyUITwoCommand.RaiseCanExecuteChanged();
                    break;
            }

            // If launching a UserControl

            // if (_RemoveUITwoHost is null) _RemoveUITwoHost = new WindowHost();
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

            // Common.EventAggregator.GetEvent<RemoveUITwoEvent>().Publish();

            // May want EventArgs

            //  EventAggregator.GetEvent<RemoveUITwoEvent>().Publish(
            //      new RemoveUITwoEventArgs()
            //      {
            //            Organization = _collectionMainViewModel.SelectedCollection.Organization,
            //            Process = _contextMainViewModel.Context.SelectedProcess
            //      });

            // Start Cut Four - Put this in PrismEvents

            // public class RemoveUITwoEvent : PubSubEvent { }

            // End Cut Four

            // Start Cut Five - Put this in places that listen for event

            //Common.EventAggregator.GetEvent<RemoveUITwoEvent>().Subscribe(RemoveUITwo);

            // End Cut Five

            if (Common.VNCLogging.EventHandler) Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // If using CommandParameter, figure out TYPE and fix above
        //public bool RemoveUITwoCanExecute(TYPE value)
        public bool RemoveUITwoCanExecute(string action)
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return (uiTwo is null) ? false : true;
        }

        #endregion

        #region RemoveUIThree Command

        public DelegateCommand<string> RemoveUIThreeCommand { get; set; }

        // If displaying UserControl
        // public static WindowHost _RemoveUIThreeHost = null;

        // If using CommandParameter, figure out TYPE here
        public string RemoveUIThreeCommandParameter;

        public string RemoveUIThreeAddContent { get; set; } = "RemoveUIThree Add";
        public string RemoveUIThreeActivateContent { get; set; } = "RemoveUIThree Activate";
        public string RemoveUIThreeDeactivateContent { get; set; } = "RemoveUIThree Deactivate";
        public string RemoveUIThreeRemoveContent { get; set; } = "RemoveUIThree Remove";
        public string RemoveUIThreeToolTip { get; set; } = "RemoveUIThree Add ToolTip";

        // Can get fancy and use Resources
        //public string RemoveUIThreeContent { get; set; } = "ViewName_RemoveUIThreeContent";
        //public string RemoveUIThreeToolTip { get; set; } = "ViewName_RemoveUIThreeContentToolTip";

        // Put these in Resource File
        //    <system:String x:Key="ViewName_RemoveUIThreeContent">RemoveUIThree</system:String>
        //    <system:String x:Key="ViewName_RemoveUIThreeContentToolTip">RemoveUIThree ToolTip</system:String>

        // If using CommandParameter, figure out TYPE here
        //public void RemoveUIThree(TYPE value)
        public void RemoveUIThree(string action)
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.EventHandler) startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);
            // TODO(crhodes)
            // Do something amazing.

            Message = "Cool, you called RemoveUIThree";

            PublishStatusMessage(Message);

            IRegion region = _regionManager.Regions[RegionNames.ViewInjectionView];

            switch (action)
            {
                case "remove":
                    region.Remove(uiThree);
                    uiThree = null;
                    RemoveUIThreeCommand.RaiseCanExecuteChanged();
                    InjectUIThreeCommand.RaiseCanExecuteChanged();
                    ModifyUIThreeCommand.RaiseCanExecuteChanged();
                    break;
            }

            // If launching a UserControl

            // if (_RemoveUIThreeHost is null) _RemoveUIThreeHost = new WindowHost();
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

            // Common.EventAggregator.GetEvent<RemoveUIThreeEvent>().Publish();

            // May want EventArgs

            //  EventAggregator.GetEvent<RemoveUIThreeEvent>().Publish(
            //      new RemoveUIThreeEventArgs()
            //      {
            //            Organization = _collectionMainViewModel.SelectedCollection.Organization,
            //            Process = _contextMainViewModel.Context.SelectedProcess
            //      });

            // Start Cut Four - Put this in PrismEvents

            // public class RemoveUIThreeEvent : PubSubEvent { }

            // End Cut Four

            // Start Cut Five - Put this in places that listen for event

            //Common.EventAggregator.GetEvent<RemoveUIThreeEvent>().Subscribe(RemoveUIThree);

            // End Cut Five

            if (Common.VNCLogging.EventHandler) Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // If using CommandParameter, figure out TYPE and fix above
        //public bool RemoveUIThreeCanExecute(TYPE value)
        public bool RemoveUIThreeCanExecute(string action)
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return (uiThree is null) ? false : true;
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

            if (Common.VNCLogging.EventHandler) Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #endregion

        #region Protected Methods (none)


        #endregion

        #region Private Methods

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
