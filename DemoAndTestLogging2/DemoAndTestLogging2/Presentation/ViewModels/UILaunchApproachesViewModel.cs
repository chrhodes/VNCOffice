using System;
using System.Windows.Input;

using Prism.Commands;
using Prism.Events;
using Prism.Services.Dialogs;

using VNC;
using VNC.Core.Events;
using VNC.Core.Mvvm;
using VNC.Core.Presentation;
using VNC.WPF.Presentation.Dx.Views;
using VNC.WPF.Presentation.Views;

namespace DemoAndTestLogging2.Presentation.ViewModels
{
    public class UILaunchApproachesViewModel : EventViewModelBase, IUILaunchApproachesViewModel, IInstanceCountVM
    {
        #region Constructors, Initialization, and Load

        public UILaunchApproachesViewModel(
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

            SayHelloCommand = new DelegateCommand(SayHello, SayHelloCanExecute);

            LaunchCylon1Command = new DelegateCommand(LaunchCylon1, LaunchCylon1CanExecute);
            LaunchCylon2Command = new DelegateCommand(LaunchCylon2, LaunchCylon2CanExecute);
            ThemedWindowHostModelessCommand = new DelegateCommand(ThemedWindowHostModeless, ThemedWindowHostModelessCanExecute);
            ThemedWindowHostModalCommand = new DelegateCommand(ThemedWindowHostModal, ThemedWindowHostModalCanExecute);

            DXLayoutControlCommand = new DelegateCommand(DXLayoutControl, DXLayoutControlCanExecute);
            DXDockLayoutControlCommand = new DelegateCommand(DXDockLayoutControl, DXDockLayoutControlCanExecute);
            DXDockLayoutManagerCommand = new DelegateCommand(DXDockLayoutManager, DXDockLayoutManagerCanExecute);

            MVVM1Command = new DelegateCommand(MVVM1, MVVM1CanExecute);
            MVVM2Command = new DelegateCommand(MVVM2, MVVM2CanExecute);
            MVVM3Command = new DelegateCommand(MVVM3, MVVM3CanExecute);
            MVVM4Command = new DelegateCommand(MVVM4, MVVM4CanExecute);
            MVVM5Command = new DelegateCommand(MVVM5, MVVM5CanExecute);
            MVVM6Command = new DelegateCommand(MVVM6, MVVM6CanExecute);

            Message = "UILaunchApproachesViewModel says hello";

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

        #region Commands

        #region SayHello Command

        public ICommand SayHelloCommand { get; private set; }

        private void SayHello()
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.EventHandler) startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Message = $"Hello from {this.GetType()}";

            PublishStatusMessage(Message);

            if (Common.VNCLogging.EventHandler) Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private bool SayHelloCanExecute()
        {
            return true;
        }

        #endregion

        #region LaunchCylon1 Command

        public DelegateCommand LaunchCylon1Command { get; set; }

        public string LaunchCylon1Content { get; set; } = "LaunchCylon1";
        public string LaunchCylon1ToolTip { get; set; } = "LaunchCylon1 ToolTip";

        // Can get fancy and use Resources
        //public string LaunchCylon1Content { get; set; } = "ViewName_LaunchCylon1Content";
        //public string LaunchCylon1ToolTip { get; set; } = "ViewName_LaunchCylon1ContentToolTip";

        // Put these in Resource File
        //    <system:String x:Key="ViewName_LaunchCylon1Content">LaunchCylon1</system:String>
        //    <system:String x:Key="ViewName_LaunchCylon1ContentToolTip">LaunchCylon1 ToolTip</system:String>

        // If using CommandParameter, figure out TYPE and fix above
        //public void LaunchCylon1(TYPE value)

        public static WindowHost cylonHost = null;

        public void LaunchCylon1()
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.EventHandler) startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);
            // TODO(crhodes)
            // Do something amazing.
            Message = "Cool, you called LaunchCylon1";

            PublishStatusMessage(Message);

            if (cylonHost is null) cylonHost = new WindowHost(EventAggregator);

            cylonHost.DisplayUserControlInHost(
                "I am a Cylon loaded by name",
                Common.DEFAULT_WINDOW_WIDTH_SMALL, Common.DEFAULT_WINDOW_HEIGHT_SMALL,
                ShowWindowMode.Modeless_Show,
                "VNC.WPF.Presentation.Views.CylonEyeBall, VNC.WPF.Presentation");

            // Uncomment this if you are telling someone else to handle this

            // Common.EventAggregator.GetEvent<LaunchCylon1Event>().Publish();

            // May want EventArgs

            //  EventAggregator.GetEvent<LaunchCylon1Event>().Publish(
            //      new LaunchCylon1EventArgs()
            //      {
            //            Organization = _collectionMainViewModel.SelectedCollection.Organization,
            //            Process = _contextMainViewModel.Context.SelectedProcess
            //      });

            // Start Cut Three - Put this in PrismEvents

            // public class LaunchCylon1Event : PubSubEvent { }

            // End Cut Three

            // Start Cut Four - Put this in places that listen for event

            //Common.EventAggregator.GetEvent<LaunchCylon1Event>().Subscribe(LaunchCylon1);

            // End Cut Four

            if (Common.VNCLogging.EventHandler) Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // If using CommandParameter, figure out TYPE and fix above
        //public bool LaunchCylon1CanExecute(TYPE value)
        public bool LaunchCylon1CanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion

        #region LaunchCylon2 Command

        public DelegateCommand LaunchCylon2Command { get; set; }

        public string LaunchCylon2Content { get; set; } = "LaunchCylon2";
        public string LaunchCylon2ToolTip { get; set; } = "LaunchCylon2 ToolTip";

        // Can get fancy and use Resources
        //public string LaunchCylon2Content { get; set; } = "ViewName_LaunchCylon2Content";
        //public string LaunchCylon2ToolTip { get; set; } = "ViewName_LaunchCylon2ContentToolTip";

        // Put these in Resource File
        //    <system:String x:Key="ViewName_LaunchCylon2Content">LaunchCylon2</system:String>
        //    <system:String x:Key="ViewName_LaunchCylon2ContentToolTip">LaunchCylon2 ToolTip</system:String>

        // If using CommandParameter, figure out TYPE and fix above
        //public void LaunchCylon2(TYPE value)

        public static WindowHost cylonHost2 = null;

        public void LaunchCylon2()
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.EventHandler) startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);
            // TODO(crhodes)
            // Do something amazing.
            Message = "Cool, you called LaunchCylon2";

            PublishStatusMessage(Message);

            if (cylonHost2 is null) cylonHost2 = new WindowHost(EventAggregator);

            cylonHost2.DisplayUserControlInHost(
                "I am a Cylon loaded by type",
                Common.DEFAULT_WINDOW_WIDTH_SMALL, Common.DEFAULT_WINDOW_HEIGHT_SMALL,
                ShowWindowMode.Modeless_Show,
                new CylonEyeBall());

            // Uncomment this if you are telling someone else to handle this

            // Common.EventAggregator.GetEvent<LaunchCylon2Event>().Publish();

            // May want EventArgs

            //  EventAggregator.GetEvent<LaunchCylon2Event>().Publish(
            //      new LaunchCylon2EventArgs()
            //      {
            //            Organization = _collectionMainViewModel.SelectedCollection.Organization,
            //            Process = _contextMainViewModel.Context.SelectedProcess
            //      });

            // Start Cut Three - Put this in PrismEvents

            // public class LaunchCylon2Event : PubSubEvent { }

            // End Cut Three

            // Start Cut Four - Put this in places that listen for event

            //Common.EventAggregator.GetEvent<LaunchCylon2Event>().Subscribe(LaunchCylon2);

            // End Cut Four

            if (Common.VNCLogging.EventHandler) Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // If using CommandParameter, figure out TYPE and fix above
        //public bool LaunchCylon2CanExecute(TYPE value)
        public bool LaunchCylon2CanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion

        #region ThemedWindowHostModeless Command

        public DelegateCommand ThemedWindowHostModelessCommand { get; set; }

        // If using CommandParameter, figure out TYPE here and above
        // and remove above declaration
        //public DelegateCommand<TYPE> ThemedWindowHostModelessCommand { get; set; }
        //public TYPE ThemedWindowHostModelessCommandParameter;
        public string ThemedWindowHostModelessContent { get; set; } = "ThemedWindowHostModeless";
        public string ThemedWindowHostModelessToolTip { get; set; } = "ThemedWindowHostModeless ToolTip";

        // Can get fancy and use Resources
        //public string ThemedWindowHostModelessContent { get; set; } = "ViewName_ThemedWindowHostModelessContent";
        //public string ThemedWindowHostModelessToolTip { get; set; } = "ViewName_ThemedWindowHostModelessContentToolTip";

        // Put these in Resource File
        //    <system:String x:Key="ViewName_ThemedWindowHostModelessContent">ThemedWindowHostModeless</system:String>
        //    <system:String x:Key="ViewName_ThemedWindowHostModelessContentToolTip">ThemedWindowHostModeless ToolTip</system:String>

        // If using CommandParameter, figure out TYPE and fix above
        //public void ThemedWindowHostModeless(TYPE value)

        private DxThemedWindowHost themedWindowHostModeless = null;

        public void ThemedWindowHostModeless()
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.EventHandler) startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);
            // TODO(crhodes)
            // Do something amazing.
            Message = "Cool, you called ThemedWindowHostModeless";

            PublishStatusMessage(Message);

            if (themedWindowHostModeless is null) themedWindowHostModeless = new DxThemedWindowHost(EventAggregator);

            themedWindowHostModeless.DisplayUserControlInHost(
              "ThemedWindowHost (Modeless)",
              Common.DEFAULT_WINDOW_WIDTH_SMALL, Common.DEFAULT_WINDOW_HEIGHT_SMALL,
              ShowWindowMode.Modeless_Show,
              new ModelessUI());

            // Uncomment this if you are telling someone else to handle this

            // Common.EventAggregator.GetEvent<ThemedWindowHostModelessEvent>().Publish();

            // May want EventArgs

            //  EventAggregator.GetEvent<ThemedWindowHostModelessEvent>().Publish(
            //      new ThemedWindowHostModelessEventArgs()
            //      {
            //            Organization = _collectionMainViewModel.SelectedCollection.Organization,
            //            Process = _contextMainViewModel.Context.SelectedProcess
            //      });

            // Start Cut Three - Put this in PrismEvents

            // public class ThemedWindowHostModelessEvent : PubSubEvent { }

            // End Cut Three

            // Start Cut Four - Put this in places that listen for event

            //Common.EventAggregator.GetEvent<ThemedWindowHostModelessEvent>().Subscribe(ThemedWindowHostModeless);

            // End Cut Four

            if (Common.VNCLogging.EventHandler) Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // If using CommandParameter, figure out TYPE and fix above
        //public bool ThemedWindowHostModelessCanExecute(TYPE value)
        public bool ThemedWindowHostModelessCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion

        #region ThemedWindowHostModal Command

        public DelegateCommand ThemedWindowHostModalCommand { get; set; }

        // If using CommandParameter, figure out TYPE here and above
        // and remove above declaration
        //public DelegateCommand<TYPE> ThemedWindowHostModalCommand { get; set; }
        //public TYPE ThemedWindowHostModalCommandParameter;
        public string ThemedWindowHostModalContent { get; set; } = "ThemedWindowHostModal";
        public string ThemedWindowHostModalToolTip { get; set; } = "ThemedWindowHostModal ToolTip";

        // Can get fancy and use Resources
        //public string ThemedWindowHostModalContent { get; set; } = "ViewName_ThemedWindowHostModalContent";
        //public string ThemedWindowHostModalToolTip { get; set; } = "ViewName_ThemedWindowHostModalContentToolTip";

        // Put these in Resource File
        //    <system:String x:Key="ViewName_ThemedWindowHostModalContent">ThemedWindowHostModal</system:String>
        //    <system:String x:Key="ViewName_ThemedWindowHostModalContentToolTip">ThemedWindowHostModal ToolTip</system:String>

        // If using CommandParameter, figure out TYPE and fix above
        //public void ThemedWindowHostModal(TYPE value)

        private DxThemedWindowHost themedWindowHostModal = null;

        public void ThemedWindowHostModal()
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.EventHandler) startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);
            // TODO(crhodes)
            // Do something amazing.
            Message = "Cool, you called ThemedWindowHostModal";

            PublishStatusMessage(Message);

            if (themedWindowHostModal is null) themedWindowHostModal = new DxThemedWindowHost(EventAggregator);

            themedWindowHostModal.DisplayUserControlInHost(
                "ThemedWindowHost (Modal)",
                Common.DEFAULT_WINDOW_WIDTH_SMALL, Common.DEFAULT_WINDOW_HEIGHT_SMALL,
                ShowWindowMode.Modal_ShowDialog,
                new ModalUI());

            // Uncomment this if you are telling someone else to handle this

            // Common.EventAggregator.GetEvent<ThemedWindowHostModalEvent>().Publish();

            // May want EventArgs

            //  EventAggregator.GetEvent<ThemedWindowHostModalEvent>().Publish(
            //      new ThemedWindowHostModalEventArgs()
            //      {
            //            Organization = _collectionMainViewModel.SelectedCollection.Organization,
            //            Process = _contextMainViewModel.Context.SelectedProcess
            //      });

            // Start Cut Three - Put this in PrismEvents

            // public class ThemedWindowHostModalEvent : PubSubEvent { }

            // End Cut Three

            // Start Cut Four - Put this in places that listen for event

            //Common.EventAggregator.GetEvent<ThemedWindowHostModalEvent>().Subscribe(ThemedWindowHostModal);

            // End Cut Four

            if (Common.VNCLogging.EventHandler) Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // If using CommandParameter, figure out TYPE and fix above
        //public bool ThemedWindowHostModalCanExecute(TYPE value)
        public bool ThemedWindowHostModalCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion

        #region DXLayoutControl Command

        public DelegateCommand DXLayoutControlCommand { get; set; }

        public string DXLayoutControlContent { get; set; } = "DXLayoutControl";
        public string DXLayoutControlToolTip { get; set; } = "DXLayoutControl ToolTip";

        // Can get fancy and use Resources
        //public string DXLayoutControlContent { get; set; } = "ViewName_DXLayoutControlContent";
        //public string DXLayoutControlToolTip { get; set; } = "ViewName_DXLayoutControlContentToolTip";

        // Put these in Resource File
        //    <system:String x:Key="ViewName_DXLayoutControlContent">DXLayoutControl</system:String>
        //    <system:String x:Key="ViewName_DXLayoutControlContentToolTip">DXLayoutControl ToolTip</system:String>

        // If using CommandParameter, figure out TYPE and fix above
        //public void DXLayoutControl(TYPE value)

        private DxThemedWindowHost dxLayoutControlHost = null;

        public void DXLayoutControl()
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.EventHandler) startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);
            // TODO(crhodes)
            // Do something amazing.
            Message = "Cool, you called DXLayoutControl";

            PublishStatusMessage(Message);

            if (dxLayoutControlHost is null) dxLayoutControlHost = new DxThemedWindowHost(EventAggregator);

            dxLayoutControlHost.DisplayUserControlInHost(
                "DxLayoutControl Test",
                Common.DEFAULT_WINDOW_WIDTH, Common.DEFAULT_WINDOW_HEIGHT,
                ShowWindowMode.Modeless_Show,
                new DxLayoutControl());

            // Uncomment this if you are telling someone else to handle this

            // Common.EventAggregator.GetEvent<DXLayoutControlEvent>().Publish();

            // May want EventArgs

            //  EventAggregator.GetEvent<DXLayoutControlEvent>().Publish(
            //      new DXLayoutControlEventArgs()
            //      {
            //            Organization = _collectionMainViewModel.SelectedCollection.Organization,
            //            Process = _contextMainViewModel.Context.SelectedProcess
            //      });

            // Start Cut Three - Put this in PrismEvents

            // public class DXLayoutControlEvent : PubSubEvent { }

            // End Cut Three

            // Start Cut Four - Put this in places that listen for event

            //Common.EventAggregator.GetEvent<DXLayoutControlEvent>().Subscribe(DXLayoutControl);

            // End Cut Four

            if (Common.VNCLogging.EventHandler) Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // If using CommandParameter, figure out TYPE and fix above
        //public bool DXLayoutControlCanExecute(TYPE value)
        public bool DXLayoutControlCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion

        #region DXDockLayoutControl Command

        public DelegateCommand DXDockLayoutControlCommand { get; set; }

        public string DXDockLayoutControlContent { get; set; } = "DXDockLayoutControl";
        public string DXDockLayoutControlToolTip { get; set; } = "DXDockLayoutControl ToolTip";

        // Can get fancy and use Resources
        //public string DXDockLayoutControlContent { get; set; } = "ViewName_DXDockLayoutControlContent";
        //public string DXDockLayoutControlToolTip { get; set; } = "ViewName_DXDockLayoutControlContentToolTip";

        // Put these in Resource File
        //    <system:String x:Key="ViewName_DXDockLayoutControlContent">DXDockLayoutControl</system:String>
        //    <system:String x:Key="ViewName_DXDockLayoutControlContentToolTip">DXDockLayoutControl ToolTip</system:String>

        // If using CommandParameter, figure out TYPE and fix above
        //public void DXDockLayoutControl(TYPE value)

        private DxThemedWindowHost dxDockLayoutControlHost = null;

        public void DXDockLayoutControl()
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.EventHandler) startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);
            // TODO(crhodes)
            // Do something amazing.
            Message = "Cool, you called DXDockLayoutControl";

            PublishStatusMessage(Message);

            if (dxDockLayoutControlHost is null) dxDockLayoutControlHost = new DxThemedWindowHost(EventAggregator);

            dxDockLayoutControlHost.DisplayUserControlInHost(
                "DxDockLayoutControl Test",
                Common.DEFAULT_WINDOW_WIDTH, Common.DEFAULT_WINDOW_HEIGHT,
                ShowWindowMode.Modeless_Show,
                new DxDockLayoutControl());

            // Uncomment this if you are telling someone else to handle this

            // Common.EventAggregator.GetEvent<DXDockLayoutControlEvent>().Publish();

            // May want EventArgs

            //  EventAggregator.GetEvent<DXDockLayoutControlEvent>().Publish(
            //      new DXDockLayoutControlEventArgs()
            //      {
            //            Organization = _collectionMainViewModel.SelectedCollection.Organization,
            //            Process = _contextMainViewModel.Context.SelectedProcess
            //      });

            // Start Cut Three - Put this in PrismEvents

            // public class DXDockLayoutControlEvent : PubSubEvent { }

            // End Cut Three

            // Start Cut Four - Put this in places that listen for event

            //Common.EventAggregator.GetEvent<DXDockLayoutControlEvent>().Subscribe(DXDockLayoutControl);

            // End Cut Four

            if (Common.VNCLogging.EventHandler) Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // If using CommandParameter, figure out TYPE and fix above
        //public bool DXDockLayoutControlCanExecute(TYPE value)
        public bool DXDockLayoutControlCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion

        #region DXDockLayoutManager Command

        public DelegateCommand DXDockLayoutManagerCommand { get; set; }

        // Can get fancy and use Resources
        //public string DXDockLayoutManagerContent { get; set; } = "ViewName_DXDockLayoutManagerContent";
        //public string DXDockLayoutManagerToolTip { get; set; } = "ViewName_DXDockLayoutManagerContentToolTip";

        // Put these in Resource File
        //    <system:String x:Key="ViewName_DXDockLayoutManagerContent">DXDockLayoutManager</system:String>
        //    <system:String x:Key="ViewName_DXDockLayoutManagerContentToolTip">DXDockLayoutManager ToolTip</system:String>

        // If using CommandParameter, figure out TYPE and fix above
        //public void DXDockLayoutManager(TYPE value)

        private DxThemedWindowHost dxDockLayoutManagerControlHost = null;

        public void DXDockLayoutManager()
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.EventHandler) startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);
            // TODO(crhodes)
            // Do something amazing.
            Message = "Cool, you called DXDockLayoutManager";

            PublishStatusMessage(Message);

            if (dxDockLayoutManagerControlHost is null) dxDockLayoutManagerControlHost = new DxThemedWindowHost(EventAggregator);

            dxDockLayoutManagerControlHost.DisplayUserControlInHost(
                "DxDocLayoutManagerControl Test",
                Common.DEFAULT_WINDOW_WIDTH, Common.DEFAULT_WINDOW_HEIGHT,
                ShowWindowMode.Modeless_Show,
                new DxDockLayoutManagerControl());

            // Uncomment this if you are telling someone else to handle this

            // Common.EventAggregator.GetEvent<DXDockLayoutManagerEvent>().Publish();

            // May want EventArgs

            //  EventAggregator.GetEvent<DXDockLayoutManagerEvent>().Publish(
            //      new DXDockLayoutManagerEventArgs()
            //      {
            //            Organization = _collectionMainViewModel.SelectedCollection.Organization,
            //            Process = _contextMainViewModel.Context.SelectedProcess
            //      });

            // Start Cut Three - Put this in PrismEvents

            // public class DXDockLayoutManagerEvent : PubSubEvent { }

            // End Cut Three

            // Start Cut Four - Put this in places that listen for event

            //Common.EventAggregator.GetEvent<DXDockLayoutManagerEvent>().Subscribe(DXDockLayoutManager);

            // End Cut Four

            if (Common.VNCLogging.EventHandler) Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // If using CommandParameter, figure out TYPE and fix above
        //public bool DXDockLayoutManagerCanExecute(TYPE value)
        public bool DXDockLayoutManagerCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion

        #region MVVM Commands

        #region MVVM1 Command

        public DelegateCommand MVVM1Command { get; set; }

        public string DXDockLayoutManagerContent { get; set; } = "DXDockLayoutManager";
        public string DXDockLayoutManagerToolTip { get; set; } = "DXDockLayoutManager ToolTip";

        public string MVVM1Content { get; set; } = "MVVM VA VM1st";
        public string MVVM1ToolTip { get; set; } = "MVVM1 ToolTip";

        // Can get fancy and use Resources
        //public string MVVM1Content { get; set; } = "ViewName_MVVM1Content";
        //public string MVVM1ToolTip { get; set; } = "ViewName_MVVM1ContentToolTip";

        // Put these in Resource File
        //    <system:String x:Key="ViewName_MVVM1Content">MVVM1</system:String>
        //    <system:String x:Key="ViewName_MVVM1ContentToolTip">MVVM1 ToolTip</system:String>

        // If using CommandParameter, figure out TYPE and fix above
        //public void MVVM1(TYPE value)

        public static DxThemedWindowHost vncMVVM_VAVM1st_Host = null;
        public void MVVM1()
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.EventHandler) startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);
            // TODO(crhodes)
            // Do something amazing.
            Message = "Cool, you called MVVM1";

            PublishStatusMessage(Message);

            // FIX(crhodes)
            // Wire things up ourselves - ViewModel First - with a little help from DI.
            // We don't have a Constructor that takes 3 arguments

            // DxThemedWindowHost.DisplayUserControlInHost(ref vncMVVM_VAVM1st_Host,
            // "MVVM ViewAViewModel First (ViewModel is passed new ViewA) - By Hand",
            // Common.DEFAULT_WINDOW_WIDTH, Common.DEFAULT_WINDOW_HEIGHT,
            // ShowWindowMode.Modeless_Show,
            // new ViewAViewModel(
            // new ViewA(),
            // (IEventAggregator)Common.Container.Resolve(typeof(EventAggregator)),
            // (DialogService)Common.Container.Resolve(typeof(DialogService))
            // )
            // );

            // Uncomment this if you are telling someone else to handle this

            // Common.EventAggregator.GetEvent<MVVM1Event>().Publish();

            // May want EventArgs

            //  EventAggregator.GetEvent<MVVM1Event>().Publish(
            //      new MVVM1EventArgs()
            //      {
            //            Organization = _collectionMainViewModel.SelectedCollection.Organization,
            //            Process = _contextMainViewModel.Context.SelectedProcess
            //      });

            // Start Cut Three - Put this in PrismEvents

            // public class MVVM1Event : PubSubEvent { }

            // End Cut Three

            // Start Cut Four - Put this in places that listen for event

            //Common.EventAggregator.GetEvent<MVVM1Event>().Subscribe(MVVM1);

            // End Cut Four

            if (Common.VNCLogging.EventHandler) Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // If using CommandParameter, figure out TYPE and fix above
        //public bool MVVM1CanExecute(TYPE value)
        public bool MVVM1CanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion

        #region MVVM2 Command

        public DelegateCommand MVVM2Command { get; set; }

        public string MVVM2Content { get; set; } = "MVVM2";
        public string MVVM2ToolTip { get; set; } = "MVVM2 ToolTip";

        // Can get fancy and use Resources
        //public string MVVM2Content { get; set; } = "ViewName_MVVM2Content";
        //public string MVVM2ToolTip { get; set; } = "ViewName_MVVM2ContentToolTip";

        // Put these in Resource File
        //    <system:String x:Key="ViewName_MVVM2Content">MVVM2</system:String>
        //    <system:String x:Key="ViewName_MVVM2ContentToolTip">MVVM2 ToolTip</system:String>

        // If using CommandParameter, figure out TYPE and fix above
        //public void MVVM2(TYPE value)
        public void MVVM2()
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.EventHandler) startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);
            // TODO(crhodes)
            // Do something amazing.
            Message = "Cool, you called MVVM2";

            PublishStatusMessage(Message);

            // Uncomment this if you are telling someone else to handle this

            // Common.EventAggregator.GetEvent<MVVM2Event>().Publish();

            // May want EventArgs

            //  EventAggregator.GetEvent<MVVM2Event>().Publish(
            //      new MVVM2EventArgs()
            //      {
            //            Organization = _collectionMainViewModel.SelectedCollection.Organization,
            //            Process = _contextMainViewModel.Context.SelectedProcess
            //      });

            // Start Cut Three - Put this in PrismEvents

            // public class MVVM2Event : PubSubEvent { }

            // End Cut Three

            // Start Cut Four - Put this in places that listen for event

            //Common.EventAggregator.GetEvent<MVVM2Event>().Subscribe(MVVM2);

            // End Cut Four

            if (Common.VNCLogging.EventHandler) Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // If using CommandParameter, figure out TYPE and fix above
        //public bool MVVM2CanExecute(TYPE value)
        public bool MVVM2CanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion

        #region MVVM3 Command

        public DelegateCommand MVVM3Command { get; set; }

        public string MVVM3Content { get; set; } = "MVVM3";
        public string MVVM3ToolTip { get; set; } = "MVVM3 ToolTip";

        // Can get fancy and use Resources
        //public string MVVM3Content { get; set; } = "ViewName_MVVM3Content";
        //public string MVVM3ToolTip { get; set; } = "ViewName_MVVM3ContentToolTip";

        // Put these in Resource File
        //    <system:String x:Key="ViewName_MVVM3Content">MVVM3</system:String>
        //    <system:String x:Key="ViewName_MVVM3ContentToolTip">MVVM3 ToolTip</system:String>

        // If using CommandParameter, figure out TYPE and fix above
        //public void MVVM3(TYPE value)
        public void MVVM3()
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.EventHandler) startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);
            // TODO(crhodes)
            // Do something amazing.
            Message = "Cool, you called MVVM3";

            PublishStatusMessage(Message);

            // Uncomment this if you are telling someone else to handle this

            // Common.EventAggregator.GetEvent<MVVM3Event>().Publish();

            // May want EventArgs

            //  EventAggregator.GetEvent<MVVM3Event>().Publish(
            //      new MVVM3EventArgs()
            //      {
            //            Organization = _collectionMainViewModel.SelectedCollection.Organization,
            //            Process = _contextMainViewModel.Context.SelectedProcess
            //      });

            // Start Cut Three - Put this in PrismEvents

            // public class MVVM3Event : PubSubEvent { }

            // End Cut Three

            // Start Cut Four - Put this in places that listen for event

            //Common.EventAggregator.GetEvent<MVVM3Event>().Subscribe(MVVM3);

            // End Cut Four

            if (Common.VNCLogging.EventHandler) Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // If using CommandParameter, figure out TYPE and fix above
        //public bool MVVM3CanExecute(TYPE value)
        public bool MVVM3CanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion

        #region MVVM4 Command

        public DelegateCommand MVVM4Command { get; set; }

        public string MVVM4Content { get; set; } = "MVVM4";
        public string MVVM4ToolTip { get; set; } = "MVVM4 ToolTip";

        // Can get fancy and use Resources
        //public string MVVM4Content { get; set; } = "ViewName_MVVM4Content";
        //public string MVVM4ToolTip { get; set; } = "ViewName_MVVM4ContentToolTip";

        // Put these in Resource File
        //    <system:String x:Key="ViewName_MVVM4Content">MVVM4</system:String>
        //    <system:String x:Key="ViewName_MVVM4ContentToolTip">MVVM4 ToolTip</system:String>

        // If using CommandParameter, figure out TYPE and fix above
        //public void MVVM4(TYPE value)
        public void MVVM4()
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.EventHandler) startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);
            // TODO(crhodes)
            // Do something amazing.
            Message = "Cool, you called MVVM4";

            PublishStatusMessage(Message);

            // Uncomment this if you are telling someone else to handle this

            // Common.EventAggregator.GetEvent<MVVM4Event>().Publish();

            // May want EventArgs

            //  EventAggregator.GetEvent<MVVM4Event>().Publish(
            //      new MVVM4EventArgs()
            //      {
            //            Organization = _collectionMainViewModel.SelectedCollection.Organization,
            //            Process = _contextMainViewModel.Context.SelectedProcess
            //      });

            // Start Cut Three - Put this in PrismEvents

            // public class MVVM4Event : PubSubEvent { }

            // End Cut Three

            // Start Cut Four - Put this in places that listen for event

            //Common.EventAggregator.GetEvent<MVVM4Event>().Subscribe(MVVM4);

            // End Cut Four

            if (Common.VNCLogging.EventHandler) Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // If using CommandParameter, figure out TYPE and fix above
        //public bool MVVM4CanExecute(TYPE value)
        public bool MVVM4CanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion

        #region MVVM5 Command

        public DelegateCommand MVVM5Command { get; set; }

        public string MVVM5Content { get; set; } = "MVVM5";
        public string MVVM5ToolTip { get; set; } = "MVVM5 ToolTip";

        // Can get fancy and use Resources
        //public string MVVM5Content { get; set; } = "ViewName_MVVM5Content";
        //public string MVVM5ToolTip { get; set; } = "ViewName_MVVM5ContentToolTip";

        // Put these in Resource File
        //    <system:String x:Key="ViewName_MVVM5Content">MVVM5</system:String>
        //    <system:String x:Key="ViewName_MVVM5ContentToolTip">MVVM5 ToolTip</system:String>

        // If using CommandParameter, figure out TYPE and fix above
        //public void MVVM5(TYPE value)
        public void MVVM5()
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.EventHandler) startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);
            // TODO(crhodes)
            // Do something amazing.
            Message = "Cool, you called MVVM5";

            PublishStatusMessage(Message);

            // Uncomment this if you are telling someone else to handle this

            // Common.EventAggregator.GetEvent<MVVM5Event>().Publish();

            // May want EventArgs

            //  EventAggregator.GetEvent<MVVM5Event>().Publish(
            //      new MVVM5EventArgs()
            //      {
            //            Organization = _collectionMainViewModel.SelectedCollection.Organization,
            //            Process = _contextMainViewModel.Context.SelectedProcess
            //      });

            // Start Cut Three - Put this in PrismEvents

            // public class MVVM5Event : PubSubEvent { }

            // End Cut Three

            // Start Cut Four - Put this in places that listen for event

            //Common.EventAggregator.GetEvent<MVVM5Event>().Subscribe(MVVM5);

            // End Cut Four

            if (Common.VNCLogging.EventHandler) Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // If using CommandParameter, figure out TYPE and fix above
        //public bool MVVM5CanExecute(TYPE value)
        public bool MVVM5CanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion

        #region MVVM6 Command

        public DelegateCommand MVVM6Command { get; set; }

        public string MVVM6Content { get; set; } = "MVVM6";
        public string MVVM6ToolTip { get; set; } = "MVVM6 ToolTip";

        // Can get fancy and use Resources
        //public string MVVM6Content { get; set; } = "ViewName_MVVM6Content";
        //public string MVVM6ToolTip { get; set; } = "ViewName_MVVM6ContentToolTip";

        // Put these in Resource File
        //    <system:String x:Key="ViewName_MVVM6Content">MVVM6</system:String>
        //    <system:String x:Key="ViewName_MVVM6ContentToolTip">MVVM6 ToolTip</system:String>

        // If using CommandParameter, figure out TYPE and fix above
        //public void MVVM6(TYPE value)
        public void MVVM6()
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.EventHandler) startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);
            // TODO(crhodes)
            // Do something amazing.
            Message = "Cool, you called MVVM6";

            PublishStatusMessage(Message);

            // Uncomment this if you are telling someone else to handle this

            // Common.EventAggregator.GetEvent<MVVM6Event>().Publish();

            // May want EventArgs

            //  EventAggregator.GetEvent<MVVM6Event>().Publish(
            //      new MVVM6EventArgs()
            //      {
            //            Organization = _collectionMainViewModel.SelectedCollection.Organization,
            //            Process = _contextMainViewModel.Context.SelectedProcess
            //      });

            // Start Cut Three - Put this in PrismEvents

            // public class MVVM6Event : PubSubEvent { }

            // End Cut Three

            // Start Cut Four - Put this in places that listen for event

            //Common.EventAggregator.GetEvent<MVVM6Event>().Subscribe(MVVM6);

            // End Cut Four

            if (Common.VNCLogging.EventHandler) Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // If using CommandParameter, figure out TYPE and fix above
        //public bool MVVM6CanExecute(TYPE value)
        public bool MVVM6CanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion

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
