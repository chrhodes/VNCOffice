using System;

using Prism.Ioc;
using Prism.Modularity;
using Prism.Regions;

using SimpleWPFApp.Core;

#if VNCEF
using SimpleWPFApp.DomainServices;
#endif

using SimpleWPFApp.Presentation.Views;
using SimpleWPFApp.Presentation.ViewModels;

using Unity;

using VNC;
using VNC.WPF.Presentation.Views;
using VNC.WPF.Presentation.ViewModels;

namespace SimpleWPFApp
{
    public class SimpleWPFAppModule : IModule
    {
        private readonly IRegionManager _regionManager;

        // 01

        public SimpleWPFAppModule(IRegionManager regionManager)
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.Constructor) startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            _regionManager = regionManager;

            if (Common.VNCLogging.Constructor) Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // 02

        public void RegisterTypes(IContainerRegistry containerRegistry)
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.ModuleInitialize) startTicks = Log.MODULE_INITIALIZE("Enter", Common.LOG_CATEGORY);

            // TODO(crhodes)
            // Maybe the Ribbon, Main, StatusBar should be move back to App.xaml.cs
            // and the App Specific stuff left here

            // TODO(crhodes)
            // This is where you pick the style of what gets loaded in the Shell.
            // The initial shell is chosen in App.xaml.cs CreateShell()

            // If you are using the Ribbon Shell and the RibbonRegion

            containerRegistry.RegisterSingleton<IRibbonViewModel, RibbonViewModel>();
            containerRegistry.RegisterSingleton<IRibbon, Ribbon>();

            // If you are using the Shell and the RibbonRegion

            //containerRegistry.RegisterSingleton<IRibbonViewModel, ShellRibbonViewModel>();
            //containerRegistry.RegisterSingleton<IRibbon, ShellRibbon>();

            // Pick one of these for the MainRegion

            // TODO(crhodes)
            // Learn how to dynamically switch this while application running

            //containerRegistry.Register<IMain, Main>();
            containerRegistry.Register<IMain, MainDxLayoutControl>();
            //containerRegistry.Register<IMain, MainDxDockLayoutManager>();

            // NOTE(crhodes)
            // Note sure why the above works as have not registered a ViewModel.
            // Why does MainDxLayoutControl(MainDxLayoutControlViewModel viewModel) get called.
            //
            // Oh, I see.
            // MainDxLayoutControl(MainDxLayoutControlViewModel viewModel)
            // Does not need DI as the ViewModel is a TYPE

            containerRegistry.RegisterSingleton<IStatusBarViewModel, StatusBarViewModel>();
            containerRegistry.RegisterSingleton<StatusBar, StatusBar>();

#if VNCEF
            // NOTE(crhodes)
            // Registering CombinedNavigation as SingleTon solves the two copies problem

            containerRegistry.RegisterSingleton<ICombinedMainViewModel, CombinedMainViewModel>();
            containerRegistry.RegisterSingleton<ICombinedMain, CombinedMain>();

            containerRegistry.RegisterSingleton<ICombinedNavigationViewModel, CombinedNavigationViewModel>();
            containerRegistry.RegisterSingleton<ICombinedNavigation, CombinedNavigation>();

            // NOTE(crhodes)
            // Seems like these belong in Module

            containerRegistry.Register<IDetailViewModel, DetailViewModel>();
            containerRegistry.RegisterSingleton<IDetail, Detail>();

            containerRegistry.RegisterSingleton<IDataService, DataService>();
            containerRegistry.RegisterSingleton<ILookupDataService, LookupDataService>();
#endif

            // NOTE(crhodes)
            // Can register a type

            //containerRegistry.Register<ViewA>();
            //containerRegistry.Register<ViewB>();
            //containerRegistry.Register<ViewC>();
            //containerRegistry.Register<ViewD>();

            //// NOTE(crhodes)
            //// Can register a type against an Interface.
            //// See OnInitialized, supra

            //containerRegistry.Register<IViewABCD, ViewABCD>();

            //// NOTE(crhodes)
            //// If do not register as Singleton, get a ViewABCDViewModel for each View above
            //containerRegistry.RegisterSingleton<IViewABCDViewModel, ViewABCDViewModel>();
            ////containerRegistry.Register<IViewABCDViewModel, ViewABCDViewModel>();

            //containerRegistry.Register<IUILaunchApproachesViewModel, UILaunchApproachesViewModel>();
            //containerRegistry.Register<UILaunchApproaches>();

            // NOTE(crhodes)
            // Observations on wiring Views and ViewModels

            // ------------------------------------------------------
            //
            // If only put ViewDiscoveryAndInjection in OnInitialized
            // and AutoWireViewModel=False
            //
            // ViewDiscoveryAndInjection() gets called and View is not wired to ViewModel
            //
            // ------------------------------------------------------

            // ++++++++++++++++++++++++++++++++++++++++++++++++++++++
            //
            // If Set AutoWireViewModel=True in Xaml,
            // ViewDiscoveryAndInjection(IViewDiscoveryAndInjectionViewModel viewModel) gets called
            //
            // and View is wired to ViewModel
            //
            // ++++++++++++++++++++++++++++++++++++++++++++++++++++++

            // ------------------------------------------------------
            //
            // If only Register ViewDiscoveryAndInjection
            // either directly or against Interface

            //containerRegistry.Register<ViewDiscoveryAndInjection>();
            //containerRegistry.Register<IViewDiscoveryAndInjection, ViewDiscoveryAndInjection>();

            // ViewDiscoveryAndInjection() gets called
            // and View is not wired to ViewModel
            //
            // Why doesn't this work as IMain works, infra.

            // Oh, I see.
            // public ViewDiscoveryAndInjection(IViewDiscoveryAndInjectionViewModel viewModel)
            // Needs DI as the viewModel is an Interface
            //
            // ------------------------------------------------------

            // ------------------------------------------------------
            //
            // If both are registered, order doesn't matter

            //containerRegistry.Register<ViewDiscoveryAndInjection>();
            //containerRegistry.Register<ViewDiscoveryAndInjectionViewModel>();

            // ViewDiscoveryAndInjection() gets called
            // and View is not wired to ViewModel
            //
            // ------------------------------------------------------

            // ------------------------------------------------------
            //
            // If only Register ViewDiscoveryAndInjectionViewModel

            //containerRegistry.Register<ViewDiscoveryAndInjectionViewModel>();

            // ViewDiscoveryAndInjection() gets called
            // and View is not wired to ViewModel
            //
            // ------------------------------------------------------

            // ++++++++++++++++++++++++++++++++++++++++++++++++++++++
            //
            // But, if register ViewDiscoveryAndInjectionViewModel
            // against an Interface

            //containerRegistry.Register<IViewDiscoveryViewModel, ViewDiscoveryViewModel>();
            //containerRegistry.Register<IViewInjectionViewModel, ViewInjectionViewModel>();
            //containerRegistry.Register<IRegionNavigationViewModel, RegionNavigationViewModel>();

            // ViewDiscoveryAndInjection(IViewDiscoveryAndInjectionViewModel viewModel) gets called
            // and View is wired to ViewModel
            //
            // ++++++++++++++++++++++++++++++++++++++++++++++++++++++

            // Figure out how to use one Type

            //containerRegistry.Register<IFriendLookupDataService, LookupDataService>();
            //containerRegistry.Register<IProgrammingLanguageLookupDataService, LookupDataService>();
            //containerRegistry.Register<IMeetingLookupDataService, LookupDataService>();

            // These are for RegionNavigation

            //containerRegistry.RegisterForNavigation(typeof(UIOne), "uione");
            //// containerRegistry.RegisterForNavigation<UITwo>();    // name defaults to UITwo
            //containerRegistry.RegisterForNavigation<UITwo>("uitwo");
            //containerRegistry.RegisterForNavigation<UIThree>("uithree");

            // NOTE(crhodes)
            // Can also (optionally) register a ViewModel with View

            //containerRegistry.RegisterForNavigation<UIFour, UIFourViewModel>("uifour");
            //containerRegistry.RegisterForNavigation<UIFive, UIFiveViewModel>("uifive");

            // Since these do not have a ViewModel and only a parameterless constructor()
            // their datacontext is the parent control.  Probably don't want this

            //containerRegistry.RegisterForNavigation<UIOne_Beta>("uionebeta");
            //containerRegistry.RegisterForNavigation<UITwo_Beta>("uitwobeta");
            //containerRegistry.RegisterForNavigation<UIThree_Beta>("uithreebeta");

            // Can specify ViewModel to use

            //containerRegistry.RegisterForNavigation<UIFour_Beta, UIFourViewModel>("uifourbeta");
            //containerRegistry.RegisterForNavigation<UIFive_Beta, UIFiveViewModel>("uifivebeta");

            //containerRegistry.Register<IMultiStepProcessViewModel, MultiStepProcessViewModel>();
            //containerRegistry.Register<IMultiStepProcess, MultiStepProcess>();

            // NOTE(crhodes)
            // If we don't register this, first step comes up but cannot navigate
            // Need to register against interface to get the parameterized constructor of StepX(viewModel) to be called

            //containerRegistry.RegisterSingleton<IStepABCDEViewModel, StepABCDEViewModel>();
            //containerRegistry.RegisterSingleton<StepABCDEViewModel>();

            ////containerRegistry.RegisterSingleton<ICatDetailMVViewModel, CatDetailMVViewModel>();
            //containerRegistry.Register<ICatDetailMVViewModel, CatDetailMVViewModel>();

            //containerRegistry.RegisterSingleton<ICatDetailMV, StepA>();

            // NOTE(crhodes)
            // Need to Register for Navigation so RequestNavigate works.

            //containerRegistry.RegisterForNavigation<StepA, StepABCDEViewModel>("uistepa");
            //containerRegistry.RegisterForNavigation<StepB, StepABCDEViewModel>("uistepb");
            //containerRegistry.RegisterForNavigation<StepC, StepABCDEViewModel>("uistepc");
            //containerRegistry.RegisterForNavigation<StepD, StepABCDEViewModel>("uistepd");
            //containerRegistry.RegisterForNavigation<StepE, StepABCDEViewModel>("uistepe");

            //containerRegistry.RegisterForNavigation<DetailMVA, DetailMVViewModel>("uicatdetaila");
            //containerRegistry.RegisterForNavigation<DetailMVB, DetailMVViewModel>("uicatdetailb");
            //containerRegistry.RegisterForNavigation<DetailMVC, DetailMVViewModel>("uicatdetailc");
            //containerRegistry.RegisterForNavigation<DetailMVD, DetailMVViewModel>("uicatdetaild");
            //containerRegistry.RegisterForNavigation<DetailMVE, DetailMVViewModel>("uicatdetaile");

            //containerRegistry.RegisterSingleton<ICatDetailMV, CatDetailMVA>();

            if (Common.VNCLogging.ModuleInitialize) Log.MODULE_INITIALIZE("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // 03

        public void OnInitialized(IContainerProvider containerProvider)
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.ModuleInitialize) startTicks = Log.MODULE_INITIALIZE("Enter", Common.LOG_CATEGORY);

            // NOTE(crhodes)
            // using typeof(TYPE) calls constructor
            // using typeof(ITYPE) resolves type (see RegisterTypes)
            // but only wires viewmodel if IViewModel is used

            // HACK(crhodes)
            // Save the current RegionManager so we can use it in other places that do not get passed it

            //Common.DefaultRegionManager = _regionManager;
            //var rc = _regionManager.Regions.Count();

            // NOTE(crhodes)
            // RegisterViewWithRegion is ViewDiscovery
            // Options are ViewInjection or Navigation

            _regionManager.RegisterViewWithRegion(RegionNames.RibbonRegion, typeof(IRibbon));
            _regionManager.RegisterViewWithRegion(RegionNames.MainRegion, typeof(IMain));
            _regionManager.RegisterViewWithRegion(RegionNames.StatusBarRegion, typeof(StatusBar));

            _regionManager.RegisterViewWithRegion(RegionNames.VNCLoggingConfigRegion, typeof(VNC.WPF.Presentation.Dx.Views.VNCLoggingConfig));
            _regionManager.RegisterViewWithRegion(RegionNames.VNCCoreLoggingConfigRegion, typeof(VNC.WPF.Presentation.Dx.Views.VNCCoreLoggingConfig));

#if VNCEF
            // This loads CombinedMain into the IMain configured
            // in RegisterTypes (MainDxLayout or MainDxDocLayoutManager), supra

            _regionManager.RegisterViewWithRegion(RegionNames.CombinedMainRegion, typeof(ICombinedMain));

            _regionManager.RegisterViewWithRegion(RegionNames.CombinedNavigationRegion, typeof(ICombinedNavigation));

#endif
            // This is for Main

            //_regionManager.RegisterViewWithRegion(RegionNames.ContentRegion, typeof(ViewA));

            // Can register a type

            //_regionManager.RegisterViewWithRegion(RegionNames.ViewARegion, typeof(ViewA));
            //_regionManager.RegisterViewWithRegion(RegionNames.ViewBRegion, typeof(ViewB));
            //_regionManager.RegisterViewWithRegion(RegionNames.ViewCRegion, typeof(ViewC));
            //_regionManager.RegisterViewWithRegion(RegionNames.ViewDRegion, typeof(ViewD));

            // Can register an interface.  See registerTypes Infra

            //_regionManager.RegisterViewWithRegion(RegionNames.ViewABCDRegion, typeof(IViewABCD));

            //_regionManager.RegisterViewWithRegion(RegionNames.UILaunchApproaches, typeof(UILaunchApproaches));

            //_regionManager.RegisterViewWithRegion(RegionNames.ViewInjection, typeof(ViewInjection));
            //_regionManager.RegisterViewWithRegion(RegionNames.ViewDiscovery, typeof(ViewDiscovery));
            //_regionManager.RegisterViewWithRegion(RegionNames.RegionNavigation, typeof(RegionNavigation));
            //_regionManager.RegisterViewWithRegion(RegionNames.MultiStepProcess, typeof(MultiStepProcess));

            //_regionManager.RegisterViewWithRegion(RegionNames.MultiStepProcessViewMV, typeof(CatDetailMVA));

            //var ok = _regionManager.Regions.ContainsRegionWithName(RegionNames.MultiStepProcessViewMV);

            if (Common.VNCLogging.ModuleInitialize) Log.MODULE_INITIALIZE("Exit", Common.LOG_CATEGORY, startTicks);
        }
    }
}
