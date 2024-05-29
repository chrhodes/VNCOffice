using System;

using Explore.Core;
using Explore.DomainServices;
using Explore.Presentation.ViewModels;
using Explore.Presentation.Views;

using Prism.Ioc;
using Prism.Modularity;
using Prism.Regions;

using Unity;

using VNC;
using VNC.Core.Presentation.ViewModels;
using VNC.Core.Presentation.Views;

namespace Explore
{
    public class ExploreModule : IModule
    {
        private readonly IRegionManager _regionManager;

        // 01

        public ExploreModule(IRegionManager regionManager)
        {
            Int64 startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            _regionManager = regionManager;

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // 02

        public void RegisterTypes(IContainerRegistry containerRegistry)
        {
            Int64 startTicks = Log.MODULE("Enter", Common.LOG_CATEGORY);

            containerRegistry.Register<IDoorDetailViewModel, DoorDetailViewModel>();
            containerRegistry.RegisterSingleton<IDoorDetail, DoorDetail>();

            containerRegistry.Register<IViewABCViewModel, ViewABCViewModel>();
            containerRegistry.RegisterSingleton<IViewABC, ViewABC>();
            containerRegistry.RegisterSingleton<IViewABC, ViewA>();
            containerRegistry.RegisterSingleton<IViewABC, ViewB>();
            containerRegistry.RegisterSingleton<IViewABC, ViewC>();

            containerRegistry.RegisterSingleton<IDoorDataService, DoorDataService>();
            containerRegistry.RegisterSingleton<IDoorLookupDataService, DoorLookupDataService>();

            containerRegistry.RegisterDialog<OkCancelDialog, OkCancelDialogViewModel>("OkCancelDialog");

            // Figure out how to use one Type

            //containerRegistry.Register<IFriendLookupDataService, LookupDataService>();
            //containerRegistry.Register<IProgrammingLanguageLookupDataService, LookupDataService>();
            //containerRegistry.Register<IMeetingLookupDataService, LookupDataService>();

            Log.MODULE("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // 03

        public void OnInitialized(IContainerProvider containerProvider)
        {
            Int64 startTicks = Log.MODULE("Enter", Common.LOG_CATEGORY);



            Log.MODULE("Exit", Common.LOG_CATEGORY, startTicks);
        }
    }
}
