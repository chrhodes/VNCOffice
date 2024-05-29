using Prism.Ioc;
using Prism.Modularity;
using Prism.Regions;

using Unity;

using VNC;

using SupportTools_Visio.Core;

namespace ModuleA
{
    public class ModuleAModule : IModule
    {
        #region Prism 6.3

        //IUnityContainer _container;
        //IRegionManager _regionManager;

        // Need a container so we can register our views
        // and a RegionManager so we can compose views and perform View discovery.
        // Prism will pass in when this module is created.

        //public ModuleAModule(IUnityContainer container, IRegionManager regionManager)
        //{
        //    _container = container;
        //    _regionManager = regionManager;
        //}

        //public void Initialize()
        //{
        //  // 1. Register Views
        //  // Nothing happens other thn Unity knows about View

        //  //_container.RegisterType<ToolBarA>();  // If this is commented out, still works!
        //  //_container.RegisterType<ContentA>();  // If this is commented out, still works!

        //  // 2. Tell the Region where the View goes.
        //  // NB. This also registers the type

        //    // Magic strings

        //    //_regionManager.RegisterViewWithRegion("ToolBarRegion", typeof(ToolBarView));
        //    //_regionManager.RegisterViewWithRegion("ContentRegion", typeof(ContentView));

        //    // No more Magic Strings

        //    //_regionManager.RegisterViewWithRegion(RegionNames.ToolBarRegion, typeof(ToolBarView));
        //    //_regionManager.RegisterViewWithRegion(RegionNames.ContentRegion, typeof(ContentView));

        //    // Multiple ToolBar Regions

        //    IRegion region = _regionManager.Regions[RegionNames.ToolBarRegionA];

        //    region.Add(_container.Resolve<ToolBarView>());
        //    region.Add(_container.Resolve<ToolBarView>());
        //    region.Add(_container.Resolve<ToolBarView>());
        //    region.Add(_container.Resolve<ToolBarView>());
        //    region.Add(_container.Resolve<ToolBarView>());

        //    _regionManager.RegisterViewWithRegion(RegionNames.ContentRegionA, typeof(ContentView));
        //}
        #endregion

        // 01
        public void RegisterTypes(IContainerRegistry containerRegistry)
        {
            long startTicks = Log.MODULE_INITIALIZE("Enter", Common.LOG_CATEGORY, 0);


            Log.MODULE_INITIALIZE("Exit", Common.LOG_CATEGORY, 0, startTicks);
        }

        // 02
        public void OnInitialized(IContainerProvider containerProvider)
        {
            long startTicks = Log.MODULE_INITIALIZE("Enter", Common.LOG_CATEGORY, 0);

            var regionManager = containerProvider.Resolve<IRegionManager>();

            // Multiple ToolBar Regions

            //IRegion region = regionManager.Regions[RegionNames.ToolBarRegionA];

            //region.Add(containerProvider.Resolve<ToolBarView>());
            //region.Add(containerProvider.Resolve<ToolBarView>());
            //region.Add(containerProvider.Resolve<ToolBarView>());
            //region.Add(containerProvider.Resolve<ToolBarView>());
            //region.Add(containerProvider.Resolve<ToolBarView>());

            regionManager.RegisterViewWithRegion(RegionNames.ToolBarRegionA, typeof(ToolBarView));

            // NOTE(crhodes)
            // Can't get this to work.  Hum.  May have to eschew multiple toolbars :)

            //IRegion region = regionManager.Regions[RegionNames.ToolBarRegionA];

            //region.Add(containerProvider.Resolve<ToolBarView>());
            //region.Add(containerProvider.Resolve<ToolBarView>());

            regionManager.RegisterViewWithRegion(RegionNames.ContentRegionA, typeof(ContentView));

            Log.MODULE_INITIALIZE("Exit", Common.LOG_CATEGORY, 0, startTicks);
        }
    }
}
