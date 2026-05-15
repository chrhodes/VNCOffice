using Prism.Ioc;
using Prism.Modularity;
using Prism.Regions;

namespace ModuleOptions
{
    public class ModuleOptionsModule : IModule
    {
        public void RegisterTypes(IContainerRegistry containerRegistry)
        {
            containerRegistry.Register<IExcelOptions, ExcelOptions>();
            containerRegistry.Register<IExcelOptionsViewModel, ExcelOptionsViewModel>();
        }

        public void OnInitialized(IContainerProvider containerProvider)
        {
            var regionManager = containerProvider.Resolve<IRegionManager>();

            regionManager.RegisterViewWithRegion("PrismContent", typeof(ExcelOptions));
        }
    }
}
