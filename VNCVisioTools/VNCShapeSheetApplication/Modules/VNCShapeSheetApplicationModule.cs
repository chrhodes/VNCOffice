using System.Reflection;

using Prism.Ioc;
using Prism.Modularity;
using Prism.Regions;

using Unity;

using VNC;

//using VNCShapeSheetApplication.Presentation.ViewModels;
//using VNCShapeSheetApplication.Presentation.Views;

namespace VNCShapeSheetApplication.Modules
{
    public class VNCShapeSheetApplicationModule : IModule
    {
        // 01
        public void RegisterTypes(IContainerRegistry containerRegistry)
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");



            //Log.MODULE_INITIALIZE("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // 02
        public void OnInitialized(IContainerProvider containerProvider)
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            var regionManager = containerProvider.Resolve<IRegionManager>();



            //Log.MODULE_INITIALIZE("Exit", Common.LOG_CATEGORY, startTicks);
        }
    }
}
