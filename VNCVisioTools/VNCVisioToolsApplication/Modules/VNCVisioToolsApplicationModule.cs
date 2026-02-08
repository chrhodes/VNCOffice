using System.Reflection;

using Prism.Ioc;
using Prism.Modularity;
using Prism.Regions;

//using VNCVisioToolsApplication.Presentation.ViewModels;

using Unity;

using VNC;

using VNCVisioToolsApplication.Core;
using VNCVisioToolsApplication.Presentation.ViewModels;
using VNCVisioToolsApplication.Presentation.Views;

namespace VNCVisioToolsApplication.Modules
{
    public class VNCVisioToolsApplicationModule : IModule
    {
        // 01
        public void RegisterTypes(IContainerRegistry containerRegistry)
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            containerRegistry.Register<IViewCViewModel, ViewCViewModel>();

            //Log.MODULE_INITIALIZE("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // 02
        public void OnInitialized(IContainerProvider containerProvider)
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            var regionManager = containerProvider.Resolve<IRegionManager>();

            // NOTE(crhodes)
            // This is for PrismRegionTest.xaml

            regionManager.RegisterViewWithRegion(RegionNames.EditTextRegion, typeof(EditText));

            regionManager.RegisterViewWithRegion(RegionNames.EditControlPointsRegion, typeof(EditControlPoints));

            regionManager.RegisterViewWithRegion(RegionNames.EditParagraphRegion, typeof(EditParagraph));

            //Log.MODULE_INITIALIZE("Exit", Common.LOG_CATEGORY, startTicks);
        }
    }
}
