using System;
using System.Reflection;

using Prism.Ioc;
using Prism.Modularity;
using Prism.Navigation.Regions;

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
            Int64 startTicks = 0;
            startTicks = Common.WriteToDebugWindow("Enter", true);
            if (Common.VNCLogLevel.ModuleInitialize) startTicks = Log.MODULE_INITIALIZE("Enter", Common.LOG_CATEGORY);

            containerRegistry.Register<IViewCViewModel, ViewCViewModel>();

            Common.WriteToDebugWindow("Exit", startTicks, true);
            if (Common.VNCLogLevel.ModuleInitialize) Log.MODULE_INITIALIZE("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // 02
        public void OnInitialized(IContainerProvider containerProvider)
        {
            Int64 startTicks = 0;
            startTicks = Common.WriteToDebugWindow("Enter", true);
            if (Common.VNCLogLevel.ModuleInitialize) startTicks = Log.MODULE_INITIALIZE("Enter", Common.LOG_CATEGORY);

            var regionManager = containerProvider.Resolve<IRegionManager>();

            // NOTE(crhodes)
            // This is for PrismRegionTest.xaml

            regionManager.RegisterViewWithRegion(RegionNames.EditTextRegion, typeof(EditText));

            regionManager.RegisterViewWithRegion(RegionNames.EditControlPointsRegion, typeof(EditControlPoints));

            regionManager.RegisterViewWithRegion(RegionNames.EditParagraphRegion, typeof(EditParagraph));

            Common.WriteToDebugWindow("Exit", startTicks, true);
            if (Common.VNCLogLevel.ModuleInitialize) Log.MODULE_INITIALIZE("Exit", Common.LOG_CATEGORY, startTicks);
        }
    }
}
