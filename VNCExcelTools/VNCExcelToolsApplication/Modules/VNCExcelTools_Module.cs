using System;
using System.Reflection;

using Prism.Ioc;
using Prism.Modularity;
using Prism.Navigation.Regions;

//using VNCExcelToolsApplication.Presentation.ViewModels;

using Unity;

using VNC;

using VNCExcelToolsApplication.Core;
using VNCExcelToolsApplication.Presentation.ViewModels;
using VNCExcelToolsApplication.Presentation.Views;

namespace VNCExcelToolsApplication.Modules
{
    public class VNCExcelToolsApplicationModule : IModule
    {
        // 01
        public void RegisterTypes(IContainerRegistry containerRegistry)
        {
            Int64 startTicks = 0;
            startTicks = Common.WriteToDebugWindow("Enter", true);
            if (Common.VNCLogLevel.ModuleInitialize) startTicks = Log.MODULE_INITIALIZE("Enter", Common.LOG_CATEGORY);

            containerRegistry.Register<IFolderMapViewModel, FolderMapViewModel>();
            containerRegistry.RegisterSingleton<IFolderMap, FolderMap>();

            containerRegistry.Register<ICreateFolderMapViewModel, CreateFolderMapViewModel>();
            containerRegistry.RegisterSingleton<ICreateFolderMap, CreateFolderMap>();

            //containerRegistry.Register<ILoggingConfigurationViewModel, LoggingConfigurationViewModel>();
            //containerRegistry.RegisterSingleton<ILoggingConfiguration, LoggingConfiguration>();

            containerRegistry.Register<ITestExcelLoggingViewModel, TestExcelLoggingViewModel>();
            containerRegistry.RegisterSingleton<ITestExcelLogging, TestExcelLogging>();

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

            regionManager.RegisterViewWithRegion(RegionNames.VNCLoggingConfigRegion, typeof(VNC.WPF.Presentation.Dx.Views.VNCLoggingConfig));
            regionManager.RegisterViewWithRegion(RegionNames.VNCCoreLoggingConfigRegion, typeof(VNC.WPF.Presentation.Dx.Views.VNCCoreLoggingConfig));

            // NOTE(crhodes)
            // This is for PrismRegionTest.xaml

            // regionManager.RegisterViewWithRegion(RegionNames.EditTextRegion, typeof(EditText));

            // regionManager.RegisterViewWithRegion(RegionNames.EditControlPointsRegion, typeof(EditControlPoints));

            // regionManager.RegisterViewWithRegion(RegionNames.EditParagraphRegion, typeof(EditParagraph));

            Common.WriteToDebugWindow("Exit", startTicks, true);
            if (Common.VNCLogLevel.ModuleInitialize) Log.MODULE_INITIALIZE("Exit", Common.LOG_CATEGORY, startTicks);
        }
    }
}
