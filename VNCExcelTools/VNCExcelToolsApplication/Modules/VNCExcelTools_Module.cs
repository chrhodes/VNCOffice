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
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            containerRegistry.Register<IFolderMapViewModel, FolderMapViewModel>();
            containerRegistry.RegisterSingleton<IFolderMap, FolderMap>();

            containerRegistry.Register<ICreateFolderMapViewModel, CreateFolderMapViewModel>();
            containerRegistry.RegisterSingleton<ICreateFolderMap, CreateFolderMap>();
            //Log.MODULE_INITIALIZE("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // 02
        public void OnInitialized(IContainerProvider containerProvider)
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            var regionManager = containerProvider.Resolve<IRegionManager>();

            // NOTE(crhodes)
            // This is for PrismRegionTest.xaml

            // regionManager.RegisterViewWithRegion(RegionNames.EditTextRegion, typeof(EditText));

            // regionManager.RegisterViewWithRegion(RegionNames.EditControlPointsRegion, typeof(EditControlPoints));

            // regionManager.RegisterViewWithRegion(RegionNames.EditParagraphRegion, typeof(EditParagraph));

            //Log.MODULE_INITIALIZE("Exit", Common.LOG_CATEGORY, startTicks);
        }
    }
}
