using System;
using System.Windows;

using ModuleA;

using Prism.Ioc;
using Prism.Modularity;

namespace SupportTools_Visio.Application
{
    public class PrismApp : Prism.Unity.PrismApplication
    {
        protected override Window CreateShell()
        {
            //if (SupportTools_Visio.Ribbon.windowHostLocal is null)
            //{
            //    Ribbon.windowHostLocal = new Presentation.Views.WindowHost();
            //}

            return Ribbon.windowHostLocal;
            ////return null;
            ////throw new NotImplementedException();
        }

        protected override void RegisterTypes(IContainerRegistry containerRegistry)
        {
            //throw new NotImplementedException();
        }

        protected override void ConfigureModuleCatalog(IModuleCatalog moduleCatalog)
        {
            Type moduleAType = typeof(ModuleAModule);

            moduleCatalog.AddModule(new ModuleInfo()
            {
                ModuleName = moduleAType.Name,
                ModuleType = moduleAType.AssemblyQualifiedName,
                InitializationMode = InitializationMode.WhenAvailable
                // InitializationMode = InitializationMode.OnDemand
            });

            base.ConfigureModuleCatalog(moduleCatalog);
        }
    }
}
