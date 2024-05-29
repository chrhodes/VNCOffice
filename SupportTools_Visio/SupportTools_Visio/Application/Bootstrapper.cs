using System;
using System.Windows;
using System.Windows.Controls;

//using ModuleA;
//using Explore;

using Prism;
using Prism.Ioc;
using Prism.Modularity;
using Prism.Regions;
using Prism.Unity;

using SupportTools_Visio.Modules;
using SupportTools_Visio.Presentation.ViewModels;
using SupportTools_Visio.Presentation.Views;

using VNC;
using VNC.Core.Mvvm.Prism;
using VNC.Core.Services;

//using VNC.Core.Mvvm;
//using VNC.Core.Mvvm.Prism;

namespace SupportTools_Visio.Application
{
    public class Bootstrapper : PrismBootstrapperBase
    {
        // Step 1a - Create the catalog of Modules

        protected override IModuleCatalog CreateModuleCatalog()
        {
            Int64 startTicks = Log.APPLICATION_INITIALIZE("Enter", Common.LOG_CATEGORY);

            Log.APPLICATION_INITIALIZE("Exit", Common.LOG_CATEGORY, startTicks);

            return new ConfigurationModuleCatalog();
        }

        // Step 1b - Configure the catalog of modules
        // Modules are loaded at Startup and must be a project reference

        protected override void ConfigureModuleCatalog(IModuleCatalog moduleCatalog)
        {
            Int64 startTicks = Log.APPLICATION_INITIALIZE("Enter", Common.LOG_CATEGORY);

            moduleCatalog.AddModule(typeof(SupportTools_VisioModule));

            moduleCatalog.AddModule(typeof(ModuleA.ModuleAModule));

            moduleCatalog.AddModule(typeof(EditTextModule));

            moduleCatalog.AddModule(typeof(Explore.ExploreModule));
            moduleCatalog.AddModule(typeof(Explore.CarModule));

            base.ConfigureModuleCatalog(moduleCatalog);

            Log.APPLICATION_INITIALIZE("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // Step 2 - Configure the container

        protected override void RegisterRequiredTypes(IContainerRegistry containerRegistry)
        {
            Int64 startTicks = Log.APPLICATION_INITIALIZE("Enter", Common.LOG_CATEGORY);

            // Registers all types that are required by Prism to function with the container.

            base.RegisterRequiredTypes(containerRegistry);

            Log.APPLICATION_INITIALIZE("Exit", Common.LOG_CATEGORY, startTicks);
        }

        protected override void RegisterTypes(IContainerRegistry containerRegistry)
        {
            Int64 startTicks = Log.APPLICATION_INITIALIZE("Enter", Common.LOG_CATEGORY);

            // Used to register types with the container that will be used by your application.

            containerRegistry.Register<IMessageDialogService, MessageDialogService>();

            containerRegistry.Register<IEditTextViewModel, EditTextViewModel>();
            containerRegistry.Register<EditText>();

            containerRegistry.Register<EditParagraphViewModel>();
            containerRegistry.Register<EditParagraph>();

            containerRegistry.Register<EditControlRowsViewModel>();
            containerRegistry.Register<EditControlRows>();

            Log.APPLICATION_INITIALIZE("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // Step 3 - Configure the RegionAdapters if any custom ones have been created

        // Step 4 - Create the Shell that will hold the modules in designated regions.

        protected override DependencyObject CreateShell()
        {
            Int64 startTicks = Log.APPLICATION_INITIALIZE("Enter", Common.LOG_CATEGORY);

            Log.APPLICATION_INITIALIZE("Exit (null)", Common.LOG_CATEGORY, startTicks);
            return null;
            //return Container.Resolve<Views.MainWindow>();
            //return Container.TryResolve<Views.MainWindow>();
        }

        // Step 5 - Show the MainWindow

        protected override void ConfigureRegionAdapterMappings(RegionAdapterMappings regionAdapterMappings)
        {
            Int64 startTicks = Log.APPLICATION_INITIALIZE("Enter", Common.LOG_CATEGORY);

            base.ConfigureRegionAdapterMappings(regionAdapterMappings);
            regionAdapterMappings.RegisterMapping(typeof(StackPanel), Container.Resolve<StackPanelRegionAdapter>());

            Log.APPLICATION_INITIALIZE("Exit", Common.LOG_CATEGORY, startTicks);
        }

        protected override void ConfigureDefaultRegionBehaviors(IRegionBehaviorFactory regionBehaviors)
        {
            Int64 startTicks = Log.APPLICATION_INITIALIZE("Enter", Common.LOG_CATEGORY);

            base.ConfigureDefaultRegionBehaviors(regionBehaviors);

            Log.APPLICATION_INITIALIZE("Exit", Common.LOG_CATEGORY, startTicks);
        }

        protected override IContainerExtension CreateContainerExtension()
        {
            Int64 startTicks = Log.APPLICATION_INITIALIZE("Enter", Common.LOG_CATEGORY);

            Log.APPLICATION_INITIALIZE("Exit", Common.LOG_CATEGORY, startTicks);

            return new UnityContainerExtension();
        }
    }
}
