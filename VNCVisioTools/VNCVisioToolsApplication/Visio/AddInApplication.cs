﻿using System;
using System.Reflection;

using Prism.Events;

using VNC;
using VNC.Core.Events;

using VNCVisioToolsApplication.Actions;

namespace VNCVisioToolsApplication.Visio
{
    public class AddInApplication
    {
        private static System.Windows.Application _XamlApp;

        private static Prism.Unity.PrismApplication _prismApplication;

        public static void InitializeApplication()
        {
            Int64 startTicks;
            startTicks = Log.APPLICATION_INITIALIZE("Enter", Common.LOG_CATEGORY);

            startTicks = Common.WriteToDebugWindow("InitializeWPFApplication()", true);

            //Common.CurrentUser = new WindowsPrincipal(WindowsIdentity.GetCurrent());

            //// NOTE(crhodes)
            //// We need to update VNC.Core as VNCCoreLogging and VNCLogging are null
            //// We started initializing them in 3.0+

            //VNC.Core.Common.VNCCoreLogging = new VNC.Core.VNCLoggingConfig();
            //VNC.Core.Common.VNCLogging = new VNC.Core.VNCLoggingConfig();

            GetAndSetInformation();

            CreateXamlApplication();

            InitializePrism();

            Log.APPLICATION_INITIALIZE("Exit", Common.LOG_CATEGORY, startTicks);
            Common.WriteToDebugWindow("InitializeWPFApplication()-Exit", startTicks, true);
        }

        /// <summary>
        /// LoadXamlApplicationResources
        ///
        /// Creates Xaml Resources collection in System.Windows.Application
        /// for use in Hosted applications without App.Xaml
        /// </summary>

        private static void GetAndSetInformation()
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.ApplicationInitialize) startTicks = Log.APPLICATION_INITIALIZE("Enter", Common.LOG_CATEGORY);

            // Get Information about VNC.Core

            Common.SetVersionInfoVNCCore();

            var appFileVersionInfo = System.Diagnostics.FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location);

            // Get Information about ourselves

            Common.SetVersionInfoApplication(Assembly.GetExecutingAssembly(), appFileVersionInfo);

            // Add Information about the other assemblies in our application

            // TODO(crhodes)
            // Gather VNC.Core.Information InformationXXX
            // for other Assemblies that should provide Info 
            // listed in VNCVisioToolsApplication.Common
            //
            // Use GAI
            // 
            // Extend Views\AppVersionInfo.xaml as needed
            // Update Views\AppVersionInfo.xaml.cs InitializeViewModel()

            var VNCVisioVSTOAddInAssembly = Assembly.GetAssembly(typeof(VNC.Visio.VSTOAddIn.Common));

            if (VNCVisioVSTOAddInAssembly is not null)
            {
                var VNCVisioVSTOAddInAssemblyFileVersionInfo = System.Diagnostics.FileVersionInfo
                    .GetVersionInfo(VNCVisioVSTOAddInAssembly.Location);

                Common.InformationVNCVisioVSTOAddIn = Common.GetInformation(
                    VNCVisioVSTOAddInAssembly,
                    VNCVisioVSTOAddInAssemblyFileVersionInfo);
            }

            var VNCVSTOAddInAssembly = Assembly.GetAssembly(typeof(VNC.VSTOAddIn.Common));

            if (VNCVSTOAddInAssembly is not null)
            {
                var VNCVSTOAddInAssemblyFileVersionInfo = System.Diagnostics.FileVersionInfo
                    .GetVersionInfo(VNCVSTOAddInAssembly.Location);

                Common.InformationVNCVSTOAddIn = Common.GetInformation(
                    VNCVSTOAddInAssembly,
                    VNCVSTOAddInAssemblyFileVersionInfo);
            }

            var VNCWpfPresentationAssembly = Assembly.GetAssembly(typeof(VNC.WPF.Presentation.Common));

            if (VNCWpfPresentationAssembly is not null)
            {
                var VNCWpfPresentationAssemblyFileVersionInfo = System.Diagnostics.FileVersionInfo
                    .GetVersionInfo(VNCWpfPresentationAssembly.Location);

                Common.InformationVNCWpfPresentation = Common.GetInformation(
                    VNCWpfPresentationAssembly,
                    VNCWpfPresentationAssemblyFileVersionInfo);
            }

            var VNCWpfPresentationDxAssembly = Assembly.GetAssembly(typeof(VNC.WPF.Presentation.Dx.Common));

            if (VNCWpfPresentationDxAssembly is not null)
            {
                var VNCWpfPresentationDxAssemblyFileVersionInfo = System.Diagnostics.FileVersionInfo
                        .GetVersionInfo(VNCWpfPresentationDxAssembly.Location);

                Common.InformationVNCWpfPresentationDx = Common.GetInformation(
                    VNCWpfPresentationDxAssembly,
                    VNCWpfPresentationDxAssemblyFileVersionInfo);
            }
            

            if (Common.VNCLogging.ApplicationInitialize) Log.APPLICATION_INITIALIZE("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private static void CreateXamlApplication()
        {
            //Int64 startTicks = Log.APPLICATION_INITIALIZE("Enter", Common.LOG_CATEGORY);

            Int64 startTicks = Common.WriteToDebugWindow("CreateXamlApplication()", true);

            try
            {
                // TODO(crhodes)

                // Can we just create a PrismApplication?
                // Create a WPF Application
                _XamlApp = new System.Windows.Application();

                //_prismApplication = new Application.PrismApp();

                //var defaultThemes = DevExpress.Xpf.Core.Theme.Themes;
                //ApplicationThemeHelper.ApplicationThemeName = "MetropolisDark";

                // Load the resources

                // This works

                //var resources = System.Windows.Application.LoadComponent(
                //    new Uri("SupportTools_Excel;component/Resources/Xaml/Brushes.xaml", UriKind.Relative)) as System.Windows.ResourceDictionary;

                // Now lets try with

                var resources = System.Windows.Application.LoadComponent(
                    new Uri("VNCVisioToolsApplication;component/Resources/Xaml/Application.xaml", UriKind.Relative)) as System.Windows.ResourceDictionary;

                //var resources = System.Windows.Application.LoadComponent(
                //    new Uri("pack:/SupportTools_Excel;:,,/Resources/Xaml/Application.xaml")) as System.Windows.ResourceDictionary;

                // Merge it on application level

                _XamlApp.Resources.MergedDictionaries.Add(resources);

                //_prismApplication.Resources.MergedDictionaries.Add(resources);
            }
            catch (Exception ex)
            {
                Common.DeveloperMode = true;
                Common.WriteToDebugWindow(ex.ToString(), true);
                Common.DeveloperMode = false;
            }

            //Log.APPLICATION_INITIALIZE("Exit", Common.LOG_CATEGORY, startTicks);
            Common.WriteToDebugWindow("CreateXamlApplication()-Exit", startTicks, true);
        }

        private static void InitializePrism()
        {
            Int64 startTicks = Common.WriteToDebugWindow("InitializePrism()", true);

            Common.ApplicationBootstrapper = new Bootstrapper();
            Common.ApplicationBootstrapper.Run();

            Common.EventAggregator = (IEventAggregator)Common.Container.Resolve(typeof(IEventAggregator));
            Visio_Application.statusMessageEvent = Common.EventAggregator.GetEvent<StatusMessageEvent>();
            Visio_Application.developerModeEvent = Common.EventAggregator.GetEvent<DeveloperModeEvent>();

            Common.WriteToDebugWindow("InitializePrism()-Exit", startTicks, true);
        }

        private void UnLoadXamlApplicationResources()
        {
            //Int64 startTicks = Log.APPLICATION_END("Enter", Common.LOG_CATEGORY);
            Int64 startTicks = Common.WriteToDebugWindow("UnLoadXamlApplicationResources()", true);

            try
            {
                if (null != _XamlApp)
                {
                    _XamlApp.Shutdown();
                    _XamlApp = null;
                }
                if (null != _prismApplication)
                {
                    _prismApplication.Shutdown();
                    _prismApplication = null;
                }
            }
            catch (Exception ex)
            {
                Common.DeveloperMode = true;
                Common.WriteToDebugWindow(ex.ToString(), true);
                Common.DeveloperMode = false;
            }

            //Log.APPLICATION_END("Exit", Common.LOG_CATEGORY, startTicks);
            Common.WriteToDebugWindow("Exit", startTicks, true);
        }
    }
}
