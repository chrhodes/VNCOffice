using System;
using System.Reflection;

using Prism.Events;

using VNC;
using VNC.Core.Events;

using VNCExcelToolsApplication.Actions;

namespace VNCExcelToolsApplication.Excel
{
    public class AddInApplication
    {
        private static System.Windows.Application _XamlApp;

        private static Prism.Unity.PrismApplication _prismApplication;

        public static void InitializeApplication()
        {
            Int64 startTicks = 0;
            Common.WriteToDebugWindow("Enter", true);
            if (Common.VNCLogging.ApplicationInitialize) startTicks = Log.APPLICATION_INITIALIZE("Enter", Common.LOG_CATEGORY);

            //Common.CurrentUser = new WindowsPrincipal(WindowsIdentity.GetCurrent());

            //// NOTE(crhodes)
            //// We need to update VNC.Core as VNCCoreLogging and VNCLogging are null
            //// We started initializing them in 3.0+

            //VNC.Core.Common.VNCCoreLogging = new VNC.Core.VNCLoggingConfig();
            //VNC.Core.Common.VNCLogging = new VNC.Core.VNCLoggingConfig();

            GetAndSetInformation();

            CreateXamlApplication();

            InitializePrism();

            Common.WriteToDebugWindow("Exit", startTicks, true);
            if (Common.VNCLogging.ApplicationInitialize) Log.APPLICATION_INITIALIZE("Exit", Common.LOG_CATEGORY, startTicks);
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
            Common.WriteToDebugWindow("Enter", true);
            if (Common.VNCLogging.ApplicationInitialize) startTicks = Log.APPLICATION_INITIALIZE("Enter", Common.LOG_CATEGORY);

            // Get Information about VNC.Core

            Common.SetVersionInfoVNCCore();

            //var appFileVersionInfo = System.Diagnostics.FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location);

            //Common.SetVersionInfoApplication(Assembly.GetExecutingAssembly(), appFileVersionInfo);

            // Get Information about ourselves

            var VNCExcelToolsApplication_Assembly = Assembly.GetExecutingAssembly();

            if (VNCExcelToolsApplication_Assembly is not null)
            {
                var VNCExcelToolsApplication_AssemblyFileVersionInfo = System.Diagnostics.FileVersionInfo
                        .GetVersionInfo(VNCExcelToolsApplication_Assembly.Location);

                Common.InformationVNCExcelToolsApplication = Common.GetInformation(
                    VNCExcelToolsApplication_Assembly,
                    VNCExcelToolsApplication_AssemblyFileVersionInfo);
            }

            var VNCExcelToolsApplicationCore_Assembly = Assembly.GetAssembly(typeof(VNCExcelToolsApplication.Core.RegionNames));

            if (VNCExcelToolsApplicationCore_Assembly is not null)
            {
                var VNCExcelToolsApplicationCore_AssemblyFileVersionInfo = System.Diagnostics.FileVersionInfo
                        .GetVersionInfo(VNCExcelToolsApplicationCore_Assembly.Location);

                Common.InformationVNCExcelToolsApplicationCore = Common.GetInformation(
                    VNCExcelToolsApplicationCore_Assembly,
                    VNCExcelToolsApplicationCore_AssemblyFileVersionInfo);
            }

            // Add Information about the other assemblies in our application

            // TODO(crhodes)
            // Gather VNC.Core.Information InformationXXX
            // for other Assemblies that should provide Info
            // listed in VNCExcelTools.Common
            //
            // Use GAI
            //
            // Extend Views\AppVersionInfo.xaml as needed
            // Update Views\AppVersionInfo.xaml.cs InitializeViewModel()

            var VNCVSTOAddInExcel_Assembly = Assembly.GetAssembly(typeof(VNC.VSTOAddIn.Excel.Common));

            if (VNCVSTOAddInExcel_Assembly is not null)
            {
                var VNCVSTOAddInExcel_AssemblyFileVersionInfo = System.Diagnostics.FileVersionInfo
                    .GetVersionInfo(VNCVSTOAddInExcel_Assembly.Location);

                Common.InformationVNCVSTOAddInExcel = Common.GetInformation(
                    VNCVSTOAddInExcel_Assembly,
                    VNCVSTOAddInExcel_AssemblyFileVersionInfo);
            }

            var VNCVSTOAddIn_Assembly = Assembly.GetAssembly(typeof(VNC.VSTOAddIn.Common));

            if (VNCVSTOAddIn_Assembly is not null)
            {
                var VNCVSTOAddIn_AssemblyFileVersionInfo = System.Diagnostics.FileVersionInfo
                    .GetVersionInfo(VNCVSTOAddIn_Assembly.Location);

                Common.InformationVNCVSTOAddIn = Common.GetInformation(
                    VNCVSTOAddIn_Assembly,
                    VNCVSTOAddIn_AssemblyFileVersionInfo);
            }

            var VNCWpfPresentation_Assembly = Assembly.GetAssembly(typeof(VNC.WPF.Presentation.Common));

            if (VNCWpfPresentation_Assembly is not null)
            {
                var VNCWpfPresentation_AssemblyFileVersionInfo = System.Diagnostics.FileVersionInfo
                    .GetVersionInfo(VNCWpfPresentation_Assembly.Location);

                Common.InformationVNCWpfPresentation = Common.GetInformation(
                    VNCWpfPresentation_Assembly,
                    VNCWpfPresentation_AssemblyFileVersionInfo);
            }

            var VNCWpfPresentationDx_Assembly = Assembly.GetAssembly(typeof(VNC.WPF.Presentation.Dx.Common));

            if (VNCWpfPresentationDx_Assembly is not null)
            {
                var VNCWpfPresentationDx_AssemblyFileVersionInfo = System.Diagnostics.FileVersionInfo
                        .GetVersionInfo(VNCWpfPresentationDx_Assembly.Location);

                Common.InformationVNCWpfPresentationDx = Common.GetInformation(
                    VNCWpfPresentationDx_Assembly,
                    VNCWpfPresentationDx_AssemblyFileVersionInfo);
            }

            Common.WriteToDebugWindow("Exit", startTicks, true);
            if (Common.VNCLogging.ApplicationInitialize) Log.APPLICATION_INITIALIZE("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private static void CreateXamlApplication()
        {
            Int64 startTicks = 0;
            startTicks = Common.WriteToDebugWindow("Enter", true);
            if (Common.VNCLogging.ApplicationInitializeLow) startTicks = Log.APPLICATION_INITIALIZE_LOW("Enter", Common.LOG_CATEGORY);

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
                    new Uri("VNCExcelToolsApplication;component/Resources/Xaml/Application.xaml", UriKind.Relative)) as System.Windows.ResourceDictionary;

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

            Common.WriteToDebugWindow("Exit", startTicks, true);
            if (Common.VNCLogging.ApplicationInitializeLow) Log.APPLICATION_INITIALIZE_LOW("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private static void InitializePrism()
        {
            Int64 startTicks = 0;
            startTicks = Common.WriteToDebugWindow("Enter", true);
            if (Common.VNCLogging.ApplicationInitializeLow) startTicks = Log.APPLICATION_INITIALIZE_LOW("Enter", Common.LOG_CATEGORY);

            Common.ApplicationBootstrapper = new Bootstrapper();
            Common.ApplicationBootstrapper.Run();

            Common.EventAggregator = (IEventAggregator)Common.Container.Resolve(typeof(IEventAggregator));
            Excel_Application.statusMessageEvent = Common.EventAggregator.GetEvent<StatusMessageEvent>();
            Excel_Application.developerModeEvent = Common.EventAggregator.GetEvent<DeveloperModeEvent>();

            Common.WriteToDebugWindow("Exit", startTicks, true);
            if (Common.VNCLogging.ApplicationInitializeLow) Log.APPLICATION_INITIALIZE_LOW("Exit", Common.LOG_CATEGORY, startTicks);
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
