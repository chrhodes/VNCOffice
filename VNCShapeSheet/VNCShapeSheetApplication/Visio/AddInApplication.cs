using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

using VNC;

namespace VNCShapeSheetApplication.Visio
{
    public class AddInApplication
    {
        private static System.Windows.Application _XamlApp;

        private static Prism.Unity.PrismApplication _prismApplication;

        public static void InitializeApplication()
        {
            //Int64 startTicks = Log.APPLICATION_INITIALIZE("Enter", Common.LOG_CATEGORY);

            Int64 startTicks = Common.WriteToDebugWindow("InitializeWPFApplication()", true);

            //Common.CurrentUser = new WindowsPrincipal(WindowsIdentity.GetCurrent());

            // TODO(crhodes)
            // We need to update VNC.Core as VNCCoreLogging and VNCLogging are null
            // We started initializing them in 3.0+

            VNC.Core.Common.VNCCoreLogging = new VNC.Core.VNCLoggingConfig();
            VNC.Core.Common.VNCLogging = new VNC.Core.VNCLoggingConfig();

            VNC.Core.Common.VNCLogging.ApplicationStart = true;

            Log.APPLICATION_START("VNCShapeSheetApplication", Common.LOG_CATEGORY);

            CreateXamlApplication();

            InitializePrism();

            Common.WriteToDebugWindow("InitializeWPFApplication()-Exit", startTicks, true);
        }

        /// <summary>
        /// LoadXamlApplicationResources
        ///
        /// Creates Xaml Resources collection in System.Windows.Application
        /// for use in Hosted applications without App.Xaml
        /// </summary>

        private static void CreateXamlApplication()
        {
            //Int64 startTicks = Log.APPLICATION_INITIALIZE("Enter", Common.LOG_CATEGORY);

            Int64 startTicks = Common.WriteToDebugWindow("CreateXamlApplication()", true);

            try
            {
                // Create a WPF Application
                _XamlApp = new System.Windows.Application();


                var resources = System.Windows.Application.LoadComponent(
                    new Uri("VNCShapeSheetApplication;component/Resources/Xaml/Application.xaml", UriKind.Relative)) as System.Windows.ResourceDictionary;


                // Merge it on application level

                _XamlApp.Resources.MergedDictionaries.Add(resources);

            }
            catch (Exception ex)
            {
                Common.DeveloperMode = true;
                Common.WriteToDebugWindow(ex.ToString(), true);
                Common.DeveloperMode = false;
            }

            Common.WriteToDebugWindow("CreateXamlApplication()-Exit", startTicks, true);
        }

        private static void InitializePrism()
        {
            Int64 startTicks = Common.WriteToDebugWindow("InitializePrism()", true);

            Common.ApplicationBootstrapper = new Bootstrapper();
            Common.ApplicationBootstrapper.Run();

            Common.WriteToDebugWindow("InitializePrism()-Exit", startTicks, true);
        }

        private void UnLoadXamlApplicationResources()
        {
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

            Common.WriteToDebugWindow("Exit", startTicks, true);
        }
    }
}
