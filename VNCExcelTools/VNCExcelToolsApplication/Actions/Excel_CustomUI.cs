using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using VNC.Core.Presentation;
using VNC.WPF.Presentation.Dx.Views;

using VNCExcelToolsApplication.Presentation.ViewModels;
using VNCExcelToolsApplication.Presentation.Views;

namespace VNCExcelToolsApplication.Actions
{
    public class Excel_CustomUI
    {
        #region FolderMap

        public static DxThemedWindowHost folderMapHost = null;

        public static void FolderMap()
        {
            if (folderMapHost is null) folderMapHost = new DxThemedWindowHost(Common.EventAggregator);

            folderMapHost.DisplayUserControlInHost(
                "Folder Map",
                300, 250,
                ShowWindowMode.Modeless_Show,
                //(CommandCockpitViewModel)Common.ApplicationBootstrapper.Container.Resolve(typeof(CommandCockpitViewModel))
                (FolderMap)Common.ApplicationBootstrapper.Container.Resolve(typeof(FolderMap))
            //new FolderMap()
            );
        }

        #endregion

        #region TestExcelLogging

        public static DxThemedWindowHost TestExcelLoggingHost = null;

        public static void TestExcelLogging()
        {
            if (TestExcelLoggingHost is null) TestExcelLoggingHost = new DxThemedWindowHost(Common.EventAggregator);

            TestExcelLoggingHost.DisplayUserControlInHost(
                "Folder Map",
                Common.DEFAULT_WINDOW_WIDTH, Common.DEFAULT_WINDOW_HEIGHT,
                ShowWindowMode.Modeless_Show,
                //(CommandCockpitViewModel)Common.ApplicationBootstrapper.Container.Resolve(typeof(CommandCockpitViewModel))
                (TestExcelLogging)Common.ApplicationBootstrapper.Container.Resolve(typeof(TestExcelLogging))
            //new TestExcelLogging()
            );
        }

        #endregion

        #region LoggingConfiguration

        public static DxThemedWindowHost LoggingConfigurationHost = null;

        public static void LoggingConfiguration()
        {
            if (LoggingConfigurationHost is null) LoggingConfigurationHost = new DxThemedWindowHost(Common.EventAggregator);

            LoggingConfigurationHost.DisplayUserControlInHost(
                "Folder Map",
                Common.DEFAULT_WINDOW_WIDTH, Common.DEFAULT_WINDOW_HEIGHT,
                ShowWindowMode.Modeless_Show,
                //(CommandCockpitViewModel)Common.ApplicationBootstrapper.Container.Resolve(typeof(CommandCockpitViewModel))
                //(LoggingConfiguration)Common.ApplicationBootstrapper.Container.Resolve(typeof(LoggingConfiguration))
                new VNCLoggingConfigMain()
            );
        }

        #endregion
    }
}
