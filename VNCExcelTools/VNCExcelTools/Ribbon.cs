using System;
using System.Configuration;
using System.IO;
using System.Threading;

using Microsoft.Office.Tools.Ribbon;

using VNC;

namespace VNCExcelTools
{
    public partial class Ribbon
    {
        // NOTE(crhodes)
        // This was moved out of designer so we can log

        public Ribbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();

            Common.WriteToDebugWindow("Ribbon()", true);

            var workingDirectory = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            var currentDirectory = Directory.GetCurrentDirectory();
            var appDomainDirectory = System.AppDomain.CurrentDomain.BaseDirectory;

            Common.WriteToDebugWindow($" - Working   Directory: {workingDirectory}", true);
            Common.WriteToDebugWindow($" - Current   Directory: {currentDirectory}", true);
            Common.WriteToDebugWindow($" - AppDomain Directory: {appDomainDirectory}", true);

            string loggingConfigurationPath = GetAppSetting("LoggingConfigurationPath");

            loggingConfigurationPath = string.IsNullOrEmpty(loggingConfigurationPath) ? currentDirectory : loggingConfigurationPath;

#if DEBUG
            Common.InitializeLogging(new VNC.Core.Logging.LogLevelConfiguration(
                filePath: loggingConfigurationPath, fileName: "vncloggingconfig-debug.json", isDebugConfig: true));

            Common.InitializeCoreLogging(new VNC.Core.Logging.LogLevelConfiguration(
                filePath: loggingConfigurationPath, fileName: "vnccoreloggingconfig-debug.json", isDebugConfig: true));
            //Common.InitializeLogging(debugConfig: true);
#else
            Common.InitializeLogging(new VNC.Core.Logging.LogLevelConfiguration(
                filePath: loggingConfigurationPath, fileName: "ConfigurationSettings.json",  isDebugConfig: false));

            Common.InitializeCoreLogging(new VNC.Core.Logging.LogLevelConfiguration(
                filePath: loggingConfigurationPath, fileName: "vnccoreloggingconfig.json", isDebugConfig: false));
            //Common.InitializeLogging();
#endif

            Int64 startTicks = 0;

            if (Common.VNCLogLevel.ApplicationStart) startTicks = Log.APPLICATION_START("Initialize SignalR", Common.LOG_CATEGORY);

            // NOTE(crhodes)
            // If don't delay a bit here, the SignalR logging infrastructure does not initialize quickly enough
            // and the first few log messages are missed.
            // NB.  All are properly recored in the log file.

            Thread.Sleep(250);

            if (Common.VNCLogLevel.ApplicationStart) startTicks = Log.APPLICATION_START("Enter/Exit", Common.LOG_CATEGORY);
        }

        private string GetAppSetting(string key)
        {
            string rawValue = ConfigurationManager.AppSettings[key];
            string settingValue = string.IsNullOrEmpty(rawValue) ? string.Empty : Environment.ExpandEnvironmentVariables(rawValue);

            if (string.IsNullOrEmpty(settingValue))
            {
                Common.WriteToDebugWindow($"Did not find setting for key: {key}  Check app.config", true);
                Log.ERROR($"Did not find setting for key: {key}  Check app.config", Common.LOG_ERROR);
            }

            return settingValue;
        }

        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            Int64 startTicks = 0;
            Common.WriteToDebugWindow("Enter/Exit", true);
            if (Common.VNCLogLevel.ApplicationStart) startTicks = Log.APPLICATION_START("Enter/Exit", Common.LOG_CATEGORY);
        }



    }
}
