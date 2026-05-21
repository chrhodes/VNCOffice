using System;
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
            var workingDirectory = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            var currentDirectory = Directory.GetCurrentDirectory();
            var appDomainDirectory = System.AppDomain.CurrentDomain.BaseDirectory;
            Common.WriteToDebugWindow($"Ribbon()", true);
            Common.WriteToDebugWindow($" - Working   Directory: {workingDirectory}", true);
            Common.WriteToDebugWindow($" - Current   Directory: {currentDirectory}", true);
            Common.WriteToDebugWindow($" - AppDomain Directory: {appDomainDirectory}", true);

            currentDirectory = @"C:\temp";
#if DEBUG
            Common.InitializeLogging(new VNC.Core.LoggingConfiguration(
                configFilePath: currentDirectory, configFile: "vncloggingconfig-debug.json",  isDebugConfig: true));

            Common.InitializeCoreLogging(new VNC.Core.LoggingConfiguration(
                configFilePath: currentDirectory, configFile: "vnccoreloggingconfig-debug.json", isDebugConfig: true));
            //Common.InitializeLogging(debugConfig: true);
#else
            Common.InitializeLogging(new VNC.Core.LoggingConfiguration(
                configFilePath: currentDirectory, configFile: "vncloggingconfig.json",  isDebugConfig: false));

            Common.InitializeCoreLogging(new VNC.Core.LoggingConfiguration(
                configFilePath: currentDirectory, configFile: "vnccoreloggingconfig.json", isDebugConfig: false));
            //Common.InitializeLogging();
#endif
            Int64 startTicks = 0;
            Common.WriteToDebugWindow("Ribbon()", true);
            if (Common.VNCLogging.ApplicationStart) startTicks = Log.APPLICATION_START("Initialize SignalR", Common.LOG_CATEGORY);

            // NOTE(crhodes)
            // If don't delay a bit here, the SignalR logging infrastructure does not initialize quickly enough
            // and the first few log messages are missed.
            // NB.  All are properly recored in the log file.

            Thread.Sleep(250);

            if (Common.VNCLogging.ApplicationStart) startTicks = Log.APPLICATION_START("Enter/Exit", Common.LOG_CATEGORY);
        }

        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            Int64 startTicks = 0;
            Common.WriteToDebugWindow("Ribbon_Load()", true);
            if (Common.VNCLogging.ApplicationStart) startTicks = Log.APPLICATION_START("Enter/Exit", Common.LOG_CATEGORY);
        }


    }
}
