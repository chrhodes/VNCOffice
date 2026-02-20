using System;
using System.Threading;

using Microsoft.Office.Tools.Ribbon;

using VNC;

namespace VNCVisioTools
{
    public partial class Ribbon
    {
        // NOTE(crhodes)
        // This was moved out of designer so we can log

        public Ribbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
#if DEBUG
            Common.InitializeLogging(debugConfig: true);
#else
            Common.InitializeLogging();
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
