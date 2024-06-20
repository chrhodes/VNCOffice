using System;
using System.Threading;

using Microsoft.Office.Tools.Ribbon;

using VNC;

namespace SupportTools_Visio
{
    public partial class Ribbon
    {
        // NOTE(crhodes)
        // This was moved out of designer so we can call bootstrapper.

        //Prism.Ioc.IContainerProvider _container;
        public Ribbon()
           : base(Globals.Factory.GetRibbonFactory())
        {
            // HACK(crhodes)
            // If don't delay a bit here, the SignalR logging infrastructure
            // does not initialize quickly enough
            // and the first few log messages are missed.
            // NB.  All are properly recored in the log file.

            Log.APPLICATION_INITIALIZE("SignalR Startup Message - Sleeping for 250ms so SignalR can load", Common.LOG_CATEGORY);

            Thread.Sleep(250);

            Int64 startTicks = Log.APPLICATION_INITIALIZE("Enter", Common.LOG_CATEGORY);

            InitializeComponent();

            // NOTE(crhodes)
            // Try moving Bootstrapper to Common so we can access UnityContainer
            Common.ApplicationBootstrapper = new Application.Bootstrapper();
            Common.ApplicationBootstrapper.Run();

            Log.APPLICATION_INITIALIZE("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #region Event Handlers

        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {

        }


        #endregion

    }
}