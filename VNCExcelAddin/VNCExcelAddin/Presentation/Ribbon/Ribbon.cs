using System;
using System.Threading;
using System.Windows;

using Microsoft.Office.Tools.Ribbon;
using Prism.Unity;
using VNCExcelAddin.Presentation.ViewModels;
using VNCExcelAddin.Presentation.Views;

using VNC;
using VNC.Core.Presentation;
using VNC.WPF.Presentation.Dx.Views;
using VNC.WPF.Presentation.Views;

using ExcelHlp = VNC.AddinHelper.Excel;
using VNCHlp = VNC.AddinHelper;

namespace VNCExcelAddin
{
    public partial class Ribbon
    {
        // NOTE(crhodes)
        // This was moved out of designer so we can call bootstrapper.

        public Ribbon()
           : base(Globals.Factory.GetRibbonFactory())
        {
            Log.APPLICATION_INITIALIZE("SignalR Startup Message - Sleeping for 250ms so SignalR can load", Common.LOG_CATEGORY);
            // HACK(crhodes)
            // See if this helps logging first few messages
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
