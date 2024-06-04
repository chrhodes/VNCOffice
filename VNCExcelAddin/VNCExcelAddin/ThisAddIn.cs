using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Interop.Excel;
using VNC;

namespace VNCExcelAddin
{
    public partial class ThisAddIn
    {
        private static System.Windows.Application _XamlApp;

        private static Prism.Unity.PrismApplication _prismApplication;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            long startTicks = Log.APPLICATION_START("Enter", Common.LOG_CATEGORY);

            try
            {
                // Not in SupportTools_Visio
                //Common.DeveloperMode = true;
                //Common.WriteToDebugWindow("ThisAddIn_Startup()");
                //Common.DeveloperMode = false;

                Globals.Ribbons.Ribbon.chkDisplayEvents.Checked = Common.DisplayEvents;
                Globals.Ribbons.Ribbon.chkEnableAppEvents.Checked = Common.HasAppEvents;

                if (Common.HasAppEvents)
                {
                    Common.AppEvents = new Events.ExcelAppEvents();
                    Common.AppEvents.ExcelApplication = Globals.ThisAddIn.Application;
                }

                XlHlp.ExcelApplication = Globals.ThisAddIn.Application;

                // For the WPF stuff we do

                InitializeWPFApplication();
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);

                Common.DeveloperMode = true;
                Common.WriteToDebugWindow(ex.ToString());
                Common.DeveloperMode = false;

                throw (ex);
            }

            Log.APPLICATION_START("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {

            long startTicks = Log.APPLICATION_END("Enter", Common.LOG_CATEGORY);

            Log.APPLICATION_END("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #region Prism and WPF support

        #endregion

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
