using System;
using System.Linq;
using System.Reflection;

using Microsoft.Office.Interop.Excel;

namespace VNCExcelToolsApplication.Events
{
    public class AddInApplicationEvents
    {
        private Application Application;
        public Application ExcelApplication
        {
            get
            {
                return Application;
            }
            set
            {
                if (Application != null)
                {
                    // Should remove all the event handlers;
                }

                Application = value;

                // NOTE(crhodes)
                // There are events that are processed by the Application.
                // Remove the event handler from AppEvents
                // Can still call method for logging, see infra.

                if (Application != null)
                {

                }
            }
        }

        #region Events Handled by Application Code

        // NOTE(crhodes)
        // The ExcelAppEvent handlers will log event to watch window.


        #endregion
    }
}
