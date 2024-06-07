using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VNC.Visio.VSTOAddIn
{
    public class Common : VNC.VSTOAddIn.Common
    {
        public static Microsoft.Office.Interop.Visio.Application VisioApplication { get; set; }

        public static void DisplayInDebugWindow(string outputLine)
        {
#if VNCLOGGING
            Log.APPLICATION($"{outputLine}", Common.LOG_CATEGORY);
#endif
            Common.WriteToDebugWindow($"{outputLine}");
        }
    }
}
