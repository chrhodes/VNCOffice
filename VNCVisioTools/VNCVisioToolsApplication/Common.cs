using System.Windows;

using Prism.Events;

using MSVisio = Microsoft.Office.Interop.Visio;

namespace VNCVisioToolsApplication
{
    public class Common : VNC.VSTOAddIn.Common
    {
        new public const string LOG_CATEGORY = "VNCVisioToolsApplication";

        public static Events.VisioAppEvents AppEvents;
        public static Events.AddInApplicationEvents AddInApplicationEvents;

        public static IEventAggregator EventAggregator = new EventAggregator();
        public static VNCVisioToolsApplication.Bootstrapper ApplicationBootstrapper;

        public static MSVisio.Application VisioApplication { get; set; }

    }
}
