using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using MSVisio = Microsoft.Office.Interop.Visio;
using VNCVisioAddIn = VNC.Visio.VSTOAddIn;

namespace VNCVisioToolsApplication.Actions
{
    public class Visio_Application
    {
        public static void DisplayInfo()
        {
            Common.WriteToDebugWindow($"{System.Reflection.MethodInfo.GetCurrentMethod().Name}");

            MSVisio.Application app = Common.VisioApplication;

            StringBuilder sb = new StringBuilder();

            Common.WriteToDebugWindow($"App.Name - {app.Name}");

            try
            {
                Common.WriteToDebugWindow($"App.ActiveDocument.Name - {app.ActiveDocument.Name}");
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow("App.ActiveDocument.Name - <none>");
            }

            try
            {
                Common.WriteToDebugWindow($"App.ActivePage.Name - {app.ActivePage.Name}");
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow("App.ActivePage.Name - <none>");
            }

            Common.WriteToDebugWindow($"App.AddonPaths - {app.AddonPaths}");
            Common.WriteToDebugWindow($"App.CommandLine - {app.CommandLine}");
            Common.WriteToDebugWindow($"App.Documents.Count - {app.Documents.Count}");
            Common.WriteToDebugWindow($"App.DrawingPaths - {app.DrawingPaths}");
            Common.WriteToDebugWindow($"App.HelpPaths - {app.HelpPaths}");
            Common.WriteToDebugWindow($"App.IsVisio32 - {app.IsVisio32}");
            Common.WriteToDebugWindow($"App.MyShapesPath - {app.MyShapesPath}");
            Common.WriteToDebugWindow($"App.Path - {app.Path}");
            Common.WriteToDebugWindow($"App.ProcessID - {app.ProcessID}");
            Common.WriteToDebugWindow($"App.ShowChanges - {app.ShowChanges}");
            Common.WriteToDebugWindow($"App.ShowProgress - {app.ShowProgress}");
            Common.WriteToDebugWindow($"App.ShowStatusBar - {app.ShowStatusBar}");
            Common.WriteToDebugWindow($"App.ShowToolBar - {app.ShowToolbar}");
            Common.WriteToDebugWindow($"App.StartupPaths - {app.StartupPaths}");
            Common.WriteToDebugWindow($"App.StencilPaths - {app.StencilPaths}");
            Common.WriteToDebugWindow($"App.TemplatePaths - {app.TemplatePaths}");
            Common.WriteToDebugWindow($"App.TraceFlags - {app.TraceFlags}");
            Common.WriteToDebugWindow($"App.UndoEnables - {app.UndoEnabled}");
            Common.WriteToDebugWindow($"App.UserName - {app.UserName}");
            Common.WriteToDebugWindow($"App.Version - {app.Version}");

            //System.Windows.Forms.MessageBox.Show(sb.ToString());
            Common.WriteToDebugWindow(sb.ToString());
        }

        public static MSVisio.Document GetActiveDocument()
        {
            Common.WriteToDebugWindow($"{System.Reflection.MethodBase.GetCurrentMethod().Name}");

            MSVisio.Application app = Common.VisioApplication;

            return app.ActiveDocument;
        }

        public static List<MSVisio.Document> GetOpenDocuments()
        {
            Common.WriteToDebugWindow($"{System.Reflection.MethodBase.GetCurrentMethod().Name}");

            MSVisio.Application app = Common.VisioApplication;

            List<MSVisio.Document> openDocs = new List<MSVisio.Document>();

            foreach (MSVisio.Document doc in app.Documents)
            {
                openDocs.Add(doc);
            }

            return openDocs;
        }

        public static void LayerManager()
        {
            Common.WriteToDebugWindow(string.Format("{0}({1})",
                System.Reflection.MethodInfo.GetCurrentMethod().Name, "TODO: Not Implemented"));

            // TODO(CHR): Launch WPF Layer Manager Window
        }

    }
}
