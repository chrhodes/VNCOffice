using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Visio = Microsoft.Office.Interop.Visio;
using VisioHelper = VNC.Visio.AddinHelper.Visio;

namespace SupportTools_Visio.Actions
{
    class Visio_Application
    {
        #region Main Function Routines

        public static void DisplayInfo()
        {
            VisioHelper.DisplayInWatchWindow(string.Format("{0}()",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));

            Visio.Application app = Globals.ThisAddIn.Application;

            StringBuilder sb = new StringBuilder();

            sb.AppendFormat("{0} - {1}\n", "App.Name", app.Name);

            try
            {
                sb.AppendFormat("{0} - {1}\n", "App.ActiveDocument.Name", app.ActiveDocument.Name);
            }
            catch (Exception ex)
            {
                sb.AppendFormat("{0} - {1}\n", "App.ActiveDocument.Name", "<none>");
            }

            try
            {
                sb.AppendFormat("{0} - {1}\n", "App.ActivePage.Name", app.ActivePage.Name);
            }
            catch (Exception ex)
            {
                sb.AppendFormat("{0} - {1}\n", "App.ActivePage.Name", "<none>");
            }

            sb.AppendFormat("{0} - {1}\n", "App.AddonPaths", app.AddonPaths);
            sb.AppendFormat("{0} - {1}\n", "App.CommandLine", app.CommandLine);
            sb.AppendFormat("{0} - {1}\n", "App.Documents.Count", app.Documents.Count);
            sb.AppendFormat("{0} - {1}\n", "App.DrawingPaths", app.DrawingPaths);
            sb.AppendFormat("{0} - {1}\n", "App.HelpPaths", app.HelpPaths);
            sb.AppendFormat("{0} - {1}\n", "App.IsVisio32", app.IsVisio32);
            sb.AppendFormat("{0} - {1}\n", "App.MyShapesPath", app.MyShapesPath);
            sb.AppendFormat("{0} - {1}\n", "App.Path", app.Path);
            sb.AppendFormat("{0} - {1}\n", "App.ProcessID", app.ProcessID);
            sb.AppendFormat("{0} - {1}\n", "App.ShowChanges", app.ShowChanges);
            sb.AppendFormat("{0} - {1}\n", "App.ShowProgress", app.ShowProgress);
            sb.AppendFormat("{0} - {1}\n", "App.ShowStatusBar", app.ShowStatusBar);
            sb.AppendFormat("{0} - {1}\n", "App.ShowToolBar", app.ShowToolbar);
            sb.AppendFormat("{0} - {1}\n", "App.StartupPaths", app.StartupPaths);
            sb.AppendFormat("{0} - {1}\n", "App.StencilPaths", app.StencilPaths);
            sb.AppendFormat("{0} - {1}\n", "App.TemplatePaths", app.TemplatePaths);
            sb.AppendFormat("{0} - {1}\n", "App.TraceFlags", app.TraceFlags);
            sb.AppendFormat("{0} - {1}\n", "App.UndoEnables", app.UndoEnabled);
            sb.AppendFormat("{0} - {1}\n", "App.UserName", app.UserName);
            sb.AppendFormat("{0} - {1}\n", "App.Version", app.Version);

            //System.Windows.Forms.MessageBox.Show(sb.ToString());
            VisioHelper.DisplayInWatchWindow(sb.ToString());
        }

        public static Visio.Document GetActiveDocument()
        {
            VisioHelper.DisplayInWatchWindow($"{System.Reflection.MethodBase.GetCurrentMethod().Name}");

            Visio.Application app = Globals.ThisAddIn.Application;

            return app.ActiveDocument;
        }

        public static List<Visio.Document> GetOpenDocuments()
        {
            VisioHelper.DisplayInWatchWindow($"{System.Reflection.MethodBase.GetCurrentMethod().Name}");

            Visio.Application app = Globals.ThisAddIn.Application;

            List<Visio.Document> openDocs = new List<Visio.Document>();

            foreach (Visio.Document doc in app.Documents)
            {
                openDocs.Add(doc);
            }

            return openDocs;
        }

        public static void LayerManager()
        {
            VisioHelper.DisplayInWatchWindow(string.Format("{0}({1})",
                System.Reflection.MethodInfo.GetCurrentMethod().Name, "TODO: Not Implemented"));

            // TODO(CHR): Launch WPF Layer Manager Window
        }

        #endregion
    }
}
