using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

using Microsoft.Office.Interop.Visio;

namespace VNC.Visio.VSTOAddIn
{
    public class Helpers
    {
        public static bool LoadStencil(Microsoft.Office.Interop.Visio.Application app, string stencilName)
        {
            bool result = false;

            try
            {
                var foo = app.Documents[stencilName];
                result = true;
            }
            catch (Exception)
            {
                // Stencil may not be open.  Try opening it

                try
                {
                    app.Documents.OpenEx(stencilName, (short)VisOpenSaveArgs.visOpenRO + (short)VisOpenSaveArgs.visOpenDocked);
                    result = true;
                }
                catch (Exception)
                {
                    MessageBox.Show($"Cannot locate or open {stencilName}, aborting.");
                }
            }

            return result;
        }
        public static string SafeFileName(string fileName)
        {
            fileName = fileName.Replace("/", "-");
            fileName = fileName.Replace(@"\", "-");
            fileName = fileName.Replace("[", "");
            fileName = fileName.Replace("]", "");
            fileName = fileName.Replace(" ", "");
            fileName = fileName.Replace(":", "-");

            return fileName;
        }

        public static string SafePageName(string pageName)
        {
            pageName = pageName.Replace("/", "-");
            pageName = pageName.Replace(@"\", "-");
            pageName = pageName.Replace("[", "");
            pageName = pageName.Replace("]", "");
            pageName = pageName.Replace(" ", "");
            pageName = pageName.Replace(":", "-");

            return pageName;
        }
    }
}
