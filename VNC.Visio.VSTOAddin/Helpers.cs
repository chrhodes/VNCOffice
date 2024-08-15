using System;
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
            //fileName = fileName.Replace(" ", "");
            fileName = fileName.Replace(":", "-");

            return fileName;
        }

        public static string SafePageName(string pageName)
        {
            pageName = pageName.Replace("/", "-");
            pageName = pageName.Replace(@"\", "-");
            pageName = pageName.Replace("[", "");
            pageName = pageName.Replace("]", "");
            //pageName = pageName.Replace(" ", "");
            pageName = pageName.Replace("\n", " ");
            pageName = pageName.Replace(":", "-");

            return pageName;
        }

        public static Boolean RowExists(Shape shape, VisSectionIndices sectionIndex, VisRowIndices rowIndex, VisExistsFlags visExistsFlags = VisExistsFlags.visExistsAnywhere)
        {
            if (0 == shape.RowExists[ (short)sectionIndex, (short)rowIndex, (short)visExistsFlags])
            { 
                return false; 
            }
            else 
            { 
                return true; 
            }
        }
    }
}
