using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace VNC.Visio.AddinHelper
{
    public class Util
    {
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
    }
}
