using System;
using System.Collections.Generic;
using System.Linq;

using MSVisio = Microsoft.Office.Interop.Visio;

using VNCVisioAddIn = VNC.Visio.VSTOAddIn;

namespace VNCVisioToolsApplication.Domain
{
    public class ClassInfoShape
    {
        #region Constructors and Load

        public ClassInfoShape(MSVisio.Shape activeShape)
        {
            // TODO(crhodes)
            // Make this reflect on properties and loop across.

            Class = VNCVisioAddIn.Helpers.GetShapePropertyAsString(activeShape, "Class");
            Namespace = VNCVisioAddIn.Helpers.GetShapePropertyAsString(activeShape, "Namespace");
            Version = VNCVisioAddIn.Helpers.GetShapePropertyAsString(activeShape, "Version");
            Color = VNCVisioAddIn.Helpers.GetShapePropertyAsString(activeShape, "Color");
            Color2 = VNCVisioAddIn.Helpers.GetShapePropertyAsString(activeShape, "Color2");
            GroupName = VNCVisioAddIn.Helpers.GetShapePropertyAsString(activeShape, "GroupName");
            SourceName = VNCVisioAddIn.Helpers.GetShapePropertyAsString(activeShape, "SourceName");
            RootPath = VNCVisioAddIn.Helpers.GetShapePropertyAsString(activeShape, "RootPath");
            AssemblyFileName = VNCVisioAddIn.Helpers.GetShapePropertyAsString(activeShape, "AssemblyFileName");
            SourceFileName = VNCVisioAddIn.Helpers.GetShapePropertyAsString(activeShape, "SourceFileName");
            ApplicationName = VNCVisioAddIn.Helpers.GetShapePropertyAsString(activeShape, "ApplicationName");
        }

        #endregion

        #region Enums, Fields, Properties, Structures

        public string ApplicationName { get; set; }
        public string AssemblyFileName { get; set; }
        public string Class { get; set; }
        public string Color { get; set; }
        public string Color2 { get; set; }
        public string GroupName { get; set; }
        public string Namespace { get; set; }
        public string RootPath { get; set; }
        public string SourceFileName { get; set; }
        public string SourceName { get; set; }
        public string Version { get; set; }

        #endregion

        #region Main Methods

        public override string ToString()
        {
            return string.Format("{0} {1} {2} {3} {4} {5} {6} {7} {8} {9} {10}",
                Class, Namespace, Version, Color, Color2, GroupName, SourceName, RootPath, AssemblyFileName, SourceFileName, ApplicationName);
        }

        #endregion
    }
}
