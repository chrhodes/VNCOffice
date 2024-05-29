using System;
using System.Collections.Generic;
using System.Linq;
using Visio = Microsoft.Office.Interop.Visio;

namespace SupportTools_Visio.Domain
{
    public class ClassInfoShape
    {
        #region Constructors and Load

        public ClassInfoShape(Visio.Shape activeShape)
        {
            // TODO(crhodes)
            // Make this reflect on properties and loop across.

            Class = Helper.GetShapePropertyAsString(activeShape, "Class");
            Namespace = Helper.GetShapePropertyAsString(activeShape, "Namespace");
            Version = Helper.GetShapePropertyAsString(activeShape, "Version");
            Color = Helper.GetShapePropertyAsString(activeShape, "Color");
            Color2 = Helper.GetShapePropertyAsString(activeShape, "Color2");
            GroupName = Helper.GetShapePropertyAsString(activeShape, "GroupName");
            SourceName = Helper.GetShapePropertyAsString(activeShape, "SourceName");
            RootPath = Helper.GetShapePropertyAsString(activeShape, "RootPath");
            AssemblyFileName = Helper.GetShapePropertyAsString(activeShape, "AssemblyFileName");
            SourceFileName = Helper.GetShapePropertyAsString(activeShape, "SourceFileName");
            ApplicationName = Helper.GetShapePropertyAsString(activeShape, "ApplicationName");
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
