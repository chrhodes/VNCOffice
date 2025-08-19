using System;

using MSVisio = Microsoft.Office.Interop.Visio;
using VNCVisioAddIn = VNC.Visio.VSTOAddIn;

namespace VNCVisioToolsApplication.Actions
{
    public class QueryWorkItemInfoShape : WorkItemShapeInfo
    {
        public QueryWorkItemInfoShape(MSVisio.Shape activeShape) : base(activeShape)
        {
            Organization = VNCVisioAddIn.Helpers.GetShapePropertyAsString(activeShape, "Organization");
        }
    }
}