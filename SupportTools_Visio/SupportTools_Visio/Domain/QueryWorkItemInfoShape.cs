using System;

using Visio = Microsoft.Office.Interop.Visio;

namespace SupportTools_Visio.Actions
{
    public class QueryWorkItemInfoShape : WorkItemShapeInfo
    {
        public QueryWorkItemInfoShape(Visio.Shape activeShape) : base(activeShape)
        {
            Organization = Helper.GetShapePropertyAsString(activeShape, "Organization");
        }
    }
}