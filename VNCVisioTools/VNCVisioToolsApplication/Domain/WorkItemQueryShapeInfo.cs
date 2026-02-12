using MSVisio = Microsoft.Office.Interop.Visio;

using VNCVisioAddIn = VNC.Visio.VSTOAddIn;

namespace VNCVisioToolsApplication.Domain
{
    public class WorkItemQueryShapeInfo : ShapeInfo
    {
        #region Constructors and Load

        public WorkItemQueryShapeInfo(MSVisio.Shape shape) : base(shape)
        {
            Organization = VNCVisioAddIn.Helpers.GetShapePropertyAsString(shape, "Organization");
            TeamProject = VNCVisioAddIn.Helpers.GetShapePropertyAsString(shape, "TeamProject");
            WorkItemType = VNCVisioAddIn.Helpers.GetShapePropertyAsString(shape, "WorkItemType");
            State = VNCVisioAddIn.Helpers.GetShapePropertyAsString(shape, "State");

            ID = VNCVisioAddIn.Helpers.GetShapePropertyAsString(shape, "ID");
        }

        #endregion

        #region Enums, Fields, Properties, Structures

        public string Organization { get; set; }

        public string TeamProject { get; set; }

        public string WorkItemType { get; set; }

        public string State { get; set; }

        public string ID { get; set; }

        //public string CreatedBy { get; set; }
        //public string CreatedDate { get; set; }
        //public string ChangedBy { get; set; }
        //public string ChangedDate { get; set; }

        //public string RelatedLinkCount { get; set; }
        //public string ExternalLinkCount { get; set; }
        //public string RemoteLinkCount { get; set; }
        //public string HyperLinkCount { get; set; }

        #endregion

    }
}
