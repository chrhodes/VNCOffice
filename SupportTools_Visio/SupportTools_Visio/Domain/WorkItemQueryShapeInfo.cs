using Visio = Microsoft.Office.Interop.Visio;

namespace SupportTools_Visio.Actions
{
    public class WorkItemQueryShapeInfo : ShapeInfo
    {
        #region Constructors and Load

        public WorkItemQueryShapeInfo(Visio.Shape shape) : base(shape)
        {
            Organization = Helper.GetShapePropertyAsString(shape, "Organization");
            TeamProject = Helper.GetShapePropertyAsString(shape, "TeamProject");
            WorkItemType = Helper.GetShapePropertyAsString(shape, "WorkItemType");
            State = Helper.GetShapePropertyAsString(shape, "State");

            ID = Helper.GetShapePropertyAsString(shape, "ID");
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
