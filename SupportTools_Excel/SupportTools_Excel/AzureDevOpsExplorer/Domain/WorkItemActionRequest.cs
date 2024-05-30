using System.Collections.Generic;

namespace SupportTools_Excel.AzureDevOpsExplorer.Domain
{
    public class WorkItemActionRequest
    {
        public string WorkItemID { get; set; }

        public List<string> WorkItemSections { get; set; } = new List<string>();

        public bool RetrieveAllWorkItemFieldData { get; set; } = false;

        public List<string> WorkItemFields { get; set; } = new List<string>();

    }
}
 