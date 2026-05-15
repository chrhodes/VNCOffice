using System.Collections.Generic;

using VNC.Core.Mvvm;

using SupportTools_Excel.Domain;
using SupportTools_Excel.AzureDevOpsExplorer.Domain;

namespace SupportTools_Excel.AzureDevOpsExplorer.Presentation.ModelWrappers
{
    public class WorkItemActionRequestWrapper : ModelWrapper<WorkItemActionRequest>
    {
        public WorkItemActionRequestWrapper() { }
        public WorkItemActionRequestWrapper(WorkItemActionRequest model) : base(model)
        {
        }

        public string WorkItemID { get { return GetValue<string>(); } set { SetValue(value); } }

        public List<string> WorkItemSections { get { return GetValue<List<string>>(); } set { SetValue(value); } }

        public bool RetrieveAllWorkItemFieldData { get { return GetValue<bool>(); } set { SetValue(value); } }

        public List<string> WorkItemFields { get { return GetValue<List<string>>(); } set { SetValue(value); } }
    }
}
