using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

using SupportTools_Excel.AzureDevOpsExplorer.Domain;
using VNC.Core.Mvvm;

namespace SupportTools_Excel.AzureDevOpsExplorer.Presentation.ModelWrappers
{
    public class WorkItemQueryWrapper : ModelWrapper<WorkItemQuery>
    {
        public WorkItemQueryWrapper() { }
        public WorkItemQueryWrapper(WorkItemQuery model) : base(model)
        {
        }

        public string Name { get { return GetValue<string>(); } set { SetValue(value); } }

        public string QueryWithTokens { get { return GetValue<string>(); } set { SetValue(value); } }

        public string Query { get { return GetValue<string>(); } set { SetValue(value); } }

        public List<string> Fields { get { return GetValue<List<string>>(); } set { SetValue(value); } }

        //public Boolean RetrieveRevisions { get { return GetValue<Boolean>(); } set { SetValue(value); } }

    }
}
