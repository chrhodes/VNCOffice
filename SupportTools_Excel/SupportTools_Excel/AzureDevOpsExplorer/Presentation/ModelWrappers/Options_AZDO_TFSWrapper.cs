using SupportTools_Excel.AzureDevOpsExplorer.Domain;
using SupportTools_Excel.Domain;
using System;
using System.Collections.Generic;
using VNC.Core.Mvvm;

namespace SupportTools_Excel.AzureDevOpsExplorer.Presentation.ModelWrappers
{
    public class Options_AZDO_TFSWrapper : ModelWrapper<Options_AZDO_TFS>
    {
        public Options_AZDO_TFSWrapper() { }
        public Options_AZDO_TFSWrapper(Options_AZDO_TFS model) : base(model)
        {
        }

        public DateTime StartDate { get { return GetValue<DateTime>(); } set { SetValue(value); } }
        public DateTime EndDate { get { return GetValue<DateTime>(); } set { SetValue(value); } }
        public int GoBackDays { get { return GetValue<int>(); } set { SetValue(value); } }
        public bool GetLastActivityDates { get { return GetValue<bool>(); } set { SetValue(value); } }
        public bool SkipIfNoActivity { get { return GetValue<bool>(); } set { SetValue(value); } }

        public bool EnableDelays { get { return GetValue<bool>(); } set { SetValue(value); } }
        public int LoopDelaySeconds { get { return GetValue<int>(); } set { SetValue(value); } }
        public Single ItemDelaySeconds { get { return GetValue<Single>(); } set { SetValue(value); } }

        public List<String> TeamProjects { get { return GetValue<List<String>>(); } set { SetValue(value); } }
        public List<String> WorkItemTypes { get { return GetValue<List<String>>(); } set { SetValue(value); } }

        public bool ShowAllNodeLevels { get { return GetValue<bool>(); } set { SetValue(value); } }

        public bool ShowIndividualItems { get { return GetValue<bool>(); } set { SetValue(value); } }
        public int LoopUpdateInterval { get { return GetValue<int>(); } set { SetValue(value); } }

        public Boolean RetrieveRevisions { get { return GetValue<Boolean>(); } set { SetValue(value); } }
        public Boolean RetrieveFieldChanges { get { return GetValue<Boolean>(); } set { SetValue(value); } }

        public int RecursionLevel { get { return GetValue<int>(); } set { SetValue(value); } }

        //public bool ShowWorkItemFieldData { get { return GetValue<bool>(); } set { SetValue(value); } }

        public bool ExportXMLTemplate { get { return GetValue<bool>(); } set { SetValue(value); } }
        public bool IncludeGlobalLists { get { return GetValue<bool>(); } set { SetValue(value); } }
        public string XMLTemplateFilePath { get { return GetValue<string>(); } set { SetValue(value); } }

        public FormatSpecifications FormatSpecs { get { return GetValue<FormatSpecifications>(); } set { SetValue(value); } }

        public bool ListChangeSetChanges { get { return GetValue<bool>(); } set { SetValue(value); } }
        public bool ListChangeSetWorkItems { get { return GetValue<bool>(); } set { SetValue(value); } }

        public WorkItemQuery WorkItemQuerySpec { get { return GetValue<WorkItemQuery>(); } set { SetValue(value); } }
    }
}
