using System;
using System.Collections.Generic;
using SupportTools_Excel.Domain;

namespace SupportTools_Excel.AzureDevOpsExplorer.Domain
{
    public class Options_AZDO_TFS : Options_Excel
    {
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        public int GoBackDays { get; set; }
        public bool GetLastActivityDates { get; set; }
        public bool SkipIfNoActivity { get; set; }

        public bool EnableDelays { get; set; } = true;
        public int LoopDelaySeconds { get; set; } = 15;
        public Single ItemDelaySeconds { get; set; } = 0.5F;

        public List<String> TeamProjects { get; set; }
        public List<String> WorkItemTypes { get; set; }

        public bool ShowAllNodeLevels { get; set; }

        public bool ShowIndividualItems { get; set; }
        public int LoopUpdateInterval { get; set; } = 50;

        public Boolean RetrieveRevisions { get; set; }
        public Boolean RetrieveFieldChanges { get; set; }

        public int RecursionLevel { get; set; } = 1;

        public bool ExportXMLTemplate { get; set; }
        public bool IncludeGlobalLists { get; set; }
        public string XMLTemplateFilePath { get; set; } = @"C:\Temp\AZDO-TFS";

        public FormatSpecifications FormatSpecs { get; set; } = new FormatSpecifications();

        public bool ListChangeSetChanges { get; set; }
        public bool ListChangeSetWorkItems { get; set; }

        public WorkItemQuery WorkItemQuerySpec { get; set; }



    }
}
