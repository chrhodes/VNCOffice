using System.Collections.Generic;

namespace SupportTools_Excel.AzureDevOpsExplorer.Domain
{
    public class TeamProjectActionRequest
    {
        // TODO(crhodes)
        // Should we initialize here or handle null elsewhere.
        public List<string> BSSections { get; set; } = new List<string>();
        public List<string> TMSections { get; set; } = new List<string>();
        public List<string> TPSections { get; set; } = new List<string>();
        public List<string> VCSSections { get; set; } = new List<string>();
        public List<string> WISSections { get; set; } = new List<string>();

    }
}
