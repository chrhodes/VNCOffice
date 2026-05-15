using VNC.Core.Mvvm;
using SupportTools_Excel.Domain;
using System.Collections.Generic;
using SupportTools_Excel.AzureDevOpsExplorer.Domain;

namespace SupportTools_Excel.AzureDevOpsExplorer.Presentation.ModelWrappers
{
    public class TeamProjectActionRequestWrapper : ModelWrapper<TeamProjectActionRequest>
    {
        public TeamProjectActionRequestWrapper() { }
        public TeamProjectActionRequestWrapper(TeamProjectActionRequest model) : base(model)
        {
        }

        // TODO(crhodes)
        // Wrap each property from the passed in model.

        public List<string> BSSections { get { return GetValue<List<string>>(); } set { SetValue(value); } }
        public List<string> TMSections { get { return GetValue<List<string>>(); } set { SetValue(value); } }
        public List<string> TPSections { get { return GetValue<List<string>>(); } set { SetValue(value); } }
        public List<string> VCSSections { get { return GetValue<List<string>>(); } set { SetValue(value); } }
        public List<string> WISSections { get { return GetValue<List<string>>(); } set { SetValue(value); } }
    }
}
