using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SupportTools_Excel.AzureDevOpsExplorer.Domain
{
    public class TestCaseRequest
    {
        public string TestID { get; set; }
        public List<string> TestSections { get; set; } = new List<string>();
    }
}
