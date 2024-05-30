﻿using System.Collections.Generic;
using SupportTools_Excel.AzureDevOpsExplorer.Domain;
using VNC.Core.Mvvm;

namespace SupportTools_Excel.AzureDevOpsExplorer.Presentation.ModelWrappers
{
    public class TestCaseRequestWrapper : ModelWrapper<TestCaseRequest>
    {
        public TestCaseRequestWrapper() { }
        public TestCaseRequestWrapper(TestCaseRequest model) : base(model)
        {
        }

        public string TestID { get { return GetValue<string>(); } set { SetValue(value); } }
        public List<string> TestSections { get { return GetValue<List<string>>(); } set { SetValue(value); } }
    }
}
