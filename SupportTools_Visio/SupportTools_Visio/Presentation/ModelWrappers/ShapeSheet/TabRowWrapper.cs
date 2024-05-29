﻿using SupportTools_Visio.Domain;
using VNC.Core.Mvvm;

namespace SupportTools_Visio.Presentation.ModelWrappers
{
    public class TabRowWrapper : ModelWrapper<Domain.TabRow>
    {
        public TabRowWrapper(TabRow model) : base(model)
        {
        }

        // TODO(crhodes)
        // This looks tricky as there are an unknown number of tabs
        public string Name { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Position1 { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Alignment1 { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Position2 { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Alignment2 { get { return GetValue<string>(); } set { SetValue(value); } }
    }
}
