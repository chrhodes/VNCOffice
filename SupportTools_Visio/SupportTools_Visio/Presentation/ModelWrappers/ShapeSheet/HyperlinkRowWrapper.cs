﻿using SupportTools_Visio.Domain;
using VNC.Core.Mvvm;

using VNCVisioAddIn = VNC.Visio.VSTOAddIn;

namespace SupportTools_Visio.Presentation.ModelWrappers
{
    public class HyperlinkRowWrapper : ModelWrapper<VNCVisioAddIn.Domain.HyperlinkRow>
    {
        public HyperlinkRowWrapper() { }

        public HyperlinkRowWrapper(VNCVisioAddIn.Domain.HyperlinkRow model) : base(model)
        {
        }

        public string Name { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Description { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Address { get { return GetValue<string>(); } set { SetValue(value); } }
        public string SubAddress { get { return GetValue<string>(); } set { SetValue(value); } }
        public string ExtraInfo { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Frame { get { return GetValue<string>(); } set { SetValue(value); } }
        public string SortKey { get { return GetValue<string>(); } set { SetValue(value); } }
        public string NewWindow { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Default { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Invisible { get { return GetValue<string>(); } set { SetValue(value); } }
    }
}
