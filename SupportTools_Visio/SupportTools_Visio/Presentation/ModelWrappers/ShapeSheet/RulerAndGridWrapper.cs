﻿using SupportTools_Visio.Domain;
using VNC.Core.Mvvm;

namespace SupportTools_Visio.Presentation.ModelWrappers
{
    public class RulerAndGridWrapper : ModelWrapper<Domain.RulerAndGrid>
    {
        public RulerAndGridWrapper()
        {
        }
        public RulerAndGridWrapper(RulerAndGrid model) : base(model)
        {
        }

        public string XRulerOrigin { get { return GetValue<string>(); } set { SetValue(value); } }
        public string YRulerOrigin { get { return GetValue<string>(); } set { SetValue(value); } }
        public string XRulerDensity { get { return GetValue<string>(); } set { SetValue(value); } }
        public string YRulerDensity { get { return GetValue<string>(); } set { SetValue(value); } }
        public string XGridOrigin { get { return GetValue<string>(); } set { SetValue(value); } }
        public string YGridOrigin { get { return GetValue<string>(); } set { SetValue(value); } }
        public string XGridDensity { get { return GetValue<string>(); } set { SetValue(value); } }
        public string YGridDensity { get { return GetValue<string>(); } set { SetValue(value); } }
        public string XGridSpacing { get { return GetValue<string>(); } set { SetValue(value); } }
        public string YGridSpacing { get { return GetValue<string>(); } set { SetValue(value); } }
    }
}
