﻿using SupportTools_Visio.Domain;
using VNC.Core.Mvvm;

using VNCVisioAddIn = VNC.Visio.VSTOAddIn;

namespace SupportTools_Visio.Presentation.ModelWrappers
{
    public class AlignmentRowWrapper : ModelWrapper<VNCVisioAddIn.Domain.AlignmentRow>
    {
        public AlignmentRowWrapper()
        {
        }
        public AlignmentRowWrapper(VNCVisioAddIn.Domain.AlignmentRow model) : base(model)
        {
        }

        // public string BevelTopType { get { return GetValue<string>(); } set { SetValue(value); } }
        // public string BevelTopWidth { get { return GetValue<string>(); } set { SetValue(value); } }
        // public string BevelTopHeight { get { return GetValue<string>(); } set { SetValue(value); } }
        // public string BevelBottomType { get { return GetValue<string>(); } set { SetValue(value); } }
        // public string BevelBottomWidth { get { return GetValue<string>(); } set { SetValue(value); } }
        // public string BevelBottomHeight { get { return GetValue<string>(); } set { SetValue(value); } }
        // public string BevelDepthColor { get { return GetValue<string>(); } set { SetValue(value); } }
        // public string BevelDepthSize { get { return GetValue<string>(); } set { SetValue(value); } }
        // public string BevelContourColor { get { return GetValue<string>(); } set { SetValue(value); } }
        // public string BevelContourSize { get { return GetValue<string>(); } set { SetValue(value); } }
        // public string BevelMaterialType { get { return GetValue<string>(); } set { SetValue(value); } }
        // public string BevelLightingType { get { return GetValue<string>(); } set { SetValue(value); } }
        // public string BevelLightingAngle { get { return GetValue<string>(); } set { SetValue(value); } }
    }
}
