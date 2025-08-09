using VNC.Core.Mvvm;

namespace VNC.Visio.VSTOAddIn.Presentation.ModelWrappers
{
    public class AlignmentRowWrapper : ModelWrapper<Domain.AlignmentRow>
    {
        public AlignmentRowWrapper()
        {
        }
        public AlignmentRowWrapper(Domain.AlignmentRow model) : base(model)
        {
        }

        // TODO(crhodes)
        // Why are these commented out?

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
