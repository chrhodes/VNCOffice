using VNC.Core.Mvvm;

using VNCVisioAddIn = VNC.Visio.VSTOAddIn;

namespace SupportTools_Visio.Presentation.ModelWrappers
{
    public class GradientPropertiesWrapper : ModelWrapper<VNCVisioAddIn.Domain.GradientProperties>
    {
        public GradientPropertiesWrapper()
        {
        }
        public GradientPropertiesWrapper(VNCVisioAddIn.Domain.GradientProperties model) : base(model)
        {
        }

        public string LineGradientDir { get { return GetValue<string>(); } set { SetValue(value); } }
        public string LineGradientAngle { get { return GetValue<string>(); } set { SetValue(value); } }
        public string FillGradientDir { get { return GetValue<string>(); } set { SetValue(value); } }
        public string FillGradientAngle { get { return GetValue<string>(); } set { SetValue(value); } }
        public string LineGradientEnabled { get { return GetValue<string>(); } set { SetValue(value); } }
        public string FillGradientEnabled { get { return GetValue<string>(); } set { SetValue(value); } }
        public string RotateGradientWithShape { get { return GetValue<string>(); } set { SetValue(value); } }
        public string UseGroupGradient { get { return GetValue<string>(); } set { SetValue(value); } }
    }
}
