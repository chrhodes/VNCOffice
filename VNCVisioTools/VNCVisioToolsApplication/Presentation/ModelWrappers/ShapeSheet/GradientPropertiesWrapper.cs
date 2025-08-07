using VNC.Core.Mvvm;

using VNCVisioAddIn = VNC.Visio.VSTOAddIn;

namespace VNCVisioToolsApplication.Presentation.ModelWrappers
{
    public class GradientPropertiesWrapper : ModelWrapper<VNCVisioAddIn.Domain.GradientPropertiesRow>
    {
        public GradientPropertiesWrapper()
        {
        }
        public GradientPropertiesWrapper(VNCVisioAddIn.Domain.GradientPropertiesRow model) : base(model)
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
