using VNC.Core.Mvvm;

namespace VNC.Visio.VSTOAddIn.Presentation.ModelWrappers
{
    public class GradientPropertiesWrapper : ModelWrapper<Domain.GradientPropertiesRow>
    {
        public GradientPropertiesWrapper()
        {
        }
        public GradientPropertiesWrapper(Domain.GradientPropertiesRow model) : base(model)
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
