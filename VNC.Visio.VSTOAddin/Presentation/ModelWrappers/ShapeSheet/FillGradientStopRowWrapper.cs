using VNC.Core.Mvvm;

namespace VNC.Visio.VSTOAddIn.Presentation.ModelWrappers
{
    public class FillGradientStopRowWrapper : ModelWrapper<Domain.FillGradientStopRow>
    {
        public FillGradientStopRowWrapper() { }

        public FillGradientStopRowWrapper(Domain.FillGradientStopRow model) : base(model)
        {
        }

        public string Color { get { return GetValue<string>(); } set { SetValue(value); } }
        public string ColorTrans { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Position { get { return GetValue<string>(); } set { SetValue(value); } }
    }
}
