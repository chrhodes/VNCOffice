
using VNC.Core.Mvvm;

namespace VNC.Visio.VSTOAddIn.Presentation.ModelWrappers
{
    public class LineGradientStopRowWrapper : ModelWrapper<Domain.LineGradientStopRow>
    {
        public LineGradientStopRowWrapper() { }

        public LineGradientStopRowWrapper(Domain.LineGradientStopRow model) : base(model)
        {
        }

        public string Name { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Color { get { return GetValue<string>(); } set { SetValue(value); } }
        public string ColorTrans { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Position { get { return GetValue<string>(); } set { SetValue(value); } }
    }
}
