
using VNC.Core.Mvvm;

using VNCVisioAddIn = VNC.Visio.VSTOAddIn;

namespace SupportTools_Visio.Presentation.ModelWrappers
{
    public class LineGradientStopRowWrapper : ModelWrapper<VNCVisioAddIn.Domain.LineGradientStopRow>
    {
        public LineGradientStopRowWrapper() { }

        public LineGradientStopRowWrapper(VNCVisioAddIn.Domain.LineGradientStopRow model) : base(model)
        {
        }

        public string Name { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Color { get { return GetValue<string>(); } set { SetValue(value); } }
        public string ColorTrans { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Position { get { return GetValue<string>(); } set { SetValue(value); } }
    }
}
