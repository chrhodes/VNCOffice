using VNCVisioToolsApplication.Domain;
using VNC.Core.Mvvm;

using VNCVisioAddIn = VNC.Visio.VSTOAddIn;

namespace VNCVisioToolsApplication.Presentation.ModelWrappers
{
    public class FillGradientStopRowWrapper : ModelWrapper<VNCVisioAddIn.Domain.FillGradientStopRow>
    {
        public FillGradientStopRowWrapper() { }

        public FillGradientStopRowWrapper(VNCVisioAddIn.Domain.FillGradientStopRow model) : base(model)
        {
        }

        public string Color { get { return GetValue<string>(); } set { SetValue(value); } }
        public string ColorTrans { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Position { get { return GetValue<string>(); } set { SetValue(value); } }
    }
}
