using SupportTools_Visio.Domain;
using VNC.Core.Mvvm;

namespace SupportTools_Visio.Presentation.ModelWrappers
{
    public class FillGradientStopRowWrapper : ModelWrapper<FillGradientStopRow>
    {
        public FillGradientStopRowWrapper(FillGradientStopRow model) : base(model)
        {
        }

        public string Color { get { return GetValue<string>(); } set { SetValue(value); } }
        public string ColorTrans { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Position { get { return GetValue<string>(); } set { SetValue(value); } }
    }
}
