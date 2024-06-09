using VNC.Core.Mvvm;

using VNCVisioAddIn = VNC.Visio.VSTOAddIn;

namespace SupportTools_Visio.Presentation.ModelWrappers
{
    public class PagePropertiesWrapper : ModelWrapper<VNCVisioAddIn.Domain.PagePropertiesRow>
    {
        public PagePropertiesWrapper()
        {
        }
        public PagePropertiesWrapper(VNCVisioAddIn.Domain.PagePropertiesRow model) : base(model)
        {
        }

        public string PageWidth { get { return GetValue<string>(); } set { SetValue(value); } }
        public string PageHeight { get { return GetValue<string>(); } set { SetValue(value); } }
        public string PageScale { get { return GetValue<string>(); } set { SetValue(value); } }
        public string DrawingScale { get { return GetValue<string>(); } set { SetValue(value); } }
        public string DrawingSizeType { get { return GetValue<string>(); } set { SetValue(value); } }
        public string DrawingScaleType { get { return GetValue<string>(); } set { SetValue(value); } }
        public string DrawingResizeType { get { return GetValue<string>(); } set { SetValue(value); } }
        public string InhibitSnap { get { return GetValue<string>(); } set { SetValue(value); } }
        public string UIVisibility { get { return GetValue<string>(); } set { SetValue(value); } }
        public string PageLockReplace { get { return GetValue<string>(); } set { SetValue(value); } }
        public string PageLockDuplicate { get { return GetValue<string>(); } set { SetValue(value); } }
    }
}
