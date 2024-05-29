using VNC.Core.Mvvm;

namespace SupportTools_Visio.Presentation.ModelWrappers
{
    public class PagePropertiesWrapper : ModelWrapper<Domain.PageProperties>
    {
        public PagePropertiesWrapper()
        {
        }
        public PagePropertiesWrapper(Domain.PageProperties model) : base(model)
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
