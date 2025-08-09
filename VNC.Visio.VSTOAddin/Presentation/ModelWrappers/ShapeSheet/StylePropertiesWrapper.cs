using VNC.Core.Mvvm;

namespace VNC.Visio.VSTOAddIn.Presentation.ModelWrappers
{
    public class StylePropertiesWrapper : ModelWrapper<Domain.StylePropertiesRow>
    {
        public StylePropertiesWrapper(Domain.StylePropertiesRow model) : base(model)
        {
        }

        public string EnableTextProps { get { return GetValue<string>(); } set { SetValue(value); } }
        public string EnableLineProps { get { return GetValue<string>(); } set { SetValue(value); } }
        public string EnableFillProps { get { return GetValue<string>(); } set { SetValue(value); } }
        public string HideForApply { get { return GetValue<string>(); } set { SetValue(value); } }
    }
}
