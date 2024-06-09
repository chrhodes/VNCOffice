using SupportTools_Visio.Domain;
using VNC.Core.Mvvm;

using VNCVisioAddIn = VNC.Visio.VSTOAddIn;

namespace SupportTools_Visio.Presentation.ModelWrappers
{
    public class StylePropertiesWrapper : ModelWrapper<VNCVisioAddIn.Domain.StylePropertiesRow>
    {
        public StylePropertiesWrapper(VNCVisioAddIn.Domain.StylePropertiesRow model) : base(model)
        {
        }

        public string EnableTextProps { get { return GetValue<string>(); } set { SetValue(value); } }
        public string EnableLineProps { get { return GetValue<string>(); } set { SetValue(value); } }
        public string EnableFillProps { get { return GetValue<string>(); } set { SetValue(value); } }
        public string HideForApply { get { return GetValue<string>(); } set { SetValue(value); } }
    }
}
