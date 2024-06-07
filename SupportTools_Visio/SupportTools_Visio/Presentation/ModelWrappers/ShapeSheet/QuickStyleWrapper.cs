using SupportTools_Visio.Domain;
using VNC.Core.Mvvm;

using VNCVisioAddIn = VNC.Visio.VSTOAddIn;

namespace SupportTools_Visio.Presentation.ModelWrappers
{
    public class QuickStyleWrapper : ModelWrapper<VNCVisioAddIn.Domain.QuickStyle>
    {
        public QuickStyleWrapper()
        {
        }
        public QuickStyleWrapper(VNCVisioAddIn.Domain.QuickStyle model) : base(model)
        {
        }

        public string QuickStyleLineMatrix { get { return GetValue<string>(); } set { SetValue(value); } }
        public string QuickStyleLineColor { get { return GetValue<string>(); } set { SetValue(value); } }
        public string QuickStyleFontColor { get { return GetValue<string>(); } set { SetValue(value); } }
        public string QuickStyleVariation { get { return GetValue<string>(); } set { SetValue(value); } }
        public string QuickStyleFillMatrix { get { return GetValue<string>(); } set { SetValue(value); } }
        public string QuickStyleFontMatrix { get { return GetValue<string>(); } set { SetValue(value); } }
        public string QuickStyleEffectsMatrix { get { return GetValue<string>(); } set { SetValue(value); } }
        public string QuickStyleShadowColor { get { return GetValue<string>(); } set { SetValue(value); } }
        public string QuickStyleType { get { return GetValue<string>(); } set { SetValue(value); } }
    }
}
