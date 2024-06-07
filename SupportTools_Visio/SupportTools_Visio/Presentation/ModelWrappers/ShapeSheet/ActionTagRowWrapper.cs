using SupportTools_Visio.Domain;
using VNC.Core.Mvvm;

using VNCVisioAddIn = VNC.Visio.VSTOAddIn;

namespace SupportTools_Visio.Presentation.ModelWrappers
{
    public class ActionTagRowWrapper : ModelWrapper<VNCVisioAddIn.Domain.ActionTagRow>
    {
        public ActionTagRowWrapper() { }
        public ActionTagRowWrapper(VNCVisioAddIn.Domain.ActionTagRow model) : base(model)
        {
        }

        public string Name { get { return GetValue<string>(); } set { SetValue(value); } }

        public string X { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Y { get { return GetValue<string>(); } set { SetValue(value); } }
        public string TagName { get { return GetValue<string>(); } set { SetValue(value); } }
        public string XJustify { get { return GetValue<string>(); } set { SetValue(value); } }
        public string YJustify { get { return GetValue<string>(); } set { SetValue(value); } }
        public string DisplayMode { get { return GetValue<string>(); } set { SetValue(value); } }
        public string ButtonFace { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Description { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Disabled { get { return GetValue<string>(); } set { SetValue(value); } }
    }
}
