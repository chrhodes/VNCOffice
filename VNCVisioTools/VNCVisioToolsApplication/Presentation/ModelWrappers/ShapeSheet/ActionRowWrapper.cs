using VNCVisioToolsApplication.Domain;
using VNC.Core.Mvvm;

using VNCVisioAddIn = VNC.Visio.VSTOAddIn;

namespace VNCVisioToolsApplication.Presentation.ModelWrappers
{
    public class ActionRowWrapper : ModelWrapper<VNCVisioAddIn.Domain.ActionRow>
    {
        public ActionRowWrapper() { }

        public ActionRowWrapper(VNCVisioAddIn.Domain.ActionRow model) : base(model)
        {
        }

        public string Name { get { return GetValue<string>(); } set { SetValue(value); } }

        public string Action { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Menu { get { return GetValue<string>(); } set { SetValue(value); } }
        public string TagName { get { return GetValue<string>(); } set { SetValue(value); } }
        public string ButtonFace { get { return GetValue<string>(); } set { SetValue(value); } }
        public string SortKey { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Checked { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Disabled { get { return GetValue<string>(); } set { SetValue(value); } }
        public string ReadOnly { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Invisible { get { return GetValue<string>(); } set { SetValue(value); } }
        public string BeginGroup { get { return GetValue<string>(); } set { SetValue(value); } }
        public string FlyoutChild { get { return GetValue<string>(); } set { SetValue(value); } }
    }
}
