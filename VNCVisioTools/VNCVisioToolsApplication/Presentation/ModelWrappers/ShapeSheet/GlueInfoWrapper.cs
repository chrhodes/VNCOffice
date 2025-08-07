using VNCVisioToolsApplication.Domain;
using VNC.Core.Mvvm;

using VNCVisioAddIn = VNC.Visio.VSTOAddIn;

namespace VNCVisioToolsApplication.Presentation.ModelWrappers
{
    public class GlueInfoWrapper : ModelWrapper<VNCVisioAddIn.Domain.GlueInfoRow>
    {
        public GlueInfoWrapper()
        {
        }
        public GlueInfoWrapper(VNCVisioAddIn.Domain.GlueInfoRow model) : base(model)
        {
        }

        public string BegTrigger { get { return GetValue<string>(); } set { SetValue(value); } }
        public string EndTrigger { get { return GetValue<string>(); } set { SetValue(value); } }
        public string GlueType { get { return GetValue<string>(); } set { SetValue(value); } }
        public string WalkPreference { get { return GetValue<string>(); } set { SetValue(value); } }
    }
}
