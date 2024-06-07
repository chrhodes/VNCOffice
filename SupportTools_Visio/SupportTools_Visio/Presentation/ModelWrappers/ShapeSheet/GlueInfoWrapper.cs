using SupportTools_Visio.Domain;
using VNC.Core.Mvvm;

using VNCVisioAddIn = VNC.Visio.VSTOAddIn;

namespace SupportTools_Visio.Presentation.ModelWrappers
{
    public class GlueInfoWrapper : ModelWrapper<VNCVisioAddIn.Domain.GlueInfo>
    {
        public GlueInfoWrapper()
        {
        }
        public GlueInfoWrapper(VNCVisioAddIn.Domain.GlueInfo model) : base(model)
        {
        }

        public string BegTrigger { get { return GetValue<string>(); } set { SetValue(value); } }
        public string EndTrigger { get { return GetValue<string>(); } set { SetValue(value); } }
        public string GlueType { get { return GetValue<string>(); } set { SetValue(value); } }
        public string WalkPreference { get { return GetValue<string>(); } set { SetValue(value); } }
    }
}
