using SupportTools_Visio.Domain;
using VNC.Core.Mvvm;

namespace SupportTools_Visio.Presentation.ModelWrappers
{
    public class GlueInfoWrapper : ModelWrapper<Domain.GlueInfo>
    {
        public GlueInfoWrapper()
        {
        }
        public GlueInfoWrapper(GlueInfo model) : base(model)
        {
        }

        public string BegTrigger { get { return GetValue<string>(); } set { SetValue(value); } }
        public string EndTrigger { get { return GetValue<string>(); } set { SetValue(value); } }
        public string GlueType { get { return GetValue<string>(); } set { SetValue(value); } }
        public string WalkPreference { get { return GetValue<string>(); } set { SetValue(value); } }
    }
}
