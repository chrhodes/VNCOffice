using VNC.Core.Mvvm;

namespace VNC.Visio.VSTOAddIn.Presentation.ModelWrappers
{
    public class GlueInfoWrapper : ModelWrapper<Domain.GlueInfoRow>
    {
        public GlueInfoWrapper()
        {
        }
        public GlueInfoWrapper(Domain.GlueInfoRow model) : base(model)
        {
        }

        public string BegTrigger { get { return GetValue<string>(); } set { SetValue(value); } }
        public string EndTrigger { get { return GetValue<string>(); } set { SetValue(value); } }
        public string GlueType { get { return GetValue<string>(); } set { SetValue(value); } }
        public string WalkPreference { get { return GetValue<string>(); } set { SetValue(value); } }
    }
}
