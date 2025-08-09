using VNC.Core.Mvvm;

using VNCVisioAddIn = VNC.Visio.VSTOAddIn;

namespace SupportTools_Visio.Presentation.ModelWrappers
{
    public class LayerMembershipWrapper : ModelWrapper<VNCVisioAddIn.Domain.LayerMembershipRow>
    {
        public LayerMembershipWrapper()
        {
        }
        public LayerMembershipWrapper(VNCVisioAddIn.Domain.LayerMembershipRow model) : base(model)
        {
        } 
        public string Name { get { return GetValue<string>(); } set { SetValue(value); } }
    }
}
