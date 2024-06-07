using VNC.Core.Mvvm;

using VNCVisioAddIn = VNC.Visio.VSTOAddIn;

namespace SupportTools_Visio.Presentation.ModelWrappers
{
    public class LayerMembershipWrapper : ModelWrapper<VNCVisioAddIn.Domain.LayerMembership>
    {
        public LayerMembershipWrapper()
        {
        }
        public LayerMembershipWrapper(VNCVisioAddIn.Domain.LayerMembership model) : base(model)
        {
        } 
        public string Name { get { return GetValue<string>(); } set { SetValue(value); } }
    }
}
