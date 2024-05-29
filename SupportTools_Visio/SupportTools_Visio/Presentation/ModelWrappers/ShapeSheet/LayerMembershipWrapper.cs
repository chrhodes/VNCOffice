using SupportTools_Visio.Domain;
using VNC.Core.Mvvm;

namespace SupportTools_Visio.Presentation.ModelWrappers
{
    public class LayerMembershipWrapper : ModelWrapper<Domain.LayerMembership>
    {
        public LayerMembershipWrapper()
        {
        }
        public LayerMembershipWrapper(LayerMembership model) : base(model)
        {
        } 
        public string Name { get { return GetValue<string>(); } set { SetValue(value); } }
    }
}
