using VNC.Core.Mvvm;

namespace VNC.Visio.VSTOAddIn.Presentation.ModelWrappers
{
    public class LayerMembershipWrapper : ModelWrapper<Domain.LayerMembershipRow>
    {
        public LayerMembershipWrapper()
        {
        }
        public LayerMembershipWrapper(Domain.LayerMembershipRow model) : base(model)
        {
        } 
        public string Name { get { return GetValue<string>(); } set { SetValue(value); } }
    }
}
