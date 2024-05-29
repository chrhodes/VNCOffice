using SupportTools_Visio.Domain;
using VNC.Core.Mvvm;

namespace SupportTools_Visio.Presentation.ModelWrappers
{
    public class ControlsRowWrapper : ModelWrapper<Domain.ControlsRow>
    {
        public ControlsRowWrapper(ControlsRow model) : base(model)
        {
        }

        public string Name { get { return GetValue<string>(); } set { SetValue(value); } }

        public string X { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Y { get { return GetValue<string>(); } set { SetValue(value); } }
        public string XDynamics { get { return GetValue<string>(); } set { SetValue(value); } }
        public string YDynamics { get { return GetValue<string>(); } set { SetValue(value); } }
        public string XBehavior { get { return GetValue<string>(); } set { SetValue(value); } }
        public string YBehavior { get { return GetValue<string>(); } set { SetValue(value); } }
        public string CanGlue { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Tip { get { return GetValue<string>(); } set { SetValue(value); } }
    }
}
