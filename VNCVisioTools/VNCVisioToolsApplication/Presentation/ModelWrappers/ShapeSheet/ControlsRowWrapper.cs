using VNCVisioToolsApplication.Domain;
using VNC.Core.Mvvm;

using VNCVisioAddIn = VNC.Visio.VSTOAddIn;

namespace VNCVisioToolsApplication.Presentation.ModelWrappers
{
    public class ControlsRowWrapper : ModelWrapper<VNCVisioAddIn.Domain.ControlsRow>
    {
        public ControlsRowWrapper() { }

        public ControlsRowWrapper(VNCVisioAddIn.Domain.ControlsRow model) : base(model)
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
