using SupportTools_Visio.Domain;
using VNC.Core.Mvvm;

using VNCVisioAddIn = VNC.Visio.VSTOAddIn;

namespace SupportTools_Visio.Presentation.ModelWrappers
{
    public class OneDEndPointsWrapper : ModelWrapper<VNCVisioAddIn.Domain.OneDEndPoints>
    {
        public OneDEndPointsWrapper()
        {
        }
        public OneDEndPointsWrapper(VNCVisioAddIn.Domain.OneDEndPoints model) : base(model)
        {
        }

        public string BeginX { get { return GetValue<string>(); } set { SetValue(value); } }
        public string BeginY { get { return GetValue<string>(); } set { SetValue(value); } }
        public string EndX { get { return GetValue<string>(); } set { SetValue(value); } }
        public string EndY { get { return GetValue<string>(); } set { SetValue(value); } }
    }
}
