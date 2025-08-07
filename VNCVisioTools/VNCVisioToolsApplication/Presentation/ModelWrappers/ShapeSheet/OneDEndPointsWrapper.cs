using VNCVisioToolsApplication.Domain;
using VNC.Core.Mvvm;

using VNCVisioAddIn = VNC.Visio.VSTOAddIn;

namespace VNCVisioToolsApplication.Presentation.ModelWrappers
{
    public class OneDEndPointsWrapper : ModelWrapper<VNCVisioAddIn.Domain.OneDEndPointsRow>
    {
        public OneDEndPointsWrapper()
        {
        }
        public OneDEndPointsWrapper(VNCVisioAddIn.Domain.OneDEndPointsRow model) : base(model)
        {
        }

        public string BeginX { get { return GetValue<string>(); } set { SetValue(value); } }
        public string BeginY { get { return GetValue<string>(); } set { SetValue(value); } }
        public string EndX { get { return GetValue<string>(); } set { SetValue(value); } }
        public string EndY { get { return GetValue<string>(); } set { SetValue(value); } }
    }
}
