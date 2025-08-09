using VNC.Core.Mvvm;

namespace VNC.Visio.VSTOAddIn.Presentation.ModelWrappers
{
    public class OneDEndPointsWrapper : ModelWrapper<Domain.OneDEndPointsRow>
    {
        public OneDEndPointsWrapper()
        {
        }
        public OneDEndPointsWrapper(Domain.OneDEndPointsRow model) : base(model)
        {
        }

        public string BeginX { get { return GetValue<string>(); } set { SetValue(value); } }
        public string BeginY { get { return GetValue<string>(); } set { SetValue(value); } }
        public string EndX { get { return GetValue<string>(); } set { SetValue(value); } }
        public string EndY { get { return GetValue<string>(); } set { SetValue(value); } }
    }
}
