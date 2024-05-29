using SupportTools_Visio.Domain;
using VNC.Core.Mvvm;

namespace SupportTools_Visio.Presentation.ModelWrappers
{
    public class OneDEndPointsWrapper : ModelWrapper<Domain.OneDEndPoints>
    {
        public OneDEndPointsWrapper()
        {
        }
        public OneDEndPointsWrapper(OneDEndPoints model) : base(model)
        {
        }

        public string BeginX { get { return GetValue<string>(); } set { SetValue(value); } }
        public string BeginY { get { return GetValue<string>(); } set { SetValue(value); } }
        public string EndX { get { return GetValue<string>(); } set { SetValue(value); } }
        public string EndY { get { return GetValue<string>(); } set { SetValue(value); } }
    }
}
