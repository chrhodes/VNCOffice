using SupportTools_Visio.Domain;
using VNC.Core.Mvvm;

using VNCVisioAddIn = VNC.Visio.VSTOAddIn;

namespace SupportTools_Visio.Presentation.ModelWrappers
{
    public class ThreeDRotationPropertiesWrapper : ModelWrapper<VNCVisioAddIn.Domain.ThreeDRotationPropertiesRow>
    {
        public ThreeDRotationPropertiesWrapper()
        {
        }
        public ThreeDRotationPropertiesWrapper(VNCVisioAddIn.Domain.ThreeDRotationPropertiesRow model) : base(model)
        {
        }

        public string RotationXAngle { get { return GetValue<string>(); } set { SetValue(value); } }
        public string RotationYAngle { get { return GetValue<string>(); } set { SetValue(value); } }
        public string RotationZAngle { get { return GetValue<string>(); } set { SetValue(value); } }
        public string RotationType { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Perspective { get { return GetValue<string>(); } set { SetValue(value); } }
        public string DistanceFromGround { get { return GetValue<string>(); } set { SetValue(value); } }
        public string KeepTextFlat { get { return GetValue<string>(); } set { SetValue(value); } }
    }
}
