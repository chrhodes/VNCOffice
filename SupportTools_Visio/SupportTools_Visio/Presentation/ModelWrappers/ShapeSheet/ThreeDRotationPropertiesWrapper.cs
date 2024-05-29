using SupportTools_Visio.Domain;
using VNC.Core.Mvvm;

namespace SupportTools_Visio.Presentation.ModelWrappers
{
    public class ThreeDRotationPropertiesWrapper : ModelWrapper<Domain.ThreeDRotationProperties>
    {
        public ThreeDRotationPropertiesWrapper()
        {
        }
        public ThreeDRotationPropertiesWrapper(ThreeDRotationProperties model) : base(model)
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
