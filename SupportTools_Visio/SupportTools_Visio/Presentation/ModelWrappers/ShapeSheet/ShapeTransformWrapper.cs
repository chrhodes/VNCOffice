using SupportTools_Visio.Domain;
using VNC.Core.Mvvm;

namespace SupportTools_Visio.Presentation.ModelWrappers
{
    public class ShapeTransformWrapper : ModelWrapper<Domain.ShapeTransform>
    {
        public ShapeTransformWrapper()
        {
        }
        public ShapeTransformWrapper(ShapeTransform model) : base(model)
        {
        }

        public string Width { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Height { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Angle { get { return GetValue<string>(); } set { SetValue(value); } }
        public string PinX { get { return GetValue<string>(); } set { SetValue(value); } }
        public string PinY { get { return GetValue<string>(); } set { SetValue(value); } }
        public string LocPinX { get { return GetValue<string>(); } set { SetValue(value); } }
        public string LocPinY { get { return GetValue<string>(); } set { SetValue(value); } }
        public string FlipX { get { return GetValue<string>(); } set { SetValue(value); } }
        public string FlipY { get { return GetValue<string>(); } set { SetValue(value); } }
        public string ResizeMode { get { return GetValue<string>(); } set { SetValue(value); } }
    }
}
