using VNC.Core.Mvvm;

namespace VNC.Visio.VSTOAddIn.Presentation.ModelWrappers
{
    public class ShapeTransformWrapper : ModelWrapper<Domain.ShapeTransformRow>
    {
        public ShapeTransformWrapper()
        {
        }
        public ShapeTransformWrapper(Domain.ShapeTransformRow model) : base(model)
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
