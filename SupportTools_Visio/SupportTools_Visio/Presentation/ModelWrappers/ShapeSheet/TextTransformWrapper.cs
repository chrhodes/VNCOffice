using SupportTools_Visio.Domain;
using VNC.Core.Mvvm;

using VNCVisioAddIn = VNC.Visio.VSTOAddIn;

namespace SupportTools_Visio.Presentation.ModelWrappers
{
    public class TextTransformWrapper : ModelWrapper<VNCVisioAddIn.Domain.TextTransformRow>
    {
        public TextTransformWrapper()
        {
        }
        public TextTransformWrapper(VNCVisioAddIn.Domain.TextTransformRow model) : base(model)
        {
        }

        public string TxtWidth { get { return GetValue<string>(); } set { SetValue(value); } }
        public string TxtHeight { get { return GetValue<string>(); } set { SetValue(value); } }
        public string TxtAngle { get { return GetValue<string>(); } set { SetValue(value); } }
        public string TxtPinX { get { return GetValue<string>(); } set { SetValue(value); } }
        public string TxtPinY { get { return GetValue<string>(); } set { SetValue(value); } }
        public string TxtLocPinX { get { return GetValue<string>(); } set { SetValue(value); } }
        public string TxtLocPinY { get { return GetValue<string>(); } set { SetValue(value); } }
    }
}
