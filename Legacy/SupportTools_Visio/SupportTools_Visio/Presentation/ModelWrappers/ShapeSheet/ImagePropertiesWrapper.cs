using SupportTools_Visio.Domain;
using VNC.Core.Mvvm;

using VNCVisioAddIn = VNC.Visio.VSTOAddIn;

namespace SupportTools_Visio.Presentation.ModelWrappers
{
    public class ImagePropertiesWrapper : ModelWrapper<VNCVisioAddIn.Domain.ImagePropertiesRow>
    {
        public ImagePropertiesWrapper()
        {
        }
        public ImagePropertiesWrapper(VNCVisioAddIn.Domain.ImagePropertiesRow model) : base(model)
        {
        }

        public string Contrast { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Gamma { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Sharpen { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Brightness { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Blur { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Denoise { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Transparency { get { return GetValue<string>(); } set { SetValue(value); } }
    }
}
