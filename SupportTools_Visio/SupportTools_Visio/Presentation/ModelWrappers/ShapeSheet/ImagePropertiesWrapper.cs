using SupportTools_Visio.Domain;
using VNC.Core.Mvvm;

namespace SupportTools_Visio.Presentation.ModelWrappers
{
    public class ImagePropertiesWrapper : ModelWrapper<Domain.ImageProperties>
    {
        public ImagePropertiesWrapper()
        {
        }
        public ImagePropertiesWrapper(ImageProperties model) : base(model)
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
