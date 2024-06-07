using SupportTools_Visio.Domain;
using VNC.Core.Mvvm;

using VNCVisioAddIn = VNC.Visio.VSTOAddIn;

namespace SupportTools_Visio.Presentation.ModelWrappers
{
    public class FillFormatWrapper : ModelWrapper<VNCVisioAddIn.Domain.FillFormat>
    {
        public FillFormatWrapper()
        {
        }
        public FillFormatWrapper(VNCVisioAddIn.Domain.FillFormat model) : base(model)
        {
        }

        public string FillForegnd { get { return GetValue<string>(); } set { SetValue(value); } }
        public string ShdwForegnd { get { return GetValue<string>(); } set { SetValue(value); } }
        public string ShapeShdwType { get { return GetValue<string>(); } set { SetValue(value); } }
        public string FillForegndTrans { get { return GetValue<string>(); } set { SetValue(value); } }
        public string ShdwForegndTrans { get { return GetValue<string>(); } set { SetValue(value); } }
        public string ShapeShdwObliqueAngle { get { return GetValue<string>(); } set { SetValue(value); } }
        public string FillBkgnd { get { return GetValue<string>(); } set { SetValue(value); } }
        public string ShdwPattern { get { return GetValue<string>(); } set { SetValue(value); } }
        public string ShapeShdwScaleFactor { get { return GetValue<string>(); } set { SetValue(value); } }
        public string FillBkgndTrans { get { return GetValue<string>(); } set { SetValue(value); } }
        public string ShapeShdwOffsetX { get { return GetValue<string>(); } set { SetValue(value); } }
        public string ShapeShdwOffsetY { get { return GetValue<string>(); } set { SetValue(value); } }
        public string ShapeShdwBlur { get { return GetValue<string>(); } set { SetValue(value); } }
        public string FillPattern { get { return GetValue<string>(); } set { SetValue(value); } }
        public string ShapeShdwShow { get { return GetValue<string>(); } set { SetValue(value); } }
    }
}
