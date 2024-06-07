using SupportTools_Visio.Domain;
using VNC.Core.Mvvm;

using VNCVisioAddIn = VNC.Visio.VSTOAddIn;

namespace SupportTools_Visio.Presentation.ModelWrappers
{ 
    public class ProtectionWrapper : ModelWrapper<VNCVisioAddIn.Domain.Protection>
    {
        public ProtectionWrapper()
        {
        }
        public ProtectionWrapper(VNCVisioAddIn.Domain.Protection model) : base(model)
        {
        }

        public string LockWidth { get { return GetValue<string>(); } set { SetValue(value); } }
        public string LockHeight { get { return GetValue<string>(); } set { SetValue(value); } }
        public string LockAspect { get { return GetValue<string>(); } set { SetValue(value); } }
        public string LockMoveX { get { return GetValue<string>(); } set { SetValue(value); } }
        public string LockMoveY { get { return GetValue<string>(); } set { SetValue(value); } }
        public string LockRotate { get { return GetValue<string>(); } set { SetValue(value); } }
        public string LockBegin { get { return GetValue<string>(); } set { SetValue(value); } }
        public string LockReplace { get { return GetValue<string>(); } set { SetValue(value); } }
        public string LockEnd { get { return GetValue<string>(); } set { SetValue(value); } }
        public string LockDelete { get { return GetValue<string>(); } set { SetValue(value); } }
        public string LockSelect { get { return GetValue<string>(); } set { SetValue(value); } }
        public string LockFormat { get { return GetValue<string>(); } set { SetValue(value); } }
        public string LockCustProp { get { return GetValue<string>(); } set { SetValue(value); } }
        public string LockTextEdit { get { return GetValue<string>(); } set { SetValue(value); } }
        public string LockVtxEdit { get { return GetValue<string>(); } set { SetValue(value); } }
        public string LockThemeIndex { get { return GetValue<string>(); } set { SetValue(value); } }
        public string LockCrop { get { return GetValue<string>(); } set { SetValue(value); } }
        public string LockGroup { get { return GetValue<string>(); } set { SetValue(value); } }
        public string LockCalcWH { get { return GetValue<string>(); } set { SetValue(value); } }
        public string LockFromGroupFormat { get { return GetValue<string>(); } set { SetValue(value); } }
        public string LockThemeColors { get { return GetValue<string>(); } set { SetValue(value); } }
        public string LockThemeEffects { get { return GetValue<string>(); } set { SetValue(value); } }
        public string LockThemeConnectors { get { return GetValue<string>(); } set { SetValue(value); } }
        public string LockThemeFonts { get { return GetValue<string>(); } set { SetValue(value); } }
        public string LockVariation { get { return GetValue<string>(); } set { SetValue(value); } }
    }
}
