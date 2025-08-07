using VNC.Core.Mvvm;

using VNCVisioAddIn = VNC.Visio.VSTOAddIn;

namespace VNCVisioToolsApplication.Presentation.ModelWrappers
{
    public class ChangeShapeBehaviorWrapper : ModelWrapper<VNCVisioAddIn.Domain.ChangeShapeBehaviorRow>
    {
        public ChangeShapeBehaviorWrapper()
        {
        }
        public ChangeShapeBehaviorWrapper(VNCVisioAddIn.Domain.ChangeShapeBehaviorRow model) : base(model)
        {
        }

        public string ReplaceLockShapeData { get { return GetValue<string>(); } set { SetValue(value); } }
        public string ReplaceLockText { get { return GetValue<string>(); } set { SetValue(value); } }
        public string ReplaceLockFormat { get { return GetValue<string>(); } set { SetValue(value); } }
        public string ReplaceCopyCells { get { return GetValue<string>(); } set { SetValue(value); } }
    }
}
