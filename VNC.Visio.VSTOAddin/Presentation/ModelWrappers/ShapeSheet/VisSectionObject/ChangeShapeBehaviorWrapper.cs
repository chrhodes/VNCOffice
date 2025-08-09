using VNC.Core.Mvvm;

namespace VNC.Visio.VSTOAddIn.Presentation.ModelWrappers
{
    public class ChangeShapeBehaviorWrapper : ModelWrapper<Domain.ChangeShapeBehaviorRow>
    {
        public ChangeShapeBehaviorWrapper()
        {
        }
        public ChangeShapeBehaviorWrapper(Domain.ChangeShapeBehaviorRow model) : base(model)
        {
        }

        public string ReplaceLockShapeData { get { return GetValue<string>(); } set { SetValue(value); } }
        public string ReplaceLockText { get { return GetValue<string>(); } set { SetValue(value); } }
        public string ReplaceLockFormat { get { return GetValue<string>(); } set { SetValue(value); } }
        public string ReplaceCopyCells { get { return GetValue<string>(); } set { SetValue(value); } }
    }
}
