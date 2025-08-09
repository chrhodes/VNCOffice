using VNC.Core.Mvvm;

namespace VNC.Visio.VSTOAddIn.Presentation.ModelWrappers
{
    public class LineFormatWrapper : ModelWrapper<Domain.LineFormatRow>
    {
        public LineFormatWrapper()
        {
        }
        public LineFormatWrapper(Domain.LineFormatRow model) : base(model)
        {
        }

        public string LinePattern { get { return GetValue<string>(); } set { SetValue(value); } }
        public string LineWeight { get { return GetValue<string>(); } set { SetValue(value); } }
        public string LineColor { get { return GetValue<string>(); } set { SetValue(value); } }
        public string LineCap { get { return GetValue<string>(); } set { SetValue(value); } }
        public string BeginArrow { get { return GetValue<string>(); } set { SetValue(value); } }
        public string EndArrow { get { return GetValue<string>(); } set { SetValue(value); } }
        public string LineColorTrans { get { return GetValue<string>(); } set { SetValue(value); } }
        public string CompoundType { get { return GetValue<string>(); } set { SetValue(value); } }
        public string BeginArrowSize { get { return GetValue<string>(); } set { SetValue(value); } }
        public string EndArrowSize { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Rounding { get { return GetValue<string>(); } set { SetValue(value); } }
    }
}
