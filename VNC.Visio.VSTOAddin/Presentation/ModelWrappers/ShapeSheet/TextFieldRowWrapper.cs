using VNC.Core.Mvvm;

namespace VNC.Visio.VSTOAddIn.Presentation.ModelWrappers
{
    public class TextFieldRowWrapper : ModelWrapper<Domain.TextFieldRow>
    {
        public TextFieldRowWrapper() { }

        public TextFieldRowWrapper(Domain.TextFieldRow model) : base(model)
        {
        }

        public string Name { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Format { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Value { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Calendar { get { return GetValue<string>(); } set { SetValue(value); } }
        public string ObjectKind { get { return GetValue<string>(); } set { SetValue(value); } }
    }
}
