using VNC.Core.Mvvm;

namespace VNC.Visio.VSTOAddIn.Presentation.ModelWrappers
{
    public class ShapeDataRowWrapper : ModelWrapper<Domain.ShapeDataRow>
    {
        public ShapeDataRowWrapper() { }

        public ShapeDataRowWrapper(Domain.ShapeDataRow model) : base(model)
        {
        }

        public string Name { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Label { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Prompt { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Type { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Format { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Value { get { return GetValue<string>(); } set { SetValue(value); } }
        public string SortKey { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Invisible { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Ask { get { return GetValue<string>(); } set { SetValue(value); } }
        public string LangID { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Calendar { get { return GetValue<string>(); } set { SetValue(value); } }
    }
}