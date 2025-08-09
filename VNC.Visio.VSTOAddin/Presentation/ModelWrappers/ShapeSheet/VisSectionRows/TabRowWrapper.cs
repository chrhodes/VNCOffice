using VNC.Core.Mvvm;

namespace VNC.Visio.VSTOAddIn.Presentation.ModelWrappers
{
    public class TabRowWrapper : ModelWrapper<Domain.TabsRow>
    {
        public TabRowWrapper() { }

        public TabRowWrapper(Domain.TabsRow model) : base(model)
        {
        }

        // TODO(crhodes)
        // This looks tricky as there are an unknown number of tabs
        public string Name { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Position1 { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Alignment1 { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Position2 { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Alignment2 { get { return GetValue<string>(); } set { SetValue(value); } }
    }
}
