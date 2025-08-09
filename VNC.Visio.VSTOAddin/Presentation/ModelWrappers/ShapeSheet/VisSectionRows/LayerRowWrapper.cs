using VNC.Core.Mvvm;

namespace VNC.Visio.VSTOAddIn.Presentation.ModelWrappers
{
    public class LayerRowWrapper : ModelWrapper<Domain.LayerRow>
    {
        public LayerRowWrapper() { }

        public LayerRowWrapper(Domain.LayerRow model) : base(model)
        {
        }

        public string Name { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Visible { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Print { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Active { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Lock { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Snap { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Glue { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Color { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Transparency { get { return GetValue<string>(); } set { SetValue(value); } }
    }
}
