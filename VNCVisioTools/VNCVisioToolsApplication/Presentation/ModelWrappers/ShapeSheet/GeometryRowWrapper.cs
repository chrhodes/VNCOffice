using VNCVisioToolsApplication.Domain;
using VNC.Core.Mvvm;

using VNCVisioAddIn = VNC.Visio.VSTOAddIn;

namespace VNCVisioToolsApplication.Presentation.ModelWrappers
{
    public class GeometryRowWrapper : ModelWrapper<VNCVisioAddIn.Domain.GeometryRow>
    {
        public GeometryRowWrapper(VNCVisioAddIn.Domain.GeometryRow model) : base(model)
        {
        }

        // TODO(crhodes)
        // This is going to take work
        public string Name { get { return GetValue<string>(); } set { SetValue(value); } }
        public string X { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Y { get { return GetValue<string>(); } set { SetValue(value); } }
        public string A { get { return GetValue<string>(); } set { SetValue(value); } }
        public string B { get { return GetValue<string>(); } set { SetValue(value); } }
        public string C { get { return GetValue<string>(); } set { SetValue(value); } }
        public string D { get { return GetValue<string>(); } set { SetValue(value); } }
    }
}
