using VNCVisioToolsApplication.Domain;
using VNC.Core.Mvvm;

using VNCVisioAddIn = VNC.Visio.VSTOAddIn;

namespace VNCVisioToolsApplication.Presentation.ModelWrappers
{
    public class ConnectionPointRowWrapper : ModelWrapper<VNCVisioAddIn.Domain.ConnectionPointRow>
    {
        public ConnectionPointRowWrapper() { }

        public ConnectionPointRowWrapper(VNCVisioAddIn.Domain.ConnectionPointRow model) : base(model)
        {
        }

        public string Name { get { return GetValue<string>(); } set { SetValue(value); } }
        public string X { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Y { get { return GetValue<string>(); } set { SetValue(value); } }
        public string DirX { get { return GetValue<string>(); } set { SetValue(value); } }
        public string A { get { return GetValue<string>(); } set { SetValue(value); } }
        public string DirY { get { return GetValue<string>(); } set { SetValue(value); } }
        public string B { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Type { get { return GetValue<string>(); } set { SetValue(value); } }
        public string C { get { return GetValue<string>(); } set { SetValue(value); } }
        public string D { get { return GetValue<string>(); } set { SetValue(value); } }
    }
}
