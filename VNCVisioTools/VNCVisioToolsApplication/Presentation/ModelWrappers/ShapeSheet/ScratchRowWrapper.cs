using VNCVisioToolsApplication.Domain;
using VNC.Core.Mvvm;

using VNCVisioAddIn = VNC.Visio.VSTOAddIn;

namespace VNCVisioToolsApplication.Presentation.ModelWrappers
{
    public class ScratchRowWrapper : ModelWrapper<VNCVisioAddIn.Domain.ScratchRow>
    {
        public ScratchRowWrapper() { }

        public ScratchRowWrapper(VNCVisioAddIn.Domain.ScratchRow model) : base(model)
        {
        }

        public string Row { get { return GetValue<string>(); } set { SetValue(value); } }
        public string X { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Y { get { return GetValue<string>(); } set { SetValue(value); } }
        public string A { get { return GetValue<string>(); } set { SetValue(value); } }
        public string B { get { return GetValue<string>(); } set { SetValue(value); } }
        public string C { get { return GetValue<string>(); } set { SetValue(value); } }
        public string D { get { return GetValue<string>(); } set { SetValue(value); } }
    }
}
