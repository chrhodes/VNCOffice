using VNC.Core.Mvvm;

using VNCVisioAddIn = VNC.Visio.VSTOAddIn;

namespace VNCVisioToolsApplication.Presentation.ModelWrappers
{
    public class EventsWrapper : ModelWrapper<VNCVisioAddIn.Domain.EventsRow>
    {
        public EventsWrapper()
        {
        }
        public EventsWrapper(VNCVisioAddIn.Domain.EventsRow model) : base(model)
        {
        }

        public string TheData { get { return GetValue<string>(); } set { SetValue(value); } }
        public string EventDblClick { get { return GetValue<string>(); } set { SetValue(value); } }
        public string EventDrop { get { return GetValue<string>(); } set { SetValue(value); } }
        public string TheText { get { return GetValue<string>(); } set { SetValue(value); } }
        public string EventXFMod { get { return GetValue<string>(); } set { SetValue(value); } }
        public string EventMultiDrop { get { return GetValue<string>(); } set { SetValue(value); } }
    }
}
