using VNC.Core.Mvvm;

namespace VNC.Visio.VSTOAddIn.Presentation.ModelWrappers
{
    public class EventsWrapper : ModelWrapper<Domain.EventsRow>
    {
        public EventsWrapper()
        {
        }
        public EventsWrapper(Domain.EventsRow model) : base(model)
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
