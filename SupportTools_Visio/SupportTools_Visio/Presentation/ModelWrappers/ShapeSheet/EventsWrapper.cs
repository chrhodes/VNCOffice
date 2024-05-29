using VNC.Core.Mvvm;

namespace SupportTools_Visio.Presentation.ModelWrappers
{
    public class EventsWrapper : ModelWrapper<Domain.Events>
    {
        public EventsWrapper()
        {
        }
        public EventsWrapper(Domain.Events model) : base(model)
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
