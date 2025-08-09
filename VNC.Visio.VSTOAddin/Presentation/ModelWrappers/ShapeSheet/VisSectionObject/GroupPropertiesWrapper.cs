using VNC.Core.Mvvm;

namespace VNC.Visio.VSTOAddIn.Presentation.ModelWrappers
{
    public class GroupPropertiesWrapper : ModelWrapper<Domain.GroupPropertiesRow>
    {
        public GroupPropertiesWrapper()
        {
        }
        public GroupPropertiesWrapper(Domain.GroupPropertiesRow model) : base(model)
        {
        }

        public string SelectMode { get { return GetValue<string>(); } set { SetValue(value); } }
        public string IsTextEditTarget { get { return GetValue<string>(); } set { SetValue(value); } }
        public string IsDropTarget { get { return GetValue<string>(); } set { SetValue(value); } }
        public string DisplayMode { get { return GetValue<string>(); } set { SetValue(value); } }
        public string IsSnapTarget { get { return GetValue<string>(); } set { SetValue(value); } }
        public string DontMoveChildren { get { return GetValue<string>(); } set { SetValue(value); } }
    }
}
