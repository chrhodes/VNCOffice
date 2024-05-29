using SupportTools_Visio.Domain;
using VNC.Core.Mvvm;

namespace SupportTools_Visio.Presentation.ModelWrappers
{
    public class GroupPropertiesWrapper : ModelWrapper<Domain.GroupProperties>
    {
        public GroupPropertiesWrapper()
        {
        }
        public GroupPropertiesWrapper(GroupProperties model) : base(model)
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
