using SupportTools_Visio.Domain;
using VNC.Core.Mvvm;

using VNCVisioAddIn = VNC.Visio.VSTOAddIn;

namespace SupportTools_Visio.Presentation.ModelWrappers
{
    public class GroupPropertiesWrapper : ModelWrapper<VNCVisioAddIn.Domain.GroupPropertiesRow>
    {
        public GroupPropertiesWrapper()
        {
        }
        public GroupPropertiesWrapper(VNCVisioAddIn.Domain.GroupPropertiesRow model) : base(model)
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
