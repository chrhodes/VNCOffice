using SupportTools_Visio.Domain;
using VNC.Core.Mvvm;

using VNCVisioAddIn = VNC.Visio.VSTOAddIn;

namespace SupportTools_Visio.Presentation.ModelWrappers
{
    public class UserDefinedCellRowWrapper : ModelWrapper<VNCVisioAddIn.Domain.UserDefinedCellRow>
    {
        public UserDefinedCellRowWrapper()
        {
        }

        public UserDefinedCellRowWrapper(VNCVisioAddIn.Domain.UserDefinedCellRow model) : base(model)
        {
        }

        public string Name { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Value { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Prompt { get { return GetValue<string>(); } set { SetValue(value); } }
    }
}
