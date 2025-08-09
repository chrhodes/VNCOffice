using VNC.Core.Mvvm;

namespace VNC.Visio.VSTOAddIn.Presentation.ModelWrappers
{
    public class UserDefinedCellRowWrapper : ModelWrapper<Domain.UserDefinedCellRow>
    {
        public UserDefinedCellRowWrapper()
        {
        }

        public UserDefinedCellRowWrapper(Domain.UserDefinedCellRow model) : base(model)
        {
        }

        public string Name { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Value { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Prompt { get { return GetValue<string>(); } set { SetValue(value); } }
    }
}
