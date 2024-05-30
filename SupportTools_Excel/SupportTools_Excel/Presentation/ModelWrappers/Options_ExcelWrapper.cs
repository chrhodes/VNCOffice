using VNC.Core.Mvvm;
using SupportTools_Excel.Domain;

namespace SupportTools_Excel.Presentation.ModelWrappers
{
    public class Options_ExcelWrapper : ModelWrapper<Domain.Options_Excel>
    {
        public Options_ExcelWrapper() { }
        public Options_ExcelWrapper(Domain.Options_Excel model) : base(model)
        {
        }

        public int StartingRow { get { return GetValue<int>(); } set { SetValue(value); } }
        public int StartingColumn { get { return GetValue<int>(); } set { SetValue(value); } }
        public bool OrientOutputVertically { get { return GetValue<bool>(); } set { SetValue(value); } }
    }
}
