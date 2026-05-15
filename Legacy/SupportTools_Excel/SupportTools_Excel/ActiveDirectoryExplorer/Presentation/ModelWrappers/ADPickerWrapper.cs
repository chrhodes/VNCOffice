using System.Collections.Generic;

using VNC.Core.Mvvm;

using SupportTools_Excel.ActiveDirectoryExplorer.Domain;

namespace SupportTools_Excel.ActiveDirectoryExplorer.Presentation.ModelWrappers
{
    public class ADPickerWrapper : ModelWrapper<ADPicker>
    {
        public ADPickerWrapper() { }
        public ADPickerWrapper(Domain.ADPicker model) : base(model)
        {
        }

        // TODO(crhodes)
        // Wrap each property from the passed in model.

        public string StringProperty { get { return GetValue<string>(); } set { SetValue(value); } }
        public int IntProperty { get { return GetValue<int>(); } set { SetValue(value); } }
    }
}
