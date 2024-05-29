using System.Collections.Generic;

using VNC.Core.Mvvm;

using SupportTools_Visio.Domain;

namespace SupportTools_Visio.Presentation.ModelWrappers
{
    public class DuplicatePageWrapper : ModelWrapper<Domain.DuplicatePage>
    {
        public DuplicatePageWrapper() { }
        public DuplicatePageWrapper(Domain.DuplicatePage model) : base(model)
        {
        }

        // TODO(crhodes)
        // Wrap each property from the passed in model.

        public string StringProperty { get { return GetValue<string>(); } set { SetValue(value); } }
        public int IntProperty { get { return GetValue<int>(); } set { SetValue(value); } }
    }
}
