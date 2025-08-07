using System.Collections.Generic;

using VNC.Core.Mvvm;

using VNCVisioToolsApplication.Domain;

namespace VNCVisioToolsApplication.Presentation.ModelWrappers
{
    public class DuplicatePageWrapper : ModelWrapper<Domain.DuplicatePage>
    {
        public DuplicatePageWrapper() { }
        public DuplicatePageWrapper(DuplicatePage model) : base(model)
        {
        }

        // TODO(crhodes)
        // Wrap each property from the passed in model.

        public string StringProperty { get { return GetValue<string>(); } set { SetValue(value); } }
        public int IntProperty { get { return GetValue<int>(); } set { SetValue(value); } }
    }
}
