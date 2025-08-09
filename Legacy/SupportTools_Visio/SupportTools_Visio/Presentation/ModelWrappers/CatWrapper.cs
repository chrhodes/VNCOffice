using System.Collections.Generic;

using VNC.Core.Mvvm;

using SupportTools_Visio.Domain;

namespace SupportTools_Visio.Presentation.Presentation.ModelWrappers
{
    public class CatWrapper : ModelWrapper<Domain.Cat>
    {
        public CatWrapper() { }
        public CatWrapper(Domain.Cat model) : base(model)
        {
        }

        // TODO(crhodes)
        // Wrap each property from the passed in model.

        public string StringProperty { get { return GetValue<string>(); } set { SetValue(value); } }
        public int IntProperty { get { return GetValue<int>(); } set { SetValue(value); } }
    }
}
