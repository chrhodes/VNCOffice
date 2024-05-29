using SupportTools_Visio.Domain;
using VNC.Core.Mvvm;

namespace SupportTools_Visio.Presentation.ModelWrappers
{
    public class MiscellaneousWrapper : ModelWrapper<Domain.Miscellaneous>
    {
        public MiscellaneousWrapper()
        {
        }
        public MiscellaneousWrapper(Miscellaneous model) : base(model)
        {
        } 
        public string NoObjHandles { get { return GetValue<string>(); } set { SetValue(value); } }
        public string NoCtlHandle { get { return GetValue<string>(); } set { SetValue(value); } }
        public string NoAlignBox { get { return GetValue<string>(); } set { SetValue(value); } }
        public string NonPrinting { get { return GetValue<string>(); } set { SetValue(value); } }
        public string LangID { get { return GetValue<string>(); } set { SetValue(value); } }
        public string HideText { get { return GetValue<string>(); } set { SetValue(value); } }
        public string UpdateAlignBox { get { return GetValue<string>(); } set { SetValue(value); } }
        public string DynFeedback { get { return GetValue<string>(); } set { SetValue(value); } }
        public string NoLiveDynamics { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Calendar { get { return GetValue<string>(); } set { SetValue(value); } }
        public string ObjType { get { return GetValue<string>(); } set { SetValue(value); } }
        public string IsDropSource { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Comment { get { return GetValue<string>(); } set { SetValue(value); } }
        public string DropOnPageScale { get { return GetValue<string>(); } set { SetValue(value); } }
        public string LocalizeMerge { get { return GetValue<string>(); } set { SetValue(value); } }
        public string NoProofing { get { return GetValue<string>(); } set { SetValue(value); } }
    }
}
