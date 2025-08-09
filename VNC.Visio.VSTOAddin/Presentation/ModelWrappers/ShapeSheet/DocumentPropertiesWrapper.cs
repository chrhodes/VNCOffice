using VNC.Core.Mvvm;

namespace VNC.Visio.VSTOAddIn.Presentation.ModelWrappers
{
    public class DocumentPropertiesWrapper : ModelWrapper<Domain.DocumentPropertiesRow>
    {
        public DocumentPropertiesWrapper()
        {
        }
        public DocumentPropertiesWrapper(Domain.DocumentPropertiesRow model) : base(model)
        {
        }
        public string PreviewQuality { get { return GetValue<string>(); } set { SetValue(value); } }
        public string OutputFormat { get { return GetValue<string>(); } set { SetValue(value); } }
        public string PreviewScope { get { return GetValue<string>(); } set { SetValue(value); } }
        public string LockPreview { get { return GetValue<string>(); } set { SetValue(value); } }
        public string AddMarkup { get { return GetValue<string>(); } set { SetValue(value); } }
        public string ViewMarkup { get { return GetValue<string>(); } set { SetValue(value); } }
        public string DocLangID { get { return GetValue<string>(); } set { SetValue(value); } }
        public string DocLockReplace { get { return GetValue<string>(); } set { SetValue(value); } }
        public string NoCoauth { get { return GetValue<string>(); } set { SetValue(value); } }
        public string DocLockDuplicatePage { get { return GetValue<string>(); } set { SetValue(value); } }
    }
}
