﻿using SupportTools_Visio.Domain;
using VNC.Core.Mvvm;

namespace SupportTools_Visio.Presentation.ModelWrappers
{
    public class DocumentPropertiesWrapper : ModelWrapper<Domain.DocumentProperties>
    {
        public DocumentPropertiesWrapper()
        {
        }
        public DocumentPropertiesWrapper(DocumentProperties model) : base(model)
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
