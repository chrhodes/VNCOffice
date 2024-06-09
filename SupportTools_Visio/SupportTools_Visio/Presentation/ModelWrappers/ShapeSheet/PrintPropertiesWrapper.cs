using SupportTools_Visio.Domain;
using VNC.Core.Mvvm;

using VNCVisioAddIn = VNC.Visio.VSTOAddIn;

namespace SupportTools_Visio.Presentation.ModelWrappers
{
    public class PrintPropertiesWrapper : ModelWrapper<VNCVisioAddIn.Domain.PrintPropertiesRow>
    {
        public PrintPropertiesWrapper()
        {
        }

        public PrintPropertiesWrapper(VNCVisioAddIn.Domain.PrintPropertiesRow model) : base(model)
        {
        }

        public string PageLeftMargin { get { return GetValue<string>(); } set { SetValue(value); } }
        public string PageTopMargin { get { return GetValue<string>(); } set { SetValue(value); } }
        public string PageRightMargin { get { return GetValue<string>(); } set { SetValue(value); } }
        public string PageBottomMargin { get { return GetValue<string>(); } set { SetValue(value); } }
        public string ScaleX { get { return GetValue<string>(); } set { SetValue(value); } }
        public string ScaleY { get { return GetValue<string>(); } set { SetValue(value); } }
        public string PagesX { get { return GetValue<string>(); } set { SetValue(value); } }
        public string PagesY { get { return GetValue<string>(); } set { SetValue(value); } }
        public string CenterX { get { return GetValue<string>(); } set { SetValue(value); } }
        public string CenterY { get { return GetValue<string>(); } set { SetValue(value); } }
        public string OnPage { get { return GetValue<string>(); } set { SetValue(value); } }
        public string PrintGrid { get { return GetValue<string>(); } set { SetValue(value); } }
        public string PrintPageOrientation { get { return GetValue<string>(); } set { SetValue(value); } }
        public string PaperKind { get { return GetValue<string>(); } set { SetValue(value); } }
        public string PaperSource { get { return GetValue<string>(); } set { SetValue(value); } }
    }
}
