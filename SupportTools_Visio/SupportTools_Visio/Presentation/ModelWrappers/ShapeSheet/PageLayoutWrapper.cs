using SupportTools_Visio.Domain;
using VNC.Core.Mvvm;

using VNCVisioAddIn = VNC.Visio.VSTOAddIn;

namespace SupportTools_Visio.Presentation.ModelWrappers
{
    public class PageLayoutWrapper : ModelWrapper<VNCVisioAddIn.Domain.PageLayout>
    {
        public PageLayoutWrapper()
        {
        }
        public PageLayoutWrapper(VNCVisioAddIn.Domain.PageLayout model) : base(model)
        {
        }

        public string PlaceStyle { get { return GetValue<string>(); } set { SetValue(value); } }
        public string PlaceDepth { get { return GetValue<string>(); } set { SetValue(value); } }
        public string PlowCode { get { return GetValue<string>(); } set { SetValue(value); } }
        public string ResizePage { get { return GetValue<string>(); } set { SetValue(value); } }
        public string DynamicsOff { get { return GetValue<string>(); } set { SetValue(value); } }
        public string EnableGrid { get { return GetValue<string>(); } set { SetValue(value); } }
        public string CtrlAsInput { get { return GetValue<string>(); } set { SetValue(value); } }
        public string LineAdjustFrom { get { return GetValue<string>(); } set { SetValue(value); } }
        public string PlaceFlip { get { return GetValue<string>(); } set { SetValue(value); } }
        public string AvoidPageBreaks { get { return GetValue<string>(); } set { SetValue(value); } }
        public string BlockSizeX { get { return GetValue<string>(); } set { SetValue(value); } }
        public string BlockSizeY { get { return GetValue<string>(); } set { SetValue(value); } }
        public string AvenueSizeX { get { return GetValue<string>(); } set { SetValue(value); } }
        public string AvenueSizeY { get { return GetValue<string>(); } set { SetValue(value); } }

        public string RouteStyle { get { return GetValue<string>(); } set { SetValue(value); } }
        public string PageLineJumpDirX { get { return GetValue<string>(); } set { SetValue(value); } }
        public string PageLineJumpDirY { get { return GetValue<string>(); } set { SetValue(value); } }
        public string LineAdjustTo { get { return GetValue<string>(); } set { SetValue(value); } }
        public string LineRouteExt { get { return GetValue<string>(); } set { SetValue(value); } }
        public string LineToNodeX { get { return GetValue<string>(); } set { SetValue(value); } }
        public string LineToNodeY { get { return GetValue<string>(); } set { SetValue(value); } }
        public string LineToLineX { get { return GetValue<string>(); } set { SetValue(value); } }
        public string LineToLineY { get { return GetValue<string>(); } set { SetValue(value); } }
        public string LineJumpFactorX { get { return GetValue<string>(); } set { SetValue(value); } }
        public string LineJumpFactorY { get { return GetValue<string>(); } set { SetValue(value); } }
        public string LineJumpCode { get { return GetValue<string>(); } set { SetValue(value); } }
        public string LineJumpStyle { get { return GetValue<string>(); } set { SetValue(value); } }
        public string PageShapeSplit { get { return GetValue<string>(); } set { SetValue(value); } }
    }
}
