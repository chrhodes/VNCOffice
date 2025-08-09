using VNC.Core.Mvvm;

namespace VNC.Visio.VSTOAddIn.Presentation.ModelWrappers
{
    public class ShapeLayoutWrapper : ModelWrapper<Domain.ShapeLayoutRow>
    {
        public ShapeLayoutWrapper()
        {
        }
        public ShapeLayoutWrapper(Domain.ShapeLayoutRow model) : base(model)
        {
        }

        public string ShapePermeableX { get { return GetValue<string>(); } set { SetValue(value); } }
        public string ShapePermeableY { get { return GetValue<string>(); } set { SetValue(value); } }
        public string ShapeFixedCode { get { return GetValue<string>(); } set { SetValue(value); } }
        public string ConLineJumpDirX { get { return GetValue<string>(); } set { SetValue(value); } } 
        public string ConLineJumpDirY { get { return GetValue<string>(); } set { SetValue(value); } }
        public string ConLineJumpCode { get { return GetValue<string>(); } set { SetValue(value); } }
        public string ShapePlaceFlip { get { return GetValue<string>(); } set { SetValue(value); } }
        public string ShapePlaceStyle { get { return GetValue<string>(); } set { SetValue(value); } }
        public string ShapePlowCode { get { return GetValue<string>(); } set { SetValue(value); } }
        public string ConLineJumpStyle { get { return GetValue<string>(); } set { SetValue(value); } }
        public string ConLineRouteExt { get { return GetValue<string>(); } set { SetValue(value); } }
        public string DisplayLevel { get { return GetValue<string>(); } set { SetValue(value); } }
        public string ShapeRouteStyle { get { return GetValue<string>(); } set { SetValue(value); } }
        public string ConFixedCode { get { return GetValue<string>(); } set { SetValue(value); } }
        public string ShapeSplit { get { return GetValue<string>(); } set { SetValue(value); } }
        public string ShapeSplittable { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Relationships { get { return GetValue<string>(); } set { SetValue(value); } }
    }
}
