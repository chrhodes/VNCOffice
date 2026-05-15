using Microsoft.Office.Interop.Excel;

namespace VNC.ExcelHelper.Domain
{
    public class FontX : Font
    {
        public Application Application => throw new System.NotImplementedException();

        public XlCreator Creator => throw new System.NotImplementedException();

        public dynamic Parent => throw new System.NotImplementedException();

        public dynamic Background { get; set; }
        public dynamic Bold { get; set; }
        public dynamic Color { get; set; }
        public dynamic ColorIndex { get; set; }
        public dynamic FontStyle { get; set; }
        public dynamic Italic { get; set; }
        public dynamic Name { get; set; }
        public dynamic OutlineFont { get; set; }
        public dynamic Shadow { get; set; }
        public dynamic Size { get; set; }
        public dynamic Strikethrough { get; set; }
        public dynamic Subscript { get; set; }
        public dynamic Superscript { get; set; }
        public dynamic Underline { get; set; }
        public dynamic ThemeColor { get; set; }
        public dynamic TintAndShade { get; set; }
        public XlThemeFont ThemeFont { get; set; }
    }
}
