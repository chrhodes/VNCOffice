using Microsoft.Office.Interop.Excel;

namespace VNC.AddinHelper
{
    public class CellFormatSpecification
    {
        public string Name;
        public XlHAlign HorizontalAlignment = XlHAlign.xlHAlignGeneral;
        public XlVAlign VerticalAlignment = XlVAlign.xlVAlignBottom;
        public XlOrientation Orientation = XlOrientation.xlHorizontal;
        public bool WrapText = false;
        public bool ShrinkToFit = false;
        public FontX Font = new FontX();
        public string NumberFormat;

        public CellFormatSpecification(string name)
        {
            Name = name;
        }

        //public CellFormatSpecification(int fontSize)
        //{
        //    Font.Size = fontSize;
        //}

        //public CellFormatSpecification(int fontSize, XlHAlign horizontalAlignment) : this()
        //{
        //    Font.Size = fontSize;
        //    HorizontalAlignment = horizontalAlignment;
        //}

        //public CellFormatSpecification(Excel.MakeBold makeBold, int fontSize) : this()
        //{
        //    Font.Bold = makeBold;
        //    Font.Size = fontSize;
        //}

        //public CellFormatSpecification(Excel.MakeBold makeBold, int fontSize, XlOrientation orientation) : this()
        //{
        //    Font.Bold = makeBold;
        //    Font.Size = fontSize;
        //    Orientation = orientation;
        //}
    }

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
