using Microsoft.Office.Interop.Excel;

namespace VNC.ExcelHelper.Domain
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
}
