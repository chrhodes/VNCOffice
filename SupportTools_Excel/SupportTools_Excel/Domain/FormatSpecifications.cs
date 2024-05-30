using System.Drawing;

using VNC.AddinHelper;

using static VNC.AddinHelper.Excel;

namespace SupportTools_Excel.Domain
{
    public class FormatSpecifications
    {
        public CellFormatSpecification DateLabel;
        public CellFormatSpecification DateContent;
        public CellFormatSpecification RedContent;
        public CellFormatSpecification WITContent;

        public bool IsInitialized { get; set; } = false;

        public void Initialize(XlLocation insertAt)
        {
            // This did not initialize everything
            //RedContent = new CellFormatSpecification("RedContent");

            RedContent = insertAt.CreateCellFormat("RedContent", insertAt.ContentLeftFormat);
            //XlLocation.InitializeFont(insertAt.ContentLeftFormat, RedContent);
            RedContent.Font.Color = Color.Red;
            RedContent.Font.Bold = true;

            WITContent = insertAt.CreateCellFormat("WITContent", insertAt.ContentLeftFormat);
            //XlLocation.InitializeFont(insertAt.ContentLeftFormat, WITContent);
            WITContent.Font.Color = Color.Green;
            WITContent.Font.Bold = true;

            DateLabel = insertAt.CreateCellFormat("DateLabel", insertAt.LabelRightFormat);
            //XlLocation.InitializeFont(insertAt.LabelRightFormat, DateLabel);
            DateLabel.Font.Color = Color.DarkOrange;

            DateContent = insertAt.CreateCellFormat("DateContent", insertAt.ContentLeftFormat);
            //XlLocation.InitializeFont(insertAt.ContentLeftFormat, DateContent);
            DateContent.Font.Color = Color.Orange;
            DateContent.Font.Bold = true;

            IsInitialized = true;
        }
    }
}
