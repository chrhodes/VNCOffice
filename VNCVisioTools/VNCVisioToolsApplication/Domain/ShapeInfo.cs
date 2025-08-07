using MSVisio = Microsoft.Office.Interop.Visio;

namespace VNCVisioToolsApplication.Actions
{
    public class ShapeInfo
    {
        #region Constructors and Load

        public ShapeInfo(MSVisio.Shape shape)
        {
            // This helps with position output relative to the activeShape

            PinX = shape.CellsU["PinX"].ResultIU;
            PinY = shape.CellsU["PinY"].ResultIU;

            Height= shape.CellsU["Height"].ResultIU;
            Width = shape.CellsU["Width"].ResultIU;
        }

        #endregion

        #region Enums, Fields, Properties, Structures

        public double PinX { get; set; }
        public double PinY { get; set; }

        public double Height { get; set; }
        public double Width { get; set; }

        #endregion

    }
}
