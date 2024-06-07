using System.Xml.Linq;

namespace VNC.Visio.VSTOAddIn.Domain
{
    public class ShapeDataRow
    {
        public string Name { get; set; }
        public string Label { get; set; }
        public string Prompt { get; set; }
        public string Type { get; set; }
        public string Format { get; set; }
        public string Value { get; set; }
        public string SortKey { get; set; }
        public string Invisible { get; set; }
        public string Ask { get; set; }
        public string LangID { get; set; }
        public string Calendar { get; set; }

        public XElement ToXElement()
        {
            var shapeDataRow = new XElement("AddPropRow");

            shapeDataRow.SetAttributeValue("Row", Name);
            shapeDataRow.SetAttributeValue("Label", Label);
            shapeDataRow.SetAttributeValue("Prompt", Prompt);
            shapeDataRow.SetAttributeValue("Type", Type);
            shapeDataRow.SetAttributeValue("Format", Format);
            shapeDataRow.SetAttributeValue("Value", Value);
            shapeDataRow.SetAttributeValue("SortKey", SortKey);
            shapeDataRow.SetAttributeValue("Invisible", Invisible);
            shapeDataRow.SetAttributeValue("Ask", Ask);
            shapeDataRow.SetAttributeValue("LangID", LangID);
            shapeDataRow.SetAttributeValue("Calendar", Calendar);

            //shapeDataRow.SetAttributeValue("Calendar", Calendar);
            //shapeDataRow.SetAttributeValue("LangID", LangID);
            //shapeDataRow.SetAttributeValue("Ask", Ask);
            //shapeDataRow.SetAttributeValue("Invisible", Invisible);
            //shapeDataRow.SetAttributeValue("SortKey", SortKey);
            //shapeDataRow.SetAttributeValue("Value", Value);
            //shapeDataRow.SetAttributeValue("Format", Format);
            //shapeDataRow.SetAttributeValue("Type", Type);
            //shapeDataRow.SetAttributeValue("Prompt", Prompt);
            //shapeDataRow.SetAttributeValue("Label", Label);
            //shapeDataRow.SetAttributeValue("Row", Name);

            return shapeDataRow;
        }
    }
}
