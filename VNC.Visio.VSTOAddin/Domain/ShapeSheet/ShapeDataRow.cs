using System.Collections.ObjectModel;
using System.Xml.Linq;

using Microsoft.Office.Interop.Visio;

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

        public static ObservableCollection<ShapeDataRow> GetRows(Shape shape)
        {
            var rows = new ObservableCollection<ShapeDataRow>();

            Section section = shape.Section[(short)VisSectionIndices.visSectionProp];

            var rowCount = section.Count;

            for (short i = 0; i < rowCount; i++)
            {
                ShapeDataRow shapeDataRow = new ShapeDataRow();

                var row = section[i];

                shapeDataRow.Name = row.NameU;

                // HACK(crhodes)
                // Trying to find a way to determine if there is a formula in the cell
                // Nothing obvious
                var fooProp = row[(short)VisCellIndices.visCustPropsPrompt];
                var fooStat = row[(short)VisCellIndices.visCustPropsPrompt].Stat;

                var fooResult = row[(short)VisCellIndices.visCustPropsPrompt].Result[VisUnitCodes.visUnitsInval];
                var fooFormula = row[(short)VisCellIndices.visXFormPinY].Formula;
                var fooFormulaU = row[(short)VisCellIndices.visCustPropsPrompt].FormulaU;
                var fooResultStrU = row[(short)VisCellIndices.visCustPropsPrompt].ResultStrU[VisUnitCodes.visUnitsString];

                var fooUnits = row[(short)VisCellIndices.visCustPropsPrompt].Units;

                //shapeDataRow.Label = row[(short)VisCellIndices.visCustPropsLabel].FormulaU;
                //shapeDataRow.Prompt = row[(short)VisCellIndices.visCustPropsPrompt].FormulaU;
                //shapeDataRow.Type = row[(short)VisCellIndices.visCustPropsType].FormulaU;
                //shapeDataRow.Format = row[(short)VisCellIndices.visCustPropsFormat].FormulaU;
                //shapeDataRow.Value = row[(short)VisCellIndices.visCustPropsValue].FormulaU;
                //shapeDataRow.SortKey = row[(short)VisCellIndices.visCustPropsSortKey].FormulaU;
                //shapeDataRow.Invisible = row[(short)VisCellIndices.visCustPropsInvis].FormulaU;
                //shapeDataRow.Ask = row[(short)VisCellIndices.visCustPropsAsk].FormulaU;
                //shapeDataRow.LangID = row[(short)VisCellIndices.visCustPropsLangID].FormulaU;
                //shapeDataRow.Calendar = row[(short)VisCellIndices.visCustPropsCalendar].FormulaU;

                shapeDataRow.Label = row[(short)VisCellIndices.visCustPropsLabel].ResultStrU[VisUnitCodes.visUnitsString];


                shapeDataRow.Prompt = row[(short)VisCellIndices.visCustPropsPrompt].ResultStrU[VisUnitCodes.visUnitsString];
                shapeDataRow.Type = row[(short)VisCellIndices.visCustPropsType].ResultStrU[VisUnitCodes.visUnitsString];
                shapeDataRow.Format = row[(short)VisCellIndices.visCustPropsFormat].ResultStrU[VisUnitCodes.visUnitsString];
                shapeDataRow.Value = row[(short)VisCellIndices.visCustPropsValue].ResultStrU[VisUnitCodes.visUnitsString];
                shapeDataRow.SortKey = row[(short)VisCellIndices.visCustPropsSortKey].ResultStrU[VisUnitCodes.visUnitsString];
                shapeDataRow.Invisible = row[(short)VisCellIndices.visCustPropsInvis].ResultStrU[VisUnitCodes.visUnitsString];
                shapeDataRow.Ask = row[(short)VisCellIndices.visCustPropsAsk].ResultStrU[VisUnitCodes.visUnitsString];
                shapeDataRow.LangID = row[(short)VisCellIndices.visCustPropsLangID].ResultStrU[VisUnitCodes.visUnitsString];
                shapeDataRow.Calendar = row[(short)VisCellIndices.visCustPropsCalendar].ResultStrU[VisUnitCodes.visUnitsString];

                rows.Add(shapeDataRow);
            }

            return rows;
        }
    }
}
