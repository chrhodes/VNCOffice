using System.Collections.ObjectModel;

using Microsoft.Office.Interop.Visio;

namespace VNC.Visio.VSTOAddIn.Domain
{
    public class CharacterRow
    {
        public string Font { get; set; }
        public string Size { get; set; }
        public string Scale { get; set; }
        public string Spacing { get; set; }
        public string Color { get; set; }
        public string Transparency { get; set; }
        public string Style { get; set; }
        public string Case { get; set; }
        public string Position { get; set; }
        public string StrikeThru { get; set; }
        public string DoubleULine { get; set; }
        public string Overline { get; set; }
        public string DoubleStrikeThrough { get; set; }
        public string AsianFont { get; set; }
        public string ComplexScriptFont { get; set; }
        public string ComplexScriptSize { get; set; }
        public string LangID { get; set; }

        public static CharacterRow GetRow(Shape shape)
        {
            CharacterRow row = new CharacterRow();

            Section section = shape.Section[(short)VisSectionIndices.visSectionCharacter];
            Row sectionRow = section[0];

            // TODO(crhodes)
            // Handle multiple rows

            row.AsianFont = sectionRow[VisCellIndices.visCharacterAsianFont].FormulaU;

            return row;
        }

        public static ObservableCollection<CharacterRow> GetRows(Shape shape)
        {
            var rows = new ObservableCollection<CharacterRow>();

            Section section = shape.Section[(short)VisSectionIndices.visSectionCharacter];

            var rowCount = section.Count;

            for (short i = 0; i < rowCount; i++)
            {
                CharacterRow characterRow = new CharacterRow();

                var row = section[i];

                //layerRow.Name = row[(short)VisCellIndices.visLayerName].FormulaU;
                //layerRow.Visible = row[(short)VisCellIndices.visLayerVisible].FormulaU;
                //layerRow.Print = row[(short)VisCellIndices.visLayerPrint].FormulaU;
                //layerRow.Active = row[(short)VisCellIndices.visLayerActive].FormulaU;
                //layerRow.Lock = row[(short)VisCellIndices.visLayerLock].FormulaU;
                //layerRow.Snap = row[(short)VisCellIndices.visLayerSnap].FormulaU;
                //layerRow.Glue = row[(short)VisCellIndices.visLayerGlue].FormulaU;
                //layerRow.Color = row[(short)VisCellIndices.visLayerColor].FormulaU;
                //layerRow.Transparency = row[(short)VisCellIndices.visLayerColorTrans].FormulaU;

                //// NOTE(crhodes)
                //// There are a few more VisCellIndices.  See what they do
                ////VisCellIndices.visLayerMember
                ////VisCellIndices.visLayerStatus

                rows.Add(characterRow);
            }

            return rows;
        }
    }
}
