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
    }
}
