using System.Collections.ObjectModel;

using Microsoft.Office.Interop.Visio;

namespace VNC.Visio.VSTOAddIn.Domain
{
    public class ControlsRow
    {
        public string Name { get; set; }

        public string X { get; set; }
        public string Y { get; set; }
        public string XDynamics { get; set; }
        public string YDynamics { get; set; }
        public string XBehavior { get; set; }
        public string YBehavior { get; set; }
        public string CanGlue { get; set; }
        public string Tip { get; set; }

        public static ControlsRow GetRow(Shape shape)
        {
            ControlsRow controlRow = new ControlsRow();

            Section controlRowsSection = shape.Section[(short)VisSectionIndices.visSectionControls];
            Row firstControlRow = controlRowsSection[0];

            // TODO(crhodes)
            // Handle multiple ControlRows

            controlRow.Name = firstControlRow.Name;

            controlRow.X = firstControlRow[VisCellIndices.visCtlX].FormulaU;
            controlRow.Y = firstControlRow[VisCellIndices.visCtlY].FormulaU;
            controlRow.XDynamics = firstControlRow[VisCellIndices.visCtlXDyn].FormulaU;
            controlRow.YDynamics = firstControlRow[VisCellIndices.visCtlYDyn].FormulaU;
            controlRow.XBehavior = firstControlRow[VisCellIndices.visCtlXCon].FormulaU;
            controlRow.YBehavior = firstControlRow[VisCellIndices.visCtlYCon].FormulaU;
            controlRow.CanGlue = firstControlRow[VisCellIndices.visCtlGlue].FormulaU;
            controlRow.Tip = firstControlRow[VisCellIndices.visCtlTip].FormulaU;

            return controlRow;
        }

        public static ObservableCollection<ControlsRow> Get_ControlsRows(Shape shape)
        {
            var rows = new ObservableCollection<ControlsRow>();

            Section section = shape.Section[(short)VisSectionIndices.visSectionControls];

            var rowCount = section.Count;

            for (short i = 0; i < rowCount; i++)
            {
                ControlsRow controlsRow = new ControlsRow();

                var row = section[i];

                controlsRow.Name = row.NameU;

                controlsRow.X = row[VisCellIndices.visCtlX].FormulaU;
                controlsRow.Y = row[VisCellIndices.visCtlY].FormulaU;
                controlsRow.XDynamics = row[VisCellIndices.visCtlXDyn].FormulaU;
                controlsRow.YDynamics = row[VisCellIndices.visCtlYDyn].FormulaU;
                controlsRow.XBehavior = row[VisCellIndices.visCtlXCon].FormulaU;
                controlsRow.YBehavior = row[VisCellIndices.visCtlYCon].FormulaU;
                controlsRow.CanGlue = row[VisCellIndices.visCtlGlue].FormulaU;
                controlsRow.Tip = row[VisCellIndices.visCtlTip].FormulaU;

                rows.Add(controlsRow);
            }

            return rows;
        }
    }
}
