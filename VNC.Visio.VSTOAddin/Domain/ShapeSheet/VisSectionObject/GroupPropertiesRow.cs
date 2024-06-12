using System;
using System.Windows;


using Microsoft.Office.Interop.Visio;

namespace VNC.Visio.VSTOAddIn.Domain
{
    public class GroupPropertiesRow
    {
        public string SelectMode { get; set; }
        public string IsTextEditTarget { get; set; }
        public string IsDropTarget { get; set; }
        public string DisplayMode { get; set; }
        public string IsSnapTarget { get; set; }
        public string DontMoveChildren { get; set; }

        public static GroupPropertiesRow GetRow(Shape shape)
        {
            GroupPropertiesRow row = new GroupPropertiesRow();

            Section section = shape.Section[(short)VisSectionIndices.visSectionObject];

            if (Helpers.RowExists(shape, VisSectionIndices.visSectionObject, VisRowIndices.visRowGroup))
            {
                Row sectionRow = section[(short)VisRowIndices.visRowGroup];

                row.SelectMode = sectionRow[VisCellIndices.visGroupSelectMode].FormulaU;
                row.IsTextEditTarget = sectionRow[VisCellIndices.visGroupIsTextEditTarget].FormulaU;
                row.IsDropTarget = sectionRow[VisCellIndices.visGroupIsDropTarget].FormulaU;
                row.DisplayMode = sectionRow[VisCellIndices.visGroupDisplayMode].FormulaU;
                row.IsSnapTarget = sectionRow[VisCellIndices.visGroupIsSnapTarget].FormulaU;
                row.DontMoveChildren = sectionRow[VisCellIndices.visGroupDontMoveChildren].FormulaU;
            }
            else
            {
                MessageBox.Show("No visRowGroup exists");
            }

            return row;
        }

        public static void SetRow(Shape shape, GroupPropertiesRow groupProperties)
        {
            Section section = shape.Section[(short)VisSectionIndices.visSectionObject];

            if (Helpers.RowExists(shape, VisSectionIndices.visSectionObject, VisRowIndices.visRowGroup))
            {
                Row sectionRow = section[(short)VisRowIndices.visRowGroup];

                groupProperties.SelectMode = sectionRow[VisCellIndices.visGroupSelectMode].FormulaU;
                groupProperties.IsTextEditTarget = sectionRow[VisCellIndices.visGroupIsTextEditTarget].FormulaU;
                groupProperties.IsDropTarget = sectionRow[VisCellIndices.visGroupIsDropTarget].FormulaU;
                groupProperties.DisplayMode = sectionRow[VisCellIndices.visGroupDisplayMode].FormulaU;
                groupProperties.IsSnapTarget = sectionRow[VisCellIndices.visGroupIsSnapTarget].FormulaU;
                groupProperties.DontMoveChildren = sectionRow[VisCellIndices.visGroupDontMoveChildren].FormulaU;
            }
            else
            {
                MessageBox.Show("No visRowGroup exists");
            }
        }
    }
}
