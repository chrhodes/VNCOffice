using System;
using System.Windows;

using Microsoft.Office.Interop.Visio;

namespace VNC.Visio.VSTOAddIn.Domain
{
    public class LayerMembershipRow
    {
        // TODO(crhodes)
        // Not clear what this is all about in ShapeSheet

        public string Name { get; set; }

        public static LayerMembershipRow GetRow(Shape shape)
        {
            LayerMembershipRow row = new LayerMembershipRow();

            Section section = shape.Section[(short)VisSectionIndices.visSectionObject];

            if (Helpers.RowExists(shape, VisSectionIndices.visSectionObject, VisRowIndices.visRowLayerMem))
            {
                Row sectionRow = section[(short)VisRowIndices.visRowLayerMem];

                row.Name = sectionRow[VisCellIndices.visLayerMember].FormulaU;
            }
            else
            {
                MessageBox.Show("No visRowLayerMem exists");
            }

            return row;
        }

        public static void SetRow(Shape shape, LayerMembershipRow layerMembership)
        {
            try
            {
                Section section = shape.Section[(short)VisSectionIndices.visSectionObject];

                if (Helpers.RowExists(shape, VisSectionIndices.visSectionObject, VisRowIndices.visRowLayerMem))
                {
                    Row sectionRow = section[(short)VisRowIndices.visRowLayerMem];

                    sectionRow[VisCellIndices.visLayerMember].FormulaU = layerMembership.Name;
                }
                else
                {
                    MessageBox.Show("No visRowLayerMem exists");
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }
    }
}
