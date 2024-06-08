using System;

using Microsoft.Office.Interop.Visio;

namespace VNC.Visio.VSTOAddIn.Domain
{
    public class LayerMembership
    {
        // TODO(crhodes)
        // Not clear what this is all about in ShapeSheet

        public string Name { get; set; }

        public static LayerMembership GetRow(Shape shape)
        {
            LayerMembership row = new LayerMembership();

            Section section = shape.Section[(short)VisSectionIndices.visSectionObject];
            Row sectionRow = section[(short)VisRowIndices.visRowLayerMem];

            row.Name = sectionRow[VisCellIndices.visLayerMember].FormulaU;

            return row;
        }

        public static void Set_LayerMembership_Section(Shape shape, LayerMembership layerMembership)
        {
            try
            {
                Section section = shape.Section[(short)VisSectionIndices.visSectionObject];
                Row sectionRow = section[(short)VisRowIndices.visRowLayerMem];

                sectionRow[VisCellIndices.visLayerMember].FormulaU = layerMembership.Name;
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }
    }
}
