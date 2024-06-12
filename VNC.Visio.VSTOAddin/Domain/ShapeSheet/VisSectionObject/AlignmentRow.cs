using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

using Microsoft.Office.Interop.Visio;

namespace VNC.Visio.VSTOAddIn.Domain
{
    public class AlignmentRow
    {
        // TODO(crhodes)
        // populate

        public static AlignmentRow GetRow(Shape shape)
        {
            AlignmentRow row = new AlignmentRow();

            Section section = shape.Section[(short)VisSectionIndices.visSectionObject];

            if (Helpers.RowExists(shape, VisSectionIndices.visSectionObject, VisRowIndices.visRowAlign))
            {
                Row sectionRow = section[(short)VisRowIndices.visRowAlign];

                // TODO(crhodes)
                // Implement
            }
            else
            {
                MessageBox.Show("No visRowAlign exists");
            }

            return row;
        }

        public static void SetRow(Shape shape, AlignmentRow alignmentRow)
        {
            // TODO(crhodes)
            // Implement

        }
    }
}
