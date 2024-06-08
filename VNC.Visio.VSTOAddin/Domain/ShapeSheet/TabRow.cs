using System;

using Microsoft.Office.Interop.Visio;

namespace VNC.Visio.VSTOAddIn.Domain
{
    public class TabRow
    {
        // TODO(crhodes)
        // This looks tricky as there are an unknown number of tabs
        public string Name { get; set; }
        public string Position1 { get; set; }
        public string Alignment1 { get; set; }
        public string Position2 { get; set; }
        public string Alignment2 { get; set; }

        public static TabRow GetRow(Shape shape)
        {
            throw new NotImplementedException();
        }
    }
}
