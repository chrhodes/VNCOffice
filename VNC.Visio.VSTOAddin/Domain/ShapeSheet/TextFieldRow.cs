using System;

using Microsoft.Office.Interop.Visio;

namespace VNC.Visio.VSTOAddIn.Domain
{
    public class TextFieldRow
    {
        public string Name { get; set; }
        public string Format { get; set; }
        public string Value { get; set; }
        public string Calendar { get; set; }
        public string ObjectKind { get; set; }

        public static TextFieldRow GetRow(Shape shape)
        {
            throw new NotImplementedException();
        }
    }
}
