namespace VNC.Visio.VSTOAddIn.Domain
{
    //public class TextBlockFormat
    //{
    //    public string LeftMargin = "Char.Size/2";
    //    public string RightMargin = "Char.Size/2";
    //    public string TextBkgnd = "0";
    //    public string TopMargin = "Char.Size/4";
    //    public string BottomMargin = "Char.Size/4";
    //    public string TextBkgndTrans = "0%";
    //    public string TextDirection = "0";
    //    public string VerticalAlign = "1";
    //    public string DefaultTabStop = "0.5 in";
    //}

    public class TextBlockFormat
    {
        public TextBlockFormat()
        {
            // Initialize Defaults
            //LeftMargin = "Char.Size/2";
            //RightMargin = "Char.Size/2";
            //TopMargin = "Char.Size/4";
            //BottomMargin = "Char.Size/4";
        }

        public string LeftMargin { get; set; } = "Char.Size/2";
        public string TopMargin { get; set; } = "Char.Size/2";
        public string RightMargin { get; set; } = "Char.Size/4";
        public string BottomMargin { get; set; } = "Char.Size/4";

        public string TextBkgnd { get; set; } = "0";
        public string TextBkgndTrans { get; set; } = "0%";
        public string TextDirection { get; set; } = "0";
        public string VerticalAlign { get; set; } = "1";
        public string DefaultTabStop { get; set; } = "0.5 in";
    }
}
