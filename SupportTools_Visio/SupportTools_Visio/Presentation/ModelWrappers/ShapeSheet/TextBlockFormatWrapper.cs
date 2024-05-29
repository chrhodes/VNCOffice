using VNC.Core.Mvvm;

namespace SupportTools_Visio.Presentation.ModelWrappers
{
    public class TextBlockFormatWrapper : ModelWrapper<Domain.TextBlockFormat>
    {
        public TextBlockFormatWrapper()
        {
        }
        public TextBlockFormatWrapper(Domain.TextBlockFormat model) : base(model)
        {
        }

        public string LeftMargin
        { 
            get { return GetValue<string>(); }
            set { SetValue(value); }
        }
        public string TopMargin { get { return GetValue<string>(); } set { SetValue(value); } }
        public string RightMargin { get { return GetValue<string>(); } set { SetValue(value); } }
        public string BottomMargin { get { return GetValue<string>(); } set { SetValue(value); } }

        public string TextBkgnd { get { return GetValue<string>(); } set { SetValue(value); } }
        public string TextBkgndTrans { get { return GetValue<string>(); } set { SetValue(value); } }
        public string TextDirection { get { return GetValue<string>(); } set { SetValue(value); } }
        public string VerticalAlign { get { return GetValue<string>(); } set { SetValue(value); } }
        public string DefaultTabStop { get { return GetValue<string>(); } set { SetValue(value); } }
    }
}
