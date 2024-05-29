using VNC.Core.Mvvm;

namespace SupportTools_Visio.Presentation.ModelWrappers
{
    public class CharacterRowWrapper : ModelWrapper<Domain.CharacterRow>
    {
        public CharacterRowWrapper(Domain.CharacterRow model) : base(model)
        {
        }

        public string Font { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Size { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Scale { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Spacing { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Color { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Transparency { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Style { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Case { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Position { get { return GetValue<string>(); } set { SetValue(value); } }
        public string StrikeThru { get { return GetValue<string>(); } set { SetValue(value); } }
        public string DoubleULine { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Overline { get { return GetValue<string>(); } set { SetValue(value); } }
        public string DoubleStrikeThrough { get { return GetValue<string>(); } set { SetValue(value); } }
        public string AsianFont { get { return GetValue<string>(); } set { SetValue(value); } }
        public string ComplexScriptFont { get { return GetValue<string>(); } set { SetValue(value); } }
        public string ComplexScriptSize { get { return GetValue<string>(); } set { SetValue(value); } }
        public string LangID { get { return GetValue<string>(); } set { SetValue(value); } }
    }
}
