using VNC.Core.Mvvm;

using VNCVisioAddIn = VNC.Visio.VSTOAddIn;

namespace SupportTools_Visio.Presentation.ModelWrappers
{
    public class ParagraphWrapper : ModelWrapper<VNCVisioAddIn.Domain.ParagraphRow>
    {
        public ParagraphWrapper(VNCVisioAddIn.Domain.ParagraphRow model) : base(model)
        {
        }

        public string IndFirst { get { return GetValue<string>(); } set { SetValue(value); } }
        public string IndLeft { get { return GetValue<string>(); } set { SetValue(value); } }
        public string IndRight { get { return GetValue<string>(); } set { SetValue(value); } }
        public string SpLine { get { return GetValue<string>(); } set { SetValue(value); } }
        public string SpBefore { get { return GetValue<string>(); } set { SetValue(value); } }
        public string SpAfter { get { return GetValue<string>(); } set { SetValue(value); } }
        public string HAlign { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Bullet { get { return GetValue<string>(); } set { SetValue(value); } }
        public string BulletString { get { return GetValue<string>(); } set { SetValue(value); } }
        public string BulletFont { get { return GetValue<string>(); } set { SetValue(value); } }
        public string TextPosAfterBullet { get { return GetValue<string>(); } set { SetValue(value); } }
        public string BulletSize { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Flags { get { return GetValue<string>(); } set { SetValue(value); } }
    }
}
