using SupportTools_Visio.Domain;
using VNC.Core.Mvvm;

using VNCVisioAddIn = VNC.Visio.VSTOAddIn;

namespace SupportTools_Visio.Presentation.ModelWrappers
{
    public class AdditionalEffectPropertiesWrapper : ModelWrapper<VNCVisioAddIn.Domain.AdditionalEffectProperties>
    {
        public AdditionalEffectPropertiesWrapper()
        {
        }
        public AdditionalEffectPropertiesWrapper(VNCVisioAddIn.Domain.AdditionalEffectProperties model) : base(model)
        {
        }

        public string ReflectionTrans { get { return GetValue<string>(); } set { SetValue(value); } }
        public string ReflectionSize { get { return GetValue<string>(); } set { SetValue(value); } }
        public string ReflectionDist { get { return GetValue<string>(); } set { SetValue(value); } }
        public string ReflectionBlur { get { return GetValue<string>(); } set { SetValue(value); } }
        public string SketchEnabled { get { return GetValue<string>(); } set { SetValue(value); } }
        public string SketchSeed { get { return GetValue<string>(); } set { SetValue(value); } }
        public string SketchAmount { get { return GetValue<string>(); } set { SetValue(value); } }
        public string SketchLineWeight { get { return GetValue<string>(); } set { SetValue(value); } }
        public string SketchLineChange { get { return GetValue<string>(); } set { SetValue(value); } }
        public string SketchFillChange { get { return GetValue<string>(); } set { SetValue(value); } }
        public string GlowColor { get { return GetValue<string>(); } set { SetValue(value); } }
        public string GlowColorTrans { get { return GetValue<string>(); } set { SetValue(value); } }
        public string GlowSize { get { return GetValue<string>(); } set { SetValue(value); } }
        public string SoftEdgesSize { get { return GetValue<string>(); } set { SetValue(value); } }
    }
}
