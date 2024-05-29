using SupportTools_Visio.Domain;
using VNC.Core.Mvvm;

namespace SupportTools_Visio.Presentation.ModelWrappers
{
    public class ThemePropertiesWrapper : ModelWrapper<Domain.ThemeProperties>
    {
        public ThemePropertiesWrapper()
        {
        }
        public ThemePropertiesWrapper(ThemeProperties model) : base(model)
        {
        }

        public string ConnectorSchemeIndex { get { return GetValue<string>(); } set { SetValue(value); } }
        public string EffectSchemeIndex { get { return GetValue<string>(); } set { SetValue(value); } }
        public string ColorSchemeIndex { get { return GetValue<string>(); } set { SetValue(value); } }
        public string FontSchemeIndex { get { return GetValue<string>(); } set { SetValue(value); } }
        public string ThemeIndex { get { return GetValue<string>(); } set { SetValue(value); } }
        public string VariationColorIndex { get { return GetValue<string>(); } set { SetValue(value); } }
        public string VariationStyleIndex { get { return GetValue<string>(); } set { SetValue(value); } }
        public string EmbellishmentIndex { get { return GetValue<string>(); } set { SetValue(value); } }
    }
}
