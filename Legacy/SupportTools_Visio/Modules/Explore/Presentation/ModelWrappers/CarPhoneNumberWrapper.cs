using Explore.Domain;

using VNC.Core.Mvvm;

namespace Explore.Presentation.ModelWrappers
{
    public class CarPhoneNumberWrapper : ModelWrapper<CarPhoneNumber>
    {
        public CarPhoneNumberWrapper(CarPhoneNumber model) : base(model)
        {
        }

        public string Number
        {
            get { return GetValue<string>(); }
            set { SetValue(value); }
        }

    }
}
