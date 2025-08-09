using System.Threading.Tasks;

using VNC.Core.Mvvm;

namespace Explore.Presentation.ViewModels
{
    public interface ICarNavigationViewModel : IViewModel
    {
        Task LoadAsync();
    }
}
