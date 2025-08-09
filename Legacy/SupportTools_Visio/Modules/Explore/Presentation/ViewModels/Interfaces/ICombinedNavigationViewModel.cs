using System.Threading.Tasks;

using VNC.Core.Mvvm;

namespace Explore.Presentation.ViewModels
{
    public interface ICombinedNavigationViewModel : IViewModel
    {
        Task LoadAsync();
    }
}
