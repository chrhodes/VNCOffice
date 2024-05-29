using System.Threading.Tasks;

using VNC.Core.Mvvm;

namespace Explore.Presentation.ViewModels
{
    public interface ICombinedMainViewModel : IViewModel
    {
        Task LoadAsync();
    }
}
