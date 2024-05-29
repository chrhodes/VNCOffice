using System.Threading.Tasks;

using VNC.Core.Mvvm;

namespace Explore.Presentation.ViewModels
{
    public interface ICarDetailViewModel : IViewModel
    {
        Task LoadAsync(int id);
    }
}
