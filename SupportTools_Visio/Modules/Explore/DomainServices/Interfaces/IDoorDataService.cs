using System.Threading.Tasks;

using Explore.Domain;

using VNC.Core.DomainServices;

namespace Explore.DomainServices
{
    public interface IDoorDataService : IGenericRepository<Door>
    {
        Task<bool> IsReferencedByCarAsync(int id);
    }
}
