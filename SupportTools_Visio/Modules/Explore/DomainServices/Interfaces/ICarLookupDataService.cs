using System.Collections.Generic;
using System.Threading.Tasks;

using VNC.Core.DomainServices;

namespace Explore.DomainServices
{
    public interface ICarLookupDataService
    {
        Task<IEnumerable<LookupItem>> GetCarLookupAsync();
    }
}
