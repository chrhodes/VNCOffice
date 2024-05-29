using System.Collections.Generic;
using System.Threading.Tasks;

using VNC.Core.DomainServices;

namespace Explore.DomainServices
{
    public interface IDoorLookupDataService
    {
        Task<IEnumerable<LookupItem>> GetDoorLookupAsync();
    }
}
