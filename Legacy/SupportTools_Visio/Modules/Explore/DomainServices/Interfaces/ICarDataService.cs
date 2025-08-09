using Explore.Domain;

using VNC.Core.DomainServices;

namespace Explore.DomainServices
{
    public interface ICarDataService : IGenericRepository<Car>
    {
        void RemovePhoneNumber(CarPhoneNumber model);
    }
}
