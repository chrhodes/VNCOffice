using System.Data.Entity;

using Explore.Domain;

namespace Explore.Persistence.Data
{
    public interface IExploreDbContext
    {
        int SaveChanges();

        DbSet<Car> CarsSet { get; set; }
    }
}
