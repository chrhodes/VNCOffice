using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Threading.Tasks;

using Explore.Persistence.Data;

using VNC;

using VNC.Core.DomainServices;

namespace Explore.DomainServices
{
    public class CarLookupDataService : ICarLookupDataService
    {

        #region Constructors, Initialization, and Load

        public CarLookupDataService(Func<ExploreDbContext> context)
        {
            Int64 startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            _contextCreator = context;

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region Enums


        #endregion

        #region Structures


        #endregion

        #region Fields and Properties

        private Func<ExploreDbContext> _contextCreator;

        #endregion

        #region Event Handlers


        #endregion

        #region Public Methods

        public async Task<IEnumerable<LookupItem>> GetCarLookupAsync()
        {
            Int64 startTicks = Log.DOMAINSERVICES("(CarLookupDataService) Enter", Common.LOG_CATEGORY);

            IEnumerable<LookupItem> result;

            using (var ctx = _contextCreator())
            {
                result = await ctx.CarsSet.AsNoTracking()
                  .Select(f =>
                  new LookupItem
                  {
                      Id = f.Id,
                      DisplayMember = f.FieldString
                  })
                  .ToListAsync();
            }

            Log.DOMAINSERVICES("(CarLookupDataService) Exit", Common.LOG_CATEGORY, startTicks);

            return result;
        }

        #endregion

        #region Protected Methods


        #endregion

        #region Private Methods


        #endregion

    }
}
