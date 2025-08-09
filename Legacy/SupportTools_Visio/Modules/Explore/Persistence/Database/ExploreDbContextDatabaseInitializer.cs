using System;
using System.Data.Entity;

using VNC;

namespace Explore.Persistence.Data
{
    public class ExploreDbContextDatabaseInitializer : CreateDatabaseIfNotExists<ExploreDbContext>
    {
        protected override void Seed(ExploreDbContext context)
        {
            Int64 startTicks = Log.PERSISTENCE("Enter", Common.LOG_CATEGORY);

            base.Seed(context);

            Log.PERSISTENCE("Exit", Common.LOG_CATEGORY, startTicks);
        }
    }
}
