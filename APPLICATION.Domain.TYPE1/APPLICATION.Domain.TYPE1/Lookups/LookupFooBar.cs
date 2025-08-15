using VNC.Core.DomainServices;

namespace APPLICATION.Domain.Domain.Lookups
{
    public class LookupFooBar : ILookupItem<int>
    {
        public int Id { get; set; }
        public string DisplayMember { get; set; }
    }
}
