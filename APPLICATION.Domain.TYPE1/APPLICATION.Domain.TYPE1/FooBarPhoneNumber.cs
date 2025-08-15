using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

using VNC.Core.DomainServices;

namespace APPLICATION.Domain.Domain
{
    public class FooBarPhoneNumber : IEntity<int>, IModificationHistory, IOptimistic
    {
        #region IEntity<int>

        public int Id { get; set; }

        #endregion

        [Phone]
        [Required]
        public string Number { get; set; }

        public int FooBarId { get; set; }

        public FooBar FooBar { get; set; }

        #region IModificationHistory

        public DateTime? DateModified { get; set; }

        public DateTime? DateCreated { get; set; }

        public Boolean? IsDirty { get; set; }

        #endregion

        #region IOptimistic

        // Need to have data annotation here.  
        // Presence in interface ignored.
        [Timestamp]
        public byte[] RowVersion { get; set; }

        #endregion
    }
}
