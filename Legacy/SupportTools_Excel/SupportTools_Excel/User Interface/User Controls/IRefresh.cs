using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SupportTools_Excel.User_Interface
{
    /// <summary>
    /// Implement this interface in a Page
    /// to be notified when the currently
    /// logged in user changes.
    /// </summary>
    public interface IRefresh
    {
        /// <summary>
        /// Called by MainForm when the currently
        /// logged in user changes.
        /// </summary>
        void Refresh();
    }
}
