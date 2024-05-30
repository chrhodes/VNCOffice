using System.Collections.Generic;

using VNC.Core.Mvvm;

namespace ModuleOptions
{
    public class ExcelOptionsViewModel : IExcelOptionsViewModel
    {
        // View 1st approach.  
        // ViewModel is not passed a View in constructor
        public ExcelOptionsViewModel()
        {

        }

        // ViewModel first approach.  
        // View is passed in constructor

        //public ExcelOptionsViewModel(IExcelOptions view)
        //{
        //    View = view;
        //    // Point the view to this ViewModel
        //    View.ViewModel = this;
        //}

        public IView View
        {
            get;
            set;
        }

        public List<string> TeamProjects = new List<string>();
    }
}
