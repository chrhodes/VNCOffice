using System;
using System.Linq;
using System.Windows.Controls;

using VNC.Core.Mvvm;

namespace ModuleOptions
{
    public partial class ExcelOptions : UserControl, IExcelOptions
    {
        // View 1st approach.  
        // ViewModel is passed in constructor
        // Container must create ViewModel first so it can be passed in.

        //public AZDOOptions()
        //{
        //}
        public ExcelOptions(IExcelOptionsViewModel viewModel)
        {
            InitializeComponent();
            ViewModel = viewModel;
        }

        public IViewModel ViewModel
        {
            get
            {
                return (IExcelOptionsViewModel)DataContext;
            }
            set
            {
                DataContext = value;
            }
        }
    }
}
