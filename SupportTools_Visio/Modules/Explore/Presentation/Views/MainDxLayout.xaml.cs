using System;
using System.Windows;

using Explore.Presentation.ViewModels;

using VNC;
using VNC.Core.Mvvm;

namespace Explore.Presentation.Views
{
    public partial class MainDxLayout : ViewBase, IMain, IInstanceCountV
    {
        public MainDxLayoutViewModel _viewModel;

        public MainDxLayout(MainDxLayoutViewModel viewModel)
        {
            Int64 startTicks = Log.CONSTRUCTOR($"Enter viewModel({viewModel.GetType()}", Common.LOG_CATEGORY);

            InstanceCountVP++;
            InitializeComponent();

            _viewModel = viewModel;
            DataContext = _viewModel;

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #region IInstanceCount

        private static int _instanceCountV;

        public int InstanceCountV
        {
            get => _instanceCountV;
            set => _instanceCountV = value;
        }

        private static int _instanceCountVP;

        public int InstanceCountVP
        {
            get => _instanceCountVP;
            set => _instanceCountVP = value;
        }

        #endregion        
    }
}
