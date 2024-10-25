﻿using System;
using System.Windows;

using Explore.Presentation.ViewModels;

using VNC;
using VNC.Core.Mvvm;

namespace Explore.Presentation.Views
{
    public partial class Main : ViewBase, IMain, IInstanceCountV
    {
        public MainViewModel _viewModel;

        public Main(MainViewModel viewModel)
        {
            Int64 startTicks = Log.CONSTRUCTOR($"Enter viewModel({viewModel.GetType()})", Common.LOG_CATEGORY);

            InstanceCountVP++;
            InitializeComponent();

            _viewModel = viewModel;
            DataContext = _viewModel;

            Log.CONSTRUCTOR(String.Format("Exit"), Common.LOG_CATEGORY, startTicks);
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
