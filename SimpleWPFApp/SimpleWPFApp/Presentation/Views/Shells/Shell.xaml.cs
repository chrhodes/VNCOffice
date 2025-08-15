using System;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;

using SimpleWPFApp.Presentation.ViewModels;

using VNC;
using VNC.Core.Mvvm;

namespace SimpleWPFApp.Presentation.Views
{
    // FIX(crhodes)
    // Do not know why this doesn't work

    //public partial class Shell : WindowBase
    public partial class Shell : Window, IInstanceCountV, INotifyPropertyChanged
    {
        #region Constructors, Intialization, and Load

        public Shell()
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.Constructor) startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            InstanceCountV++;

            InitializeComponent();

            // Wire up ViewModel if needed

            // If View First with ViewModel in Xaml

            // ViewModel = (IViewModel)DataContext;

            // Can create directly

            // ViewModel = ViewModel();

            // Can use ourselves for everything

            //DataContext = this;

            InitializeView();

            if (Common.VNCLogging.Constructor) Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public Shell(ShellViewModel viewModel)
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.Constructor) startTicks = Log.CONSTRUCTOR($"Enter viewModel({viewModel.GetType()})", Common.LOG_CATEGORY);

            InstanceCountVP++;

            InitializeComponent();

            ViewModel = viewModel;      // AppVersionInfo needs this.
            DataContext = viewModel;

            // For the rare case where the ViewModel needs to know about the View
            // ViewModel.View = this;

            InitializeView();

            if (Common.VNCLogging.Constructor) Log.CONSTRUCTOR(String.Format("Exit"), Common.LOG_CATEGORY, startTicks);
        }

        private void InitializeView()
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.ViewLow) startTicks = Log.VIEW_LOW("Enter", Common.LOG_CATEGORY);

            // Store information about the View, DataContext, and ViewModel 
            // for the DeveloperInfo control. Useful for debugging binding issues
            // Set the DataConext to us.

            ViewType = this.GetType().ToString().Split('.').Last();
            ViewModelType = ViewModel?.GetType().ToString().Split('.').Last();
            ViewDataContextType = this.DataContext?.GetType().ToString().Split('.').Last();
            spDeveloperInfo.DataContext = this;

            // TODO(crhodes)
            // Put things here that initialize the View
            // Hook eventhandlers, etc.

            Common.CurrentShell = this;
            DeveloperUIMode = Common.DeveloperUIMode;

            // Establish any additional DataContext(s) to things held in this View  


            if (Common.VNCLogging.ViewLow) Log.VIEW_LOW("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region Fields and Properties

        private string _viewType;

        public string ViewType
        {
            get => _viewType;
            set
            {
                if (_viewType == value)
                {
                    return;
                }

                _viewType = value;
                OnPropertyChanged();
            }
        }

        private string _viewDataContextType;

        public string ViewDataContextType
        {
            get => _viewDataContextType;
            set
            {
                if (_viewDataContextType == value)
                {
                    return;
                }

                _viewDataContextType = value;
                OnPropertyChanged();
            }
        }

        private ShellViewModel viewModel;
        public ShellViewModel ViewModel
        {
            get => viewModel; set
            {
                if (viewModel == value)
                    return;
                viewModel = value;
                OnPropertyChanged();
            }
        }

        private string _viewModelType;

        public string ViewModelType
        {
            get => _viewModelType;
            set
            {
                if (_viewModelType == value)
                {
                    return;
                }

                _viewModelType = value;
                OnPropertyChanged();
            }
        }

        private Visibility _developerUIMode = Visibility.Collapsed;
        public Visibility DeveloperUIMode
        {
            get => _developerUIMode;
            set
            {
                if (_developerUIMode == value)
                    return;
                _developerUIMode = value;
                OnPropertyChanged();
            }
        }

        private Size _windowSize;
        public Size WindowSize
        {
            get => _windowSize;
            set
            {
                if (_windowSize == value)
                    return;
                _windowSize = value;
                OnPropertyChanged();
            }
        }

        #endregion

        #region EventHandlers

        private void thisControl_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            var newSize = e.NewSize;
            var previousSize = e.PreviousSize;
            WindowSize = newSize;
        }

        #endregion

        #region INotifyPropertyChanged

        public event PropertyChangedEventHandler PropertyChanged;

        // This is the traditional approach - requires string name to be passed in

        //private void OnPropertyChanged(string propertyName)
        //{
        //    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        //}

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            long startTicks = 0;
            if (Common.VNCLogging.INPC) startTicks = Log.VIEW_LOW($"Enter ({propertyName})", Common.LOG_CATEGORY);

            // This is the new CompilerServices attribute!

            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

            if (Common.VNCLogging.INPC) Log.VIEW_LOW("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

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
