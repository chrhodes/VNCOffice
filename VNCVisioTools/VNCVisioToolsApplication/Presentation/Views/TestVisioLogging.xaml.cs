using System;
using System.Linq;
using System.Windows;

using VNC;
using VNC.Core.Mvvm;

using VNCVisioToolsApplication.Presentation.ViewModels;

namespace VNCVisioToolsApplication.Presentation.Views
{
    public partial class TestVisioLogging : ViewBase, ITestVisioLogging, IInstanceCountV
    {
        #region Fields (none)



        #endregion

        #region Constructors, Initialization, and Load

        public TestVisioLogging()
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.Constructor) startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            InstanceCountV++;

            InitializeComponent();

            // Wire up ViewModel if needed

            // If View First with ViewModel in Xaml

            // ViewModel = (ITestExcelLoggingViewModel)DataContext;

            // Can create directly

            // ViewModel = TestExcelLoggingViewModel();

            // ViewModel = new TestExcelLoggingViewModel(
            // Common.EventAggregator,
            // (DialogService)Common.Container.Resolve(typeof(DialogService)));

            // Can use ourselves for everything

            // DataContext = this;

            // If no DataContext is set,
            // DataContext will come from parent.

            InitializeView();

            if (Common.VNCLogging.Constructor) Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public TestVisioLogging(ITestVisioLoggingViewModel viewModel)
        {
            Int64 startTicks = Log.CONSTRUCTOR($"Enter viewModel({viewModel.GetType()}", Common.LOG_CATEGORY);

            InstanceCountVP++;

            InitializeComponent();

            ViewModel = viewModel;  // ViewBase sets the DataContext to ViewModel

            // For the rare case where the ViewModel needs to know about the View
            // ViewModel.View = this;

            InitializeView();

            if (Common.VNCLogging.Constructor) Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void InitializeView()
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.ViewLow) startTicks = Log.VIEW_LOW("Enter", Common.LOG_CATEGORY);

            // Store information about the View, DataContext, and ViewModel
            // for the DeveloperInfo control. Useful for debugging binding issues

            ViewType = this.GetType().ToString().Split('.').Last();
            ViewModelType = ViewModel?.GetType().ToString().Split('.').Last();
            ViewDataContextType = this.DataContext?.GetType().ToString().Split('.').Last();

            // Set the DataContext to us.
            spDeveloperInfo.DataContext = this;

            // TODO(crhodes)
            // Put things here that initialize the View
            // Hook EventHandlers, etc.


            // Establish any additional DataContext(s) to things held in this View

            if (Common.VNCLogging.ViewLow) Log.VIEW_LOW("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region Enums (none)



        #endregion

        #region Structures (none)



        #endregion

        #region Properties (none)



        #endregion

        #region Event Handlers (none)



        #endregion

        #region Commands (none)



        #endregion

        #region Public Methods (none)



        #endregion

        #region Protected Methods (none)



        #endregion

        #region Private Methods (none)



        #endregion

        #region IInstanceCountV

        private static Int32 _instanceCountV;

        public Int32 InstanceCountV
        {
            get => _instanceCountV;
            set => _instanceCountV = value;
        }

        private static Int32 _instanceCountVP;

        public Int32 InstanceCountVP
        {
            get => _instanceCountVP;
            set => _instanceCountVP = value;
        }

        #endregion
    }
}
