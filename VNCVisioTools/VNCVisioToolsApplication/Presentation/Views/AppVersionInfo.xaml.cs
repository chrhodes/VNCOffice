using System;
using System.Linq;

using VNCVisioToolsApplication.Presentation.ViewModels;

using VNC;
using VNC.Core.Mvvm;
using VNC.Core;

namespace VNCVisioToolsApplication.Presentation.Views
{
    public partial class AppVersionInfo : ViewBase
    {
        #region Constructors, Initialization, and Load

        public AppVersionInfo()
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.Constructor) startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            InstanceCountV++;

            InitializeComponent();

            // Wire up ViewModel if needed

            // If View First with ViewModel in Xaml

            // ViewModel = (IAppVersionInfoViewModel)DataContext;

            // Can create directly

            // ViewModel = new AppVersionInfoViewModel();

            // Can use ourselves for everything

            //DataContext = this;

            InitializeView();

            if (Common.VNCLogging.Constructor) Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public AppVersionInfo(IAppVersionInfoViewModel viewModel)
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.Constructor) startTicks = Log.CONSTRUCTOR($"Enter viewModel({viewModel.GetType()}", Common.LOG_CATEGORY);

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
            // Set the DataContext to us.

            ViewType = this.GetType().ToString().Split('.').Last();
            ViewModelType = ViewModel?.GetType().ToString().Split('.').Last();
            ViewDataContextType = this.DataContext?.GetType().ToString().Split('.').Last();
            spDeveloperInfo.DataContext = this;

            // TODO(crhodes)
            // Put things here that initialize the View
            // Hook event handlers, etc.


            // Establish any additional DataContext(s) to things held in this View  

            // This gives us access to the ViewModelBase
            // which contains the Assembly and Runtime Information we need
            //
            // NB. This steps on the ViewModel = viewModel above.
            // Need to think through this if we put anything in AppVersionInfoViewModel

            // TODO(crhodes)
            // Maybe give a name to the control that contains everything.

            // HACK(crhodes)
            // Figure out what to do here as there is no current shell in an AddIn

            //InformationApplication = Common.InformationApplication;
            //InformationApplicationCore = Common.InformationApplicationCore;

            ////TODO(crhodes)
            //// Add additional Information InformationXXX for other assemblies

            //InformationVNCCore = Common.InformationVNCCore;

            spVNCVisioTools.DataContext = Common.InformationVNCVisioTools;
            lblRunTimeVersion.DataContext = Common.InformationVNCVisioTools;

            spVNCVisioToolsApplication.DataContext = Common.InformationVNCVisioToolsApplication;
            spVNCVisioToolsApplicationCore.DataContext = Common.InformationVNCVisioToolsApplicationCore;

            spVNCVSTOAddIn.DataContext = Common.InformationVNCVSTOAddIn;
            spVNCVisioVSTOAddIn.DataContext = Common.InformationVNCVisioVSTOAddIn;
            spVNCWpfPresentation.DataContext = Common.InformationVNCWpfPresentation;
            spVNCWpfPresentationDx.DataContext = Common.InformationVNCWpfPresentationDx;

            spVNCCore.DataContext = Common.InformationVNCCore;
  
            if (Common.VNCLogging.ViewLow) Log.VIEW_LOW("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region Enums (none)



        #endregion

        #region Structures (none)



        #endregion

        #region Fields and Properties (none)

        //public Information InformationApplication
        //{
        //    get;
        //    set;
        //}
        //public Information InformationApplicationCore
        //{
        //    get;
        //    set;
        //}

        ////TODO(crhodes)
        //// Add additional Information InformationXXX for other assemblies

        //public Information InformationVNCCore { get; set; }

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
    }
}
