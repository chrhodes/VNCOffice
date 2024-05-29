using System;
using System.Windows;

using Microsoft.Office.Tools.Ribbon;

using Prism.Events;
using Prism.Services.Dialogs;

using SupportTools_Visio.Presentation.ViewModels;
using SupportTools_Visio.Presentation.Views;

using VNC;
using VNC.Core.Presentation;
using VNC.WPF.Presentation.Dx.Views;
using VNC.WPF.Presentation.Views;

namespace SupportTools_Visio
{
    public partial class Ribbon
    {
        #region Event Handlers

        #region UI Launch Events

        private DxThemedWindowHost themedWindowHostModal = null;

        private void btnThemedWindowHostModal_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            DxThemedWindowHost.DisplayUserControlInHost(ref themedWindowHostModal,
                "ThemedWindowHost (Modal)",
                Common.DEFAULT_WINDOW_WIDTH_SMALL, Common.DEFAULT_WINDOW_HEIGHT_SMALL,
                ShowWindowMode.Modal_ShowDialog);

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private DxThemedWindowHost themedWindowHostModeless = null;

        private void btnThemedWindowHostModeless_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            DxThemedWindowHost.DisplayUserControlInHost(ref themedWindowHostModeless,
                "ThemedWindowHost (Modeless)",
                Common.DEFAULT_WINDOW_WIDTH_SMALL, Common.DEFAULT_WINDOW_HEIGHT_SMALL,
                ShowWindowMode.Modeless_Show);

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private DxWindowHost dxWindowHost = null;

        private void btnDxWindowHost_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            DxWindowHost.DisplayUserControlInHost(ref dxWindowHost,
                "DxWindowHost Test",
                Common.DEFAULT_WINDOW_WIDTH_SMALL, Common.DEFAULT_WINDOW_HEIGHT_SMALL,
                ShowWindowMode.Modeless_Show);

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static WindowHost windowHostLocal = null;

        private void btnWindowHostLocal_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            WindowHost.DisplayUserControlInHost(ref windowHostLocal,
                "WindowHost (local) Test",
                Common.DEFAULT_WINDOW_WIDTH_SMALL, Common.DEFAULT_WINDOW_HEIGHT_SMALL,
                ShowWindowMode.Modeless_Show);

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static WindowHost windowHostVNC = null;

        private void btnWindowHostVNC_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            ShowEmptyHost(windowHostVNC, "WindowHost (VNC)", ShowWindowMode.Modeless_Show);

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private static void ShowEmptyHost(Window host, string title, ShowWindowMode mode)
        {
            long startTicks = Log.PRESENTATION("Enter", Common.LOG_CATEGORY);

            if (host is null)
            {
                host = new DxThemedWindowHost();
                host.Height = Common.DEFAULT_WINDOW_HEIGHT_SMALL;
                host.Width = Common.DEFAULT_WINDOW_WIDTH_SMALL;
                host.Title = title;
            }

            if (mode == ShowWindowMode.Modal_ShowDialog)
            {
                long endTicks2 = Log.PRESENTATION("Exit", Common.LOG_CATEGORY, startTicks);

                host.Title = $"{host.GetType()} loadtime: {Log.GetDuration(startTicks, endTicks2)}";

                host.ShowDialog();
            }
            else
            {
                long endTicks2 = Log.PRESENTATION("Exit", Common.LOG_CATEGORY, startTicks);

                host.Title = $"{host.GetType()} loadtime: {Log.GetDuration(startTicks, endTicks2)}";

                host.Show();
            }

            long endTicks = Log.PRESENTATION("Exit", Common.LOG_CATEGORY, startTicks);

            host.Title = $"{host.GetType()} loadtime: {Log.GetDuration(startTicks, endTicks)}";
        }

        #endregion UI Launch Events

        #region WPF UI Events

        public static WindowHost cylonHost = null;

        private void btnLaunchCylon_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            WindowHost.DisplayUserControlInHost(ref cylonHost,
                "I am a Cylon loaded by name",
                Common.DEFAULT_WINDOW_WIDTH_SMALL, Common.DEFAULT_WINDOW_HEIGHT_SMALL,
                ShowWindowMode.Modeless_Show,
                "VNC.WPF.Presentation.Views.CylonEyeBall, VNC.WPF.Presentation");

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static WindowHost cylonHost2 = null;

        private void btnLaunchCylon2_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            WindowHost.DisplayUserControlInHost(ref cylonHost2,
                "I am a Cylon loaded by type",
                Common.DEFAULT_WINDOW_WIDTH_SMALL, Common.DEFAULT_WINDOW_HEIGHT_SMALL,
                ShowWindowMode.Modeless_Show,
                new CylonEyeBall());

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        DxLayoutControl _dxLayoutControl;

        public DxLayoutControl DxLayoutControl
        {
            get
            {
                if (_dxLayoutControl is null)
                {
                    _dxLayoutControl = new DxLayoutControl();
                }

                return _dxLayoutControl;
            }
            set
            {
                _dxLayoutControl = value;
            }
        }

        private DxThemedWindowHost dxLayoutControlHost = null;

        private void btnDxLayoutControl_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            DxThemedWindowHost.DisplayUserControlInHost(ref dxLayoutControlHost,
                "DxLayoutControl Test",
                Common.DEFAULT_WINDOW_WIDTH_LARGE, Common.DEFAULT_WINDOW_HEIGHT_LARGE,
                ShowWindowMode.Modeless_Show,
                DxLayoutControl);

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        DxDockLayoutControl _dxDockLayoutControl;

        public DxDockLayoutControl DxDockLayoutControl
        {
            get
            {
                if (_dxDockLayoutControl is null)
                {
                    _dxDockLayoutControl = new DxDockLayoutControl();
                }

                return _dxDockLayoutControl;
            }
            set
            {
                _dxDockLayoutControl = value;
            }
        }

        private DxThemedWindowHost dxDockLayoutControlHost = null;

        private void btnDxDockLayoutControl_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            DxThemedWindowHost.DisplayUserControlInHost(ref dxDockLayoutControlHost,
                "DxDockLayoutControl Test",
                Common.DEFAULT_WINDOW_WIDTH_LARGE, Common.DEFAULT_WINDOW_HEIGHT_LARGE,
                ShowWindowMode.Modeless_Show,
                DxDockLayoutControl);

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        DxDockLayoutManagerControl _dxDockLayoutControlManager;

        public DxDockLayoutManagerControl DxDockLayoutManagerControl
        {
            get
            {
                if (_dxDockLayoutControlManager is null)
                {
                    _dxDockLayoutControlManager = new DxDockLayoutManagerControl();
                }

                return _dxDockLayoutControlManager;
            }
            set
            {
                _dxDockLayoutControlManager = value;
            }
        }

        private DxThemedWindowHost dxDockLayoutManagerControlHost = null;

        private void btnDxDockLayoutManagerControl_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            DxThemedWindowHost.DisplayUserControlInHost(ref dxDockLayoutManagerControlHost,
                "DxDocLayoutManagerControl Test",
                Common.DEFAULT_WINDOW_WIDTH_LARGE, Common.DEFAULT_WINDOW_HEIGHT_LARGE,
                ShowWindowMode.Modeless_Show,
                DxDockLayoutManagerControl);

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        Presentation.Views.PrismRegionTest _prismRegionTest;

        public Presentation.Views.PrismRegionTest PrismRegionTest
        {
            get
            {
                if (_prismRegionTest is null)
                {
                    _prismRegionTest = new Presentation.Views.PrismRegionTest();
                }

                return _prismRegionTest;
            }
            set
            {
                _prismRegionTest = value;
            }
        }

        private DxThemedWindowHost prismRegionTestHost = null;

        private void btnPrismRegionTest_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            DxThemedWindowHost.DisplayUserControlInHost(ref prismRegionTestHost,
                "Prism Region Test", 
                Common.DEFAULT_WINDOW_WIDTH_LARGE, Common.DEFAULT_WINDOW_HEIGHT_LARGE,
                ShowWindowMode.Modeless_Show,
                PrismRegionTest);

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion WPF UI Events

        #region MVVM Examples

        public static DxThemedWindowHost vncMVVM_VAVM1st_Host = null;

        private void btnVNC_MVVM_VAVM1st_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            // NOTE(crhodes)
            // Wire things up ourselves - ViewModel First - with a little help from DI.

            DxThemedWindowHost.DisplayUserControlInHost(ref vncMVVM_VAVM1st_Host,
                "MVVM ViewAViewModel First (ViewModel is passed new ViewA) - By Hand",
                Common.DEFAULT_WINDOW_WIDTH, Common.DEFAULT_WINDOW_HEIGHT,
                ShowWindowMode.Modeless_Show,
                new ViewAViewModel(
                    new ViewA(),
                    (IEventAggregator)Common.ApplicationBootstrapper.Container.Resolve(typeof(EventAggregator)),
                    (DialogService)Common.ApplicationBootstrapper.Container.Resolve(typeof(DialogService))
                )
            );

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost vncMVVM_VA_Host = null;

        private void btnVNC_MVVM_VA1st_Click(object sender, RibbonControlEventArgs e)
        {
            // NOTE(crhodes)
            // This does not wire View to ViewModel
            // Because we HAVE NOT Registered ViewAViewModel in SupportTools_VisioModules
            // Parameterless ViewA() constructor is called.

            DxThemedWindowHost.DisplayUserControlInHost(ref vncMVVM_VA_Host,
                "MVVM ViewA First - No Registrations - DI Resolve",
                Common.DEFAULT_WINDOW_WIDTH_SMALL, Common.DEFAULT_WINDOW_HEIGHT_SMALL,
                ShowWindowMode.Modeless_Show,
                (ViewA)Common.ApplicationBootstrapper.Container.Resolve(typeof(ViewA))
            );
        }

        public static DxThemedWindowHost vncMVVM_VAVMDI_Host = null;

        private void btnVNC_MVVM_VAVM1stDI_Click(object sender, RibbonControlEventArgs e)
        {
            // NOTE(crhodes)
            // This does wire View to ViewModel
            // Because ViewModel is passed a View (DI) and wires itself to View

            DxThemedWindowHost.DisplayUserControlInHost(ref vncMVVM_VAVMDI_Host,
                "MVVM ViewAViewModel First (ViewModel is passed new View) - DI Resolve",
                Common.DEFAULT_WINDOW_WIDTH_SMALL, Common.DEFAULT_WINDOW_HEIGHT_SMALL,
                ShowWindowMode.Modeless_Show,
                (ViewAViewModel)Common.ApplicationBootstrapper.Container.Resolve(typeof(ViewAViewModel))
            );
        }

        public static DxThemedWindowHost vncMVVM_VB_Host = null;

        private void btnVNC_MVVM_VB1st_Click(object sender, RibbonControlEventArgs e)
        {
            // NOTE(crhodes)
            // This does wire View to ViewModel
            // Because we have Registered ViewBViewModel in SupportTools_VisioModules

            DxThemedWindowHost.DisplayUserControlInHost(ref vncMVVM_VB_Host,
                "MVVM ViewB First (View is passed new ViewModel) DI Resolve",
                Common.DEFAULT_WINDOW_WIDTH_SMALL, Common.DEFAULT_WINDOW_HEIGHT_SMALL,
                ShowWindowMode.Modeless_Show,
                (ViewB)Common.ApplicationBootstrapper.Container.Resolve(typeof(ViewB))
            );
        }

        public static DxThemedWindowHost vncMVVM_VC1_Host = null;

        private void btnVNC_MVVM_VC11st_Click(object sender, RibbonControlEventArgs e)
        {
            // NOTE(crhodes)
            // This does wire View to ViewModel
            // C1 has C1() and C1(ViewModel) constructors. No DI Registrations
            // NB.  AutoWireViewModel=false

            DxThemedWindowHost.DisplayUserControlInHost(ref vncMVVM_VC1_Host,
                "MVVM ViewC1 First.  ViewC1 has C1() and C1(ViewModel) constructors. No DI Registrations",
                Common.DEFAULT_WINDOW_WIDTH_SMALL, Common.DEFAULT_WINDOW_HEIGHT_SMALL,
                ShowWindowMode.Modeless_Show,
                (ViewC1)Common.ApplicationBootstrapper.Container.Resolve(typeof(ViewC1))
            );
        }

        public static DxThemedWindowHost vncMVVM_VC2_Host = null;

        private void btnVNC_MVVM_VC21st_Click(object sender, RibbonControlEventArgs e)
        {
            // NOTE(crhodes)
            // This does wire View to ViewModel
            // Because we have removed the default ViewC2 Constructor
            // and Registered ViewCViewModel in SupportTools_VisioModules

            DxThemedWindowHost.DisplayUserControlInHost(ref vncMVVM_VC2_Host,
                "MVVM ViewC2 First (View is passed new ViewModel) DI Resolve",
                Common.DEFAULT_WINDOW_WIDTH_SMALL, Common.DEFAULT_WINDOW_HEIGHT_SMALL,
                ShowWindowMode.Modeless_Show,
                (ViewC2)Common.ApplicationBootstrapper.Container.Resolve(typeof(ViewC2))
            );
        }

        #endregion

        #endregion Event Handlers
    }
}