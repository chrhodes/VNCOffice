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

            if (themedWindowHostModal is null) themedWindowHostModal = new DxThemedWindowHost();

            themedWindowHostModal.DisplayUserControlInHost(
                "ThemedWindowHost (Modal)",
                Common.DEFAULT_WINDOW_WIDTH_SMALL, Common.DEFAULT_WINDOW_HEIGHT_SMALL,
                ShowWindowMode.Modal_ShowDialog);

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private DxThemedWindowHost themedWindowHostModeless = null;

        private void btnThemedWindowHostModeless_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (themedWindowHostModeless is null) themedWindowHostModeless = new DxThemedWindowHost();

            themedWindowHostModeless.DisplayUserControlInHost(
                "ThemedWindowHost (Modeless)",
                Common.DEFAULT_WINDOW_WIDTH_SMALL, Common.DEFAULT_WINDOW_HEIGHT_SMALL,
                ShowWindowMode.Modeless_Show);

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private DxWindowHost dxWindowHost = null;

        private void btnDxWindowHost_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (dxWindowHost is null) dxWindowHost = new DxWindowHost();

            dxWindowHost.DisplayUserControlInHost(
                "DxWindowHost Test",
                Common.DEFAULT_WINDOW_WIDTH_SMALL, Common.DEFAULT_WINDOW_HEIGHT_SMALL,
                ShowWindowMode.Modeless_Show);

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static WindowHost windowHostLocal = null;

        private void btnWindowHostLocal_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (windowHostLocal is null) windowHostLocal = new WindowHost();

            windowHostLocal.DisplayUserControlInHost(
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

            if (cylonHost is null) cylonHost = new WindowHost();

            cylonHost.DisplayUserControlInHost(
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

            if (cylonHost2 is null) cylonHost2 = new WindowHost();

            cylonHost2.DisplayUserControlInHost(
                "I am a Cylon loaded by type",
                Common.DEFAULT_WINDOW_WIDTH_SMALL, Common.DEFAULT_WINDOW_HEIGHT_SMALL,
                ShowWindowMode.Modeless_Show,
                new CylonEyeBall());

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private DxThemedWindowHost dxLayoutControlHost = null;

        private void btnDxLayoutControl_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (dxLayoutControlHost is null) dxLayoutControlHost = new DxThemedWindowHost();

            dxLayoutControlHost.DisplayUserControlInHost(
                "DxLayoutControl Test",
                Common.DEFAULT_WINDOW_WIDTH_LARGE, Common.DEFAULT_WINDOW_HEIGHT_LARGE,
                ShowWindowMode.Modeless_Show,
                new DxLayoutControl());

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private DxThemedWindowHost dxDockLayoutControlHost = null;

        private void btnDxDockLayoutControl_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (dxDockLayoutControlHost is null) dxDockLayoutControlHost = new DxThemedWindowHost();

            dxDockLayoutControlHost.DisplayUserControlInHost(
                "DxDockLayoutControl Test",
                Common.DEFAULT_WINDOW_WIDTH_LARGE, Common.DEFAULT_WINDOW_HEIGHT_LARGE,
                ShowWindowMode.Modeless_Show,
                new DxDockLayoutControl());

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private DxThemedWindowHost dxDockLayoutManagerControlHost = null;

        private void btnDxDockLayoutManagerControl_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (dxDockLayoutManagerControlHost is null) dxDockLayoutManagerControlHost = new DxThemedWindowHost();

            dxDockLayoutManagerControlHost.DisplayUserControlInHost(
                "DxDocLayoutManagerControl Test",
                Common.DEFAULT_WINDOW_WIDTH_LARGE, Common.DEFAULT_WINDOW_HEIGHT_LARGE,
                ShowWindowMode.Modeless_Show,
                new DxDockLayoutManagerControl());

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private DxThemedWindowHost prismRegionTestHost = null;

        private void btnPrismRegionTest_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (prismRegionTestHost is null) prismRegionTestHost = new DxThemedWindowHost();

            prismRegionTestHost.DisplayUserControlInHost(
                "Prism Region Test", 
                Common.DEFAULT_WINDOW_WIDTH_LARGE, Common.DEFAULT_WINDOW_HEIGHT_LARGE,
                ShowWindowMode.Modeless_Show,
                new PrismRegionTest());

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion WPF UI Events

        #region MVVM Examples

        public static DxThemedWindowHost vncMVVM_VAVM1st_Host = null;

        private void btnVNC_MVVM_VAVM1st_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (vncMVVM_VAVM1st_Host is null) vncMVVM_VAVM1st_Host = new DxThemedWindowHost();

            // NOTE(crhodes)
            // Wire things up ourselves - ViewModel First - with a little help from DI.

            vncMVVM_VAVM1st_Host.DisplayUserControlInHost(
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
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (vncMVVM_VA_Host is null) vncMVVM_VA_Host = new DxThemedWindowHost();
            // NOTE(crhodes)
            // This does not wire View to ViewModel
            // Because we HAVE NOT Registered ViewAViewModel in SupportTools_VisioModules
            // Parameterless ViewA() constructor is called.

            vncMVVM_VA_Host.DisplayUserControlInHost(
                "MVVM ViewA First - No Registrations - DI Resolve",
                Common.DEFAULT_WINDOW_WIDTH_SMALL, Common.DEFAULT_WINDOW_HEIGHT_SMALL,
                ShowWindowMode.Modeless_Show,
                (ViewA)Common.ApplicationBootstrapper.Container.Resolve(typeof(ViewA))
            );

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost vncMVVM_VAVMDI_Host = null;

        private void btnVNC_MVVM_VAVM1stDI_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (vncMVVM_VAVMDI_Host is null) vncMVVM_VAVMDI_Host = new DxThemedWindowHost();

            // NOTE(crhodes)
            // This does wire View to ViewModel
            // Because ViewModel is passed a View (DI) and wires itself to View

            vncMVVM_VAVMDI_Host.DisplayUserControlInHost(
                "MVVM ViewAViewModel First (ViewModel is passed new View) - DI Resolve",
                Common.DEFAULT_WINDOW_WIDTH_SMALL, Common.DEFAULT_WINDOW_HEIGHT_SMALL,
                ShowWindowMode.Modeless_Show,
                (ViewAViewModel)Common.ApplicationBootstrapper.Container.Resolve(typeof(ViewAViewModel))
            );

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost vncMVVM_VB_Host = null;

        private void btnVNC_MVVM_VB1st_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (vncMVVM_VB_Host is null) vncMVVM_VB_Host = new DxThemedWindowHost();

            // NOTE(crhodes)
            // This does wire View to ViewModel
            // Because we have Registered ViewBViewModel in SupportTools_VisioModules

            vncMVVM_VB_Host.DisplayUserControlInHost(
                "MVVM ViewB First (View is passed new ViewModel) DI Resolve",
                Common.DEFAULT_WINDOW_WIDTH_SMALL, Common.DEFAULT_WINDOW_HEIGHT_SMALL,
                ShowWindowMode.Modeless_Show,
                (ViewB)Common.ApplicationBootstrapper.Container.Resolve(typeof(ViewB))
            );

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost vncMVVM_VC1_Host = null;

        private void btnVNC_MVVM_VC11st_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (vncMVVM_VC1_Host is null) vncMVVM_VC1_Host = new DxThemedWindowHost();

            // NOTE(crhodes)
            // This does wire View to ViewModel
            // C1 has C1() and C1(ViewModel) constructors. No DI Registrations
            // NB.  AutoWireViewModel=false

            vncMVVM_VC1_Host.DisplayUserControlInHost(
                "MVVM ViewC1 First.  ViewC1 has C1() and C1(ViewModel) constructors. No DI Registrations",
                Common.DEFAULT_WINDOW_WIDTH_SMALL, Common.DEFAULT_WINDOW_HEIGHT_SMALL,
                ShowWindowMode.Modeless_Show,
                (ViewC1)Common.ApplicationBootstrapper.Container.Resolve(typeof(ViewC1))
            );

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost vncMVVM_VC2_Host = null;

        private void btnVNC_MVVM_VC21st_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (vncMVVM_VC2_Host is null) vncMVVM_VC2_Host = new DxThemedWindowHost();

            // NOTE(crhodes)
            // This does wire View to ViewModel
            // Because we have removed the default ViewC2 Constructor
            // and Registered ViewCViewModel in SupportTools_VisioModules

            vncMVVM_VC2_Host.DisplayUserControlInHost(
                "MVVM ViewC2 First (View is passed new ViewModel) DI Resolve",
                Common.DEFAULT_WINDOW_WIDTH_SMALL, Common.DEFAULT_WINDOW_HEIGHT_SMALL,
                ShowWindowMode.Modeless_Show,
                (ViewC2)Common.ApplicationBootstrapper.Container.Resolve(typeof(ViewC2))
            );

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #endregion Event Handlers
    }
}