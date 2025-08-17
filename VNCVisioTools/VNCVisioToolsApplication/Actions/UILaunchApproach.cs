using System.Reflection;
using System.Windows;

using Prism.Events;
using Prism.Services.Dialogs;

using VNC.Core.Presentation;
using VNC.WPF.Presentation.Dx.Views;
using VNC.WPF.Presentation.Views;

using VNCVisioToolsApplication.Presentation.ViewModels;
using VNCVisioToolsApplication.Presentation.Views;

namespace VNCVisioToolsApplication.Actions
{
    public class UILaunchApproach
    {
        #region UI Launch Events

        private static DxThemedWindowHost themedWindowHostModal = null;

        public static void ThemedWindowHostModal()
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            if (themedWindowHostModal is null) themedWindowHostModal = new DxThemedWindowHost();

            themedWindowHostModal.DisplayUserControlInHost(
                "ThemedWindowHost (Modal)",
                Common.DEFAULT_WINDOW_WIDTH_SMALL, Common.DEFAULT_WINDOW_HEIGHT_SMALL,
                ShowWindowMode.Modal_ShowDialog);
        }

        private static DxThemedWindowHost themedWindowHostModeless = null;

        public static void ThemedWindowHostModeless()
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            if (themedWindowHostModeless is null) themedWindowHostModeless = new DxThemedWindowHost();

            themedWindowHostModeless.DisplayUserControlInHost(
                "ThemedWindowHost (Modeless)",
                Common.DEFAULT_WINDOW_WIDTH_SMALL, Common.DEFAULT_WINDOW_HEIGHT_SMALL,
                ShowWindowMode.Modeless_Show);
        }

        private static DxWindowHost dxWindowHost = null;

        public static void DxWindowHost()
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            if (dxWindowHost is null) dxWindowHost = new DxWindowHost();

            dxWindowHost.DisplayUserControlInHost(
                "DxWindowHost Test",
                Common.DEFAULT_WINDOW_WIDTH_SMALL, Common.DEFAULT_WINDOW_HEIGHT_SMALL,
                ShowWindowMode.Modeless_Show);
        }

        public static WindowHost windowHostLocal = null;

        public static void WindowHostLocal()
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            if (windowHostLocal is null) windowHostLocal = new WindowHost(Common.EventAggregator);

            windowHostLocal.DisplayUserControlInHost(
                "WindowHost (local) Test",
                Common.DEFAULT_WINDOW_WIDTH_SMALL, Common.DEFAULT_WINDOW_HEIGHT_SMALL,
                ShowWindowMode.Modeless_Show);
        }

        public static WindowHost windowHostVNC = null;

        public static void WindowHostVNC()
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            ShowEmptyHost(windowHostVNC, "WindowHost (VNC)", ShowWindowMode.Modeless_Show);
        }

        private static void ShowEmptyHost(Window host, string title, ShowWindowMode mode)
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");
            //long startTicks = Log.PRESENTATION("Enter", Common.LOG_CATEGORY);

            if (host is null)
            {
                host = new DxThemedWindowHost();
                host.Height = Common.DEFAULT_WINDOW_HEIGHT_SMALL;
                host.Width = Common.DEFAULT_WINDOW_WIDTH_SMALL;
                host.Title = title;
            }

            if (mode == ShowWindowMode.Modal_ShowDialog)
            {
                //long endTicks2 = Log.PRESENTATION("Exit", Common.LOG_CATEGORY, startTicks);

                //host.Title = $"{host.GetType()} loadtime: {Log.GetDuration(startTicks, endTicks2)}";

                host.ShowDialog();
            }
            else
            {
                //long endTicks2 = Log.PRESENTATION("Exit", Common.LOG_CATEGORY, startTicks);

                //host.Title = $"{host.GetType()} loadtime: {Log.GetDuration(startTicks, endTicks2)}";

                host.Show();
            }

            //long endTicks = Log.PRESENTATION("Exit", Common.LOG_CATEGORY, startTicks);

            //host.Title = $"{host.GetType()} loadtime: {Log.GetDuration(startTicks, endTicks)}";
        }

        #endregion UI Launch Events

        #region WPF UI Events

        public static WindowHost cylonHost = null;

        public static void LaunchCylon()
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            if (cylonHost is null) cylonHost = new WindowHost(Common.EventAggregator);

            cylonHost.DisplayUserControlInHost(
                "I am a Cylon loaded by name",
                Common.DEFAULT_WINDOW_WIDTH_SMALL, Common.DEFAULT_WINDOW_HEIGHT_SMALL,
                ShowWindowMode.Modeless_Show,
                "VNC.WPF.Presentation.Views.CylonEyeBall, VNC.WPF.Presentation");
        }

        public static WindowHost cylonHost2 = null;

        public static void LaunchCylon2()
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            if (cylonHost2 is null) cylonHost2 = new WindowHost(Common.EventAggregator);

            cylonHost2.DisplayUserControlInHost(
                "I am a Cylon loaded by type",
                Common.DEFAULT_WINDOW_WIDTH_SMALL, Common.DEFAULT_WINDOW_HEIGHT_SMALL,
                ShowWindowMode.Modeless_Show,
                new CylonEyeBall());
        }

        private static DxThemedWindowHost dxLayoutControlHost = null;

        public static void DxLayoutControl()
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            if (dxLayoutControlHost is null) dxLayoutControlHost = new DxThemedWindowHost();

            dxLayoutControlHost.DisplayUserControlInHost(
                "DxLayoutControl Test",
                Common.DEFAULT_WINDOW_WIDTH_LARGE, Common.DEFAULT_WINDOW_HEIGHT_LARGE,
                ShowWindowMode.Modeless_Show,
                new DxLayoutControl());
        }

        private static DxThemedWindowHost dxDockLayoutControlHost = null;

        public static void DxDockLayoutControl()
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            if (dxDockLayoutControlHost is null) dxDockLayoutControlHost = new DxThemedWindowHost();

            dxDockLayoutControlHost.DisplayUserControlInHost(
                "DxDockLayoutControl Test",
                Common.DEFAULT_WINDOW_WIDTH_LARGE, Common.DEFAULT_WINDOW_HEIGHT_LARGE,
                ShowWindowMode.Modeless_Show,
                new DxDockLayoutControl());
        }

        private static DxThemedWindowHost dxDockLayoutManagerControlHost = null;

        public static void DxDockLayoutManagerControl()
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            if (dxDockLayoutManagerControlHost is null) dxDockLayoutManagerControlHost = new DxThemedWindowHost();

            dxDockLayoutManagerControlHost.DisplayUserControlInHost(
                "DxDocLayoutManagerControl Test",
                Common.DEFAULT_WINDOW_WIDTH_LARGE, Common.DEFAULT_WINDOW_HEIGHT_LARGE,
                ShowWindowMode.Modeless_Show,
                new DxDockLayoutManagerControl());
        }

        private static DxThemedWindowHost prismRegionTestHost = null;

        public static void PrismRegionTest()
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            if (prismRegionTestHost is null) prismRegionTestHost = new DxThemedWindowHost();

            prismRegionTestHost.DisplayUserControlInHost(
                "Prism Region Test",
                Common.DEFAULT_WINDOW_WIDTH_LARGE, Common.DEFAULT_WINDOW_HEIGHT_LARGE,
                ShowWindowMode.Modeless_Show,
                new PrismRegionTest());
        }

        #endregion WPF UI Events

        #region MVVM Examples

        public static DxThemedWindowHost vncMVVM_VAVM1st_Host = null;

        public static void VNC_MVVM_VAVM1st()
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

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
        }

        public static DxThemedWindowHost vncMVVM_VA_Host = null;

        public static void VNC_MVVM_VA1st()
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            if (vncMVVM_VA_Host is null) vncMVVM_VA_Host = new DxThemedWindowHost();
            // NOTE(crhodes)
            // This does not wire View to ViewModel
            // Because we HAVE NOT Registered ViewAViewModel in VNCVisioToolsApplicationModules
            // Parameterless ViewA() constructor is called.

            vncMVVM_VA_Host.DisplayUserControlInHost(
                "MVVM ViewA First - No Registrations - DI Resolve",
                Common.DEFAULT_WINDOW_WIDTH_SMALL, Common.DEFAULT_WINDOW_HEIGHT_SMALL,
                ShowWindowMode.Modeless_Show,
                (ViewA)Common.ApplicationBootstrapper.Container.Resolve(typeof(ViewA))
            );
        }

        public static DxThemedWindowHost vncMVVM_VAVMDI_Host = null;

        public static void VNC_MVVM_VAVM1stDI()
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

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
        }

        public static DxThemedWindowHost vncMVVM_VB_Host = null;

        public static void VNC_MVVM_VB1st()
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            if (vncMVVM_VB_Host is null) vncMVVM_VB_Host = new DxThemedWindowHost();

            // NOTE(crhodes)
            // This does wire View to ViewModel
            // Because we have Registered ViewBViewModel in VNCVisioToolsApplicationModules

            vncMVVM_VB_Host.DisplayUserControlInHost(
                "MVVM ViewB First (View is passed new ViewModel) DI Resolve",
                Common.DEFAULT_WINDOW_WIDTH_SMALL, Common.DEFAULT_WINDOW_HEIGHT_SMALL,
                ShowWindowMode.Modeless_Show,
                (ViewB)Common.ApplicationBootstrapper.Container.Resolve(typeof(ViewB))
            );
        }

        public static DxThemedWindowHost vncMVVM_VC1_Host = null;

        public static void VNC_MVVM_VC11st()
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

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
        }

        public static DxThemedWindowHost vncMVVM_VC2_Host = null;

        public static void VNC_MVVM_VC21st()
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            if (vncMVVM_VC2_Host is null) vncMVVM_VC2_Host = new DxThemedWindowHost();

            // NOTE(crhodes)
            // This does wire View to ViewModel
            // Because we have removed the default ViewC2 Constructor
            // and Registered ViewCViewModel in VNCVisioToolsApplicationModules

            vncMVVM_VC2_Host.DisplayUserControlInHost(
                "MVVM ViewC2 First (View is passed new ViewModel) DI Resolve",
                Common.DEFAULT_WINDOW_WIDTH_SMALL, Common.DEFAULT_WINDOW_HEIGHT_SMALL,
                ShowWindowMode.Modeless_Show,
                (ViewC2)Common.ApplicationBootstrapper.Container.Resolve(typeof(ViewC2))
            );
        }

        #endregion

    }
}
