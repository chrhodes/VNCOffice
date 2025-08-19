using Microsoft.Office.Tools.Ribbon;

using VNCVisioToolsApplication.Actions;

namespace VNCVisioTools
{
    public partial class Ribbon
    {
        #region Event Handlers

        #region UI Launch Events

        //wrap all calls to UILaunchApproach in try/catch to prevent exceptions from crashing the add-in.
        //Use Common.WriteToDebugWindow(ex.Message, force:true) to handle exception

        private void btnThemedWindowHostModal_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                UILaunchApproach.ThemedWindowHostModal();
            }
            catch (System.Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnThemedWindowHostModeless_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                UILaunchApproach.ThemedWindowHostModeless();
            }
            catch (System.Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnDxWindowHost_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                UILaunchApproach.DxWindowHost();
            }
            catch (System.Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnWindowHostLocal_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                UILaunchApproach.WindowHostLocal();
            }
            catch (System.Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnWindowHostVNC_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                UILaunchApproach.WindowHostVNC();
            }
            catch (System.Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        //private static void ShowEmptyHost(Window host, string title, ShowWindowMode mode)
        //{
        //    //UILaunchApproach.Show
        //}

        #endregion UI Launch Events

        #region WPF UI Events

        private void btnLaunchCylon_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                UILaunchApproach.LaunchCylon();
            }
            catch (System.Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnLaunchCylon2_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                UILaunchApproach.LaunchCylon2();
            }
            catch (System.Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnDxLayoutControl_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                UILaunchApproach.DxLayoutControl();
            }
            catch (System.Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnDxDockLayoutControl_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                UILaunchApproach.DxDockLayoutControl();
            }
            catch (System.Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnDxDockLayoutManagerControl_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                UILaunchApproach.DxDockLayoutManagerControl();
            }
            catch (System.Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnPrismRegionTest_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                UILaunchApproach.PrismRegionTest();
            }
            catch (System.Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        #endregion WPF UI Events

        #region MVVM Examples

        private void btnVNC_MVVM_VAVM1st_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                UILaunchApproach.VNC_MVVM_VAVM1st();
            }
            catch (System.Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnVNC_MVVM_VA1st_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                UILaunchApproach.VNC_MVVM_VA1st();
            }
            catch (System.Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnVNC_MVVM_VAVM1stDI_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                UILaunchApproach.VNC_MVVM_VAVM1stDI();
            }
            catch (System.Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnVNC_MVVM_VB1st_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                UILaunchApproach.VNC_MVVM_VB1st();
            }
            catch (System.Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnVNC_MVVM_VC11st_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                UILaunchApproach.VNC_MVVM_VC11st();
            }
            catch (System.Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        private void btnVNC_MVVM_VC21st_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                UILaunchApproach.VNC_MVVM_VC21st();
            }
            catch (System.Exception ex)
            {
                Common.WriteToDebugWindow(ex.Message, force: true);
            }
        }

        #endregion

        #endregion Event Handlers
    }
}