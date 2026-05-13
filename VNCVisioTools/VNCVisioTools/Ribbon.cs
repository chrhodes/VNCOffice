using System;
using System.Threading;

using Microsoft.Office.Tools.Ribbon;

using VNC;

namespace VNCVisioTools
{
    public partial class Ribbon
    {
        // NOTE(crhodes)
        // This was moved out of designer so we can log

        public Ribbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
#if DEBUG
            Common.InitializeLogging(debugConfig: true);
#else
            Common.InitializeLogging();
#endif
            Int64 startTicks = 0;
            Common.WriteToDebugWindow("Ribbon()", true);
            if (Common.VNCLogging.ApplicationStart) startTicks = Log.APPLICATION_START("Initialize SignalR", Common.LOG_CATEGORY);

            // NOTE(crhodes)
            // If don't delay a bit here, the SignalR logging infrastructure does not initialize quickly enough
            // and the first few log messages are missed.
            // NB.  All are properly recored in the log file.

            Thread.Sleep(250);

            if (Common.VNCLogging.ApplicationStart) startTicks = Log.APPLICATION_START("Enter/Exit", Common.LOG_CATEGORY);
        }

        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            Int64 startTicks = 0;
            Common.WriteToDebugWindow("Ribbon_Load()", true);
            if (Common.VNCLogging.ApplicationStart) startTicks = Log.APPLICATION_START("Enter/Exit", Common.LOG_CATEGORY);
        }

        private void btnGetApplicationInfo_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnGetDocumentInfo_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnGetStencilInfo_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnAddTableOfContents_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnAddHeader_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnAddFooter_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnRemoveLayers_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnSortAllPages_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnDisplayPageNames_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnSyncPageNames_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnAutoSizePagesOn_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnAutoSizePagesOff_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnUpdatePageNameShapes_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnAddNavigationLinks_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnPrintPages_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnDeletePages_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnSavePages_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnAddArchitectureBasePages_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnAddBackgroundPages_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnAddDefaultLayers_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnGetPageInfo_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnUpdatePageNameShapesPage_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnAddNavLinks_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnPrintPage_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnSavePage_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnSyncPageNamesPage_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnAutoSizePageOn_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnAutoSizePageOff_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnPageOn_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnPageOff_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnAllPageOn_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnAllPageOff_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnLayerManager_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnLockBackground_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnUnlockBackground_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnAddDefaultLayers_Page_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnRemoveLayers_Page_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnGetShapeInfo_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnAddTextControl_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnAddIsPageName_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnAddHyperLink_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnAddColorSupport_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnMakeLinkableMaster_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnAddIDSupport_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnAddIDAndTextSupport_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnMoveToBackgroundLayer_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btn0PtMargins_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btn1PtMargins_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btn2PtMargins_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnCommandCockpit_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnEditControlRows_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnEditParagraph_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnEditText_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnEditControlPoints_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnRenamePages_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnDuplicatePage_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnMovePages_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnCustomUI_Car_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnDebugWindow_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnWatchWindow_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnDisplayAddInInfo_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnToggleDeveloperMode_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnThemedWindowHostModeless_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnThemedWindowHostModal_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnWindowHostLocal_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnWindowHostVNC_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnDxWindowHost_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnLaunchCylon_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnLaunchCylon2_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnDxDockLayoutManagerControl_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnDxLayoutControl_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnDxDockLayoutControl_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnPrismRegionTest_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnVNC_MVVM_VAVM1st_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnVNC_MVVM_VA1st_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnVNC_MVVM_VAVM1stDI_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnVNC_MVVM_VB1st_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnVNC_MVVM_VC11st_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnVNC_MVVM_VC21st_Click(object sender, RibbonControlEventArgs e)
        {

        }
    }
}
