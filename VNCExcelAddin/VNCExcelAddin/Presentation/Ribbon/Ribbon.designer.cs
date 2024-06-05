namespace VNCExcelAddin
{
    partial class Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        //public Ribbon()
        //    : base(Globals.Factory.GetRibbonFactory())
        //{
        //    InitializeComponent();
        //}

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if(disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl3 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl4 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl5 = this.Factory.CreateRibbonDropDownItem();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.tabSupportTools = this.Factory.CreateRibbonTab();
            this.grpTaskPanes = this.Factory.CreateRibbonGroup();
            this.btnAppUtilities = this.Factory.CreateRibbonButton();
            this.btnExcelUtilities = this.Factory.CreateRibbonButton();
            this.btnLogParser = this.Factory.CreateRibbonButton();
            this.btnNetworkTraces = this.Factory.CreateRibbonButton();
            this.btnActiveDirectory = this.Factory.CreateRibbonButton();
            this.btnExaVault = this.Factory.CreateRibbonButton();
            this.btnRally = this.Factory.CreateRibbonButton();
            this.btnSalesforce = this.Factory.CreateRibbonButton();
            this.btnTPDevelopment = this.Factory.CreateRibbonButton();
            this.grpForms = this.Factory.CreateRibbonGroup();
            this.btnLoadTPHost_ActiveDirectory = this.Factory.CreateRibbonButton();
            this.btnLoadTFSHost = this.Factory.CreateRibbonButton();
            this.btnExplore = this.Factory.CreateRibbonButton();
            this.grpHelp = this.Factory.CreateRibbonGroup();
            this.btnAddInInfo = this.Factory.CreateRibbonButton();
            this.btnDeveloperMode = this.Factory.CreateRibbonButton();
            this.ddTheme = this.Factory.CreateRibbonDropDown();
            this.grpDebug = this.Factory.CreateRibbonGroup();
            this.btnDebugWindow = this.Factory.CreateRibbonButton();
            this.btnWatchWindow = this.Factory.CreateRibbonButton();
            this.chkEnableAppEvents = this.Factory.CreateRibbonCheckBox();
            this.chkDisplayEvents = this.Factory.CreateRibbonCheckBox();
            this.chkScreenUpdates = this.Factory.CreateRibbonCheckBox();
            this.chkDisplayXlLocationUpdates = this.Factory.CreateRibbonCheckBox();
            this.chkEnableTraceLogging = this.Factory.CreateRibbonCheckBox();
            this.tabUILaunch = this.Factory.CreateRibbonTab();
            this.grpUILaunch = this.Factory.CreateRibbonGroup();
            this.btnThemedWindowHostModeless = this.Factory.CreateRibbonButton();
            this.btnThemedWIndowHostModal = this.Factory.CreateRibbonButton();
            this.btnWindowHostLocal = this.Factory.CreateRibbonButton();
            this.btnWindowHostVNC = this.Factory.CreateRibbonButton();
            this.btnDxWindowHost = this.Factory.CreateRibbonButton();
            this.button9 = this.Factory.CreateRibbonButton();
            this.button10 = this.Factory.CreateRibbonButton();
            this.button11 = this.Factory.CreateRibbonButton();
            this.grpWPFUI = this.Factory.CreateRibbonGroup();
            this.btnLaunchCylon = this.Factory.CreateRibbonButton();
            this.btnLaunchCylonn2 = this.Factory.CreateRibbonButton();
            this.btnPrismRegionTest = this.Factory.CreateRibbonButton();
            this.btnDxLayoutControl = this.Factory.CreateRibbonButton();
            this.btnDxDockLayoutControl = this.Factory.CreateRibbonButton();
            this.btnDockLayoutManagerControl = this.Factory.CreateRibbonButton();
            this.button6 = this.Factory.CreateRibbonButton();
            this.button7 = this.Factory.CreateRibbonButton();
            this.button8 = this.Factory.CreateRibbonButton();
            this.rgMVVMExamples = this.Factory.CreateRibbonGroup();
            this.btnVNC_MVVM_VAVM1st = this.Factory.CreateRibbonButton();
            this.btnVNC_MVVM_VA1st = this.Factory.CreateRibbonButton();
            this.btnVNC_MVVM_VAVM1stDI = this.Factory.CreateRibbonButton();
            this.btnVNC_MVVM_VB1st = this.Factory.CreateRibbonButton();
            this.btnVNC_MVVM_VC11st = this.Factory.CreateRibbonButton();
            this.btnVNC_MVVM_VC21st = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.tabSupportTools.SuspendLayout();
            this.grpTaskPanes.SuspendLayout();
            this.grpForms.SuspendLayout();
            this.grpHelp.SuspendLayout();
            this.grpDebug.SuspendLayout();
            this.tabUILaunch.SuspendLayout();
            this.grpUILaunch.SuspendLayout();
            this.grpWPFUI.SuspendLayout();
            this.rgMVVMExamples.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // tabSupportTools
            // 
            this.tabSupportTools.Groups.Add(this.grpTaskPanes);
            this.tabSupportTools.Groups.Add(this.grpForms);
            this.tabSupportTools.Groups.Add(this.grpHelp);
            this.tabSupportTools.Groups.Add(this.grpDebug);
            this.tabSupportTools.Label = "Support Tools";
            this.tabSupportTools.Name = "tabSupportTools";
            // 
            // grpTaskPanes
            // 
            this.grpTaskPanes.Items.Add(this.btnAppUtilities);
            this.grpTaskPanes.Items.Add(this.btnExcelUtilities);
            this.grpTaskPanes.Items.Add(this.btnLogParser);
            this.grpTaskPanes.Items.Add(this.btnNetworkTraces);
            this.grpTaskPanes.Items.Add(this.btnActiveDirectory);
            this.grpTaskPanes.Items.Add(this.btnExaVault);
            this.grpTaskPanes.Items.Add(this.btnRally);
            this.grpTaskPanes.Items.Add(this.btnSalesforce);
            this.grpTaskPanes.Items.Add(this.btnTPDevelopment);
            this.grpTaskPanes.Label = "TaskPane Host";
            this.grpTaskPanes.Name = "grpTaskPanes";
            // 
            // btnAppUtilities
            // 
            this.btnAppUtilities.Label = "Excel Utilities";
            this.btnAppUtilities.Name = "btnAppUtilities";
            this.btnAppUtilities.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAppUtilities_Click);
            // 
            // btnExcelUtilities
            // 
            this.btnExcelUtilities.Label = "WPF Excel Utilities";
            this.btnExcelUtilities.Name = "btnExcelUtilities";
            this.btnExcelUtilities.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnExcelUtilities_Click);
            // 
            // btnLogParser
            // 
            this.btnLogParser.Label = "Log Parser";
            this.btnLogParser.Name = "btnLogParser";
            this.btnLogParser.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLogParser_Click);
            // 
            // btnNetworkTraces
            // 
            this.btnNetworkTraces.Label = "Network Traces";
            this.btnNetworkTraces.Name = "btnNetworkTraces";
            this.btnNetworkTraces.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnNetworkTraces_Click);
            // 
            // btnActiveDirectory
            // 
            this.btnActiveDirectory.Label = "Active Directory";
            this.btnActiveDirectory.Name = "btnActiveDirectory";
            this.btnActiveDirectory.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnActiveDirectory_Click);
            // 
            // btnExaVault
            // 
            this.btnExaVault.Label = "";
            this.btnExaVault.Name = "btnExaVault";
            // 
            // btnRally
            // 
            this.btnRally.Label = "";
            this.btnRally.Name = "btnRally";
            // 
            // btnSalesforce
            // 
            this.btnSalesforce.Label = "";
            this.btnSalesforce.Name = "btnSalesforce";
            // 
            // btnTPDevelopment
            // 
            this.btnTPDevelopment.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnTPDevelopment.Image = global::VNCExcelAddin.Properties.Resources.development_tools;
            this.btnTPDevelopment.Label = "Development";
            this.btnTPDevelopment.Name = "btnTPDevelopment";
            this.btnTPDevelopment.ShowImage = true;
            this.btnTPDevelopment.SuperTip = "Developer Tools";
            this.btnTPDevelopment.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnTPDevelopment_Click);
            // 
            // grpForms
            // 
            this.grpForms.Items.Add(this.btnLoadTPHost_ActiveDirectory);
            this.grpForms.Items.Add(this.btnLoadTFSHost);
            this.grpForms.Items.Add(this.btnExplore);
            this.grpForms.Label = "Winform Host";
            this.grpForms.Name = "grpForms";
            // 
            // btnLoadTPHost_ActiveDirectory
            // 
            this.btnLoadTPHost_ActiveDirectory.Label = "Active Directory";
            this.btnLoadTPHost_ActiveDirectory.Name = "btnLoadTPHost_ActiveDirectory";
            this.btnLoadTPHost_ActiveDirectory.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLoadTPHost_ActiveDirectory_Click);
            // 
            // btnLoadTFSHost
            // 
            this.btnLoadTFSHost.Label = "AZDO(TFS)";
            this.btnLoadTFSHost.Name = "btnLoadTFSHost";
            this.btnLoadTFSHost.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLoadTFSHost_Click);
            // 
            // btnExplore
            // 
            this.btnExplore.Label = "Explore";
            this.btnExplore.Name = "btnExplore";
            this.btnExplore.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnExplore_Click);
            // 
            // grpHelp
            // 
            this.grpHelp.Items.Add(this.btnAddInInfo);
            this.grpHelp.Items.Add(this.btnDeveloperMode);
            this.grpHelp.Items.Add(this.ddTheme);
            this.grpHelp.Label = "Help";
            this.grpHelp.Name = "grpHelp";
            // 
            // btnAddInInfo
            // 
            this.btnAddInInfo.Label = "AddIn Info";
            this.btnAddInInfo.Name = "btnAddInInfo";
            this.btnAddInInfo.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAddInInfo_Click);
            // 
            // btnDeveloperMode
            // 
            this.btnDeveloperMode.Label = "Developer Mode";
            this.btnDeveloperMode.Name = "btnDeveloperMode";
            this.btnDeveloperMode.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDeveloperMode_Click);
            // 
            // ddTheme
            // 
            ribbonDropDownItemImpl1.Label = "DeepBlue";
            ribbonDropDownItemImpl2.Label = "DXStyle";
            ribbonDropDownItemImpl3.Label = "LightGray";
            ribbonDropDownItemImpl4.Label = "MetropolisDark";
            ribbonDropDownItemImpl5.Label = "MetropolisLight";
            this.ddTheme.Items.Add(ribbonDropDownItemImpl1);
            this.ddTheme.Items.Add(ribbonDropDownItemImpl2);
            this.ddTheme.Items.Add(ribbonDropDownItemImpl3);
            this.ddTheme.Items.Add(ribbonDropDownItemImpl4);
            this.ddTheme.Items.Add(ribbonDropDownItemImpl5);
            this.ddTheme.Label = "Theme";
            this.ddTheme.Name = "ddTheme";
            this.ddTheme.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ddTheme_SelectionChanged);
            // 
            // grpDebug
            // 
            this.grpDebug.Items.Add(this.btnDebugWindow);
            this.grpDebug.Items.Add(this.btnWatchWindow);
            this.grpDebug.Items.Add(this.chkEnableAppEvents);
            this.grpDebug.Items.Add(this.chkDisplayEvents);
            this.grpDebug.Items.Add(this.chkScreenUpdates);
            this.grpDebug.Items.Add(this.chkDisplayXlLocationUpdates);
            this.grpDebug.Items.Add(this.chkEnableTraceLogging);
            this.grpDebug.Label = "Debug";
            this.grpDebug.Name = "grpDebug";
            this.grpDebug.Visible = false;
            // 
            // btnDebugWindow
            // 
            this.btnDebugWindow.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnDebugWindow.Image = global::VNCExcelAddin.Properties.Resources.Auto_Debug_System_icon;
            this.btnDebugWindow.Label = "Debug Window";
            this.btnDebugWindow.Name = "btnDebugWindow";
            this.btnDebugWindow.ShowImage = true;
            this.btnDebugWindow.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDebugWindow_Click);
            // 
            // btnWatchWindow
            // 
            this.btnWatchWindow.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnWatchWindow.Image = global::VNCExcelAddin.Properties.Resources.WatchWindow;
            this.btnWatchWindow.Label = "Watch Window";
            this.btnWatchWindow.Name = "btnWatchWindow";
            this.btnWatchWindow.ShowImage = true;
            this.btnWatchWindow.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnWatchWindow_Click);
            // 
            // chkEnableAppEvents
            // 
            this.chkEnableAppEvents.Checked = true;
            this.chkEnableAppEvents.Label = "Enable App Events";
            this.chkEnableAppEvents.Name = "chkEnableAppEvents";
            this.chkEnableAppEvents.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.chkEnableAppEvents_Click);
            // 
            // chkDisplayEvents
            // 
            this.chkDisplayEvents.Label = "Display Events";
            this.chkDisplayEvents.Name = "chkDisplayEvents";
            this.chkDisplayEvents.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.chkDisplayEvents_Click);
            // 
            // chkScreenUpdates
            // 
            this.chkScreenUpdates.Label = "Display Screen Updates";
            this.chkScreenUpdates.Name = "chkScreenUpdates";
            this.chkScreenUpdates.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.chkScreenUpdates_Click);
            // 
            // chkDisplayXlLocationUpdates
            // 
            this.chkDisplayXlLocationUpdates.Label = "Display XlLocation Updates";
            this.chkDisplayXlLocationUpdates.Name = "chkDisplayXlLocationUpdates";
            this.chkDisplayXlLocationUpdates.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.chkDisplayXlLocationUpdates_Click);
            // 
            // chkEnableTraceLogging
            // 
            this.chkEnableTraceLogging.Label = "Enable Trace Logging";
            this.chkEnableTraceLogging.Name = "chkEnableTraceLogging";
            this.chkEnableTraceLogging.SuperTip = "Adds Log.Trace call to all writes to WatchWindow";
            this.chkEnableTraceLogging.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.chkEnableTraceLogging_Click);
            // 
            // tabUILaunch
            // 
            this.tabUILaunch.Groups.Add(this.grpUILaunch);
            this.tabUILaunch.Groups.Add(this.grpWPFUI);
            this.tabUILaunch.Groups.Add(this.rgMVVMExamples);
            this.tabUILaunch.Label = "UI Launch Approches";
            this.tabUILaunch.Name = "tabUILaunch";
            // 
            // grpUILaunch
            // 
            this.grpUILaunch.Items.Add(this.btnThemedWindowHostModeless);
            this.grpUILaunch.Items.Add(this.btnThemedWIndowHostModal);
            this.grpUILaunch.Items.Add(this.btnWindowHostLocal);
            this.grpUILaunch.Items.Add(this.btnWindowHostVNC);
            this.grpUILaunch.Items.Add(this.btnDxWindowHost);
            this.grpUILaunch.Items.Add(this.button9);
            this.grpUILaunch.Items.Add(this.button10);
            this.grpUILaunch.Items.Add(this.button11);
            this.grpUILaunch.Label = "Hosts";
            this.grpUILaunch.Name = "grpUILaunch";
            // 
            // btnThemedWindowHostModeless
            // 
            this.btnThemedWindowHostModeless.Label = "ThemedWindow Host (Modeless)";
            this.btnThemedWindowHostModeless.Name = "btnThemedWindowHostModeless";
            this.btnThemedWindowHostModeless.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnThemedWindowHostModeless_Click);
            // 
            // btnThemedWIndowHostModal
            // 
            this.btnThemedWIndowHostModal.Label = "ThemedWindow Host (Modal)";
            this.btnThemedWIndowHostModal.Name = "btnThemedWIndowHostModal";
            this.btnThemedWIndowHostModal.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnThemedWindowHostModal_Click);
            // 
            // btnWindowHostLocal
            // 
            this.btnWindowHostLocal.Label = "WindowHost (Local)";
            this.btnWindowHostLocal.Name = "btnWindowHostLocal";
            this.btnWindowHostLocal.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnWindowHostLocal_Click);
            // 
            // btnWindowHostVNC
            // 
            this.btnWindowHostVNC.Label = "WIndowHost (VNC)";
            this.btnWindowHostVNC.Name = "btnWindowHostVNC";
            this.btnWindowHostVNC.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnWindowHostVNC_Click);
            // 
            // btnDxWindowHost
            // 
            this.btnDxWindowHost.Label = "DxWindow Host";
            this.btnDxWindowHost.Name = "btnDxWindowHost";
            this.btnDxWindowHost.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDxWindowHost_Click);
            // 
            // button9
            // 
            this.button9.Label = "";
            this.button9.Name = "button9";
            // 
            // button10
            // 
            this.button10.Label = "";
            this.button10.Name = "button10";
            // 
            // button11
            // 
            this.button11.Label = "";
            this.button11.Name = "button11";
            // 
            // grpWPFUI
            // 
            this.grpWPFUI.Items.Add(this.btnLaunchCylon);
            this.grpWPFUI.Items.Add(this.btnLaunchCylonn2);
            this.grpWPFUI.Items.Add(this.btnPrismRegionTest);
            this.grpWPFUI.Items.Add(this.btnDxLayoutControl);
            this.grpWPFUI.Items.Add(this.btnDxDockLayoutControl);
            this.grpWPFUI.Items.Add(this.btnDockLayoutManagerControl);
            this.grpWPFUI.Items.Add(this.button6);
            this.grpWPFUI.Items.Add(this.button7);
            this.grpWPFUI.Items.Add(this.button8);
            this.grpWPFUI.Label = "WPF UI";
            this.grpWPFUI.Name = "grpWPFUI";
            // 
            // btnLaunchCylon
            // 
            this.btnLaunchCylon.Label = "Launch Cylon";
            this.btnLaunchCylon.Name = "btnLaunchCylon";
            this.btnLaunchCylon.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLaunchCylon_Click);
            // 
            // btnLaunchCylonn2
            // 
            this.btnLaunchCylonn2.Label = "Launch Cylon 2";
            this.btnLaunchCylonn2.Name = "btnLaunchCylonn2";
            this.btnLaunchCylonn2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLaunchCylon2_Click);
            // 
            // btnPrismRegionTest
            // 
            this.btnPrismRegionTest.Label = "Prism Region Test";
            this.btnPrismRegionTest.Name = "btnPrismRegionTest";
            this.btnPrismRegionTest.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnPrismRegionTest_Click);
            // 
            // btnDxLayoutControl
            // 
            this.btnDxLayoutControl.Label = "DxLayoutControl";
            this.btnDxLayoutControl.Name = "btnDxLayoutControl";
            this.btnDxLayoutControl.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDxLayoutControl_Click);
            // 
            // btnDxDockLayoutControl
            // 
            this.btnDxDockLayoutControl.Label = "DxDockLayoutControl";
            this.btnDxDockLayoutControl.Name = "btnDxDockLayoutControl";
            this.btnDxDockLayoutControl.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDxDockLayoutControl_Click);
            // 
            // btnDockLayoutManagerControl
            // 
            this.btnDockLayoutManagerControl.Label = "DxDockLayoutManagerControl";
            this.btnDockLayoutManagerControl.Name = "btnDockLayoutManagerControl";
            this.btnDockLayoutManagerControl.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDxDockLayoutManagerControl_Click);
            // 
            // button6
            // 
            this.button6.Label = "";
            this.button6.Name = "button6";
            // 
            // button7
            // 
            this.button7.Label = "";
            this.button7.Name = "button7";
            // 
            // button8
            // 
            this.button8.Label = "";
            this.button8.Name = "button8";
            // 
            // rgMVVMExamples
            // 
            this.rgMVVMExamples.Items.Add(this.btnVNC_MVVM_VAVM1st);
            this.rgMVVMExamples.Items.Add(this.btnVNC_MVVM_VA1st);
            this.rgMVVMExamples.Items.Add(this.btnVNC_MVVM_VAVM1stDI);
            this.rgMVVMExamples.Items.Add(this.btnVNC_MVVM_VB1st);
            this.rgMVVMExamples.Items.Add(this.btnVNC_MVVM_VC11st);
            this.rgMVVMExamples.Items.Add(this.btnVNC_MVVM_VC21st);
            this.rgMVVMExamples.Label = "MVVM Examples";
            this.rgMVVMExamples.Name = "rgMVVMExamples";
            // 
            // btnVNC_MVVM_VAVM1st
            // 
            this.btnVNC_MVVM_VAVM1st.Label = "VNC MVVM VAVM 1st";
            this.btnVNC_MVVM_VAVM1st.Name = "btnVNC_MVVM_VAVM1st";
            this.btnVNC_MVVM_VAVM1st.SuperTip = "ViewModel First by Hand";
            this.btnVNC_MVVM_VAVM1st.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnVNC_MVVM_VAVM1st_Click);
            // 
            // btnVNC_MVVM_VA1st
            // 
            this.btnVNC_MVVM_VA1st.Label = "VNC MVVM VA1st (DI)";
            this.btnVNC_MVVM_VA1st.Name = "btnVNC_MVVM_VA1st";
            this.btnVNC_MVVM_VA1st.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnVNC_MVVM_VA1st_Click);
            // 
            // btnVNC_MVVM_VAVM1stDI
            // 
            this.btnVNC_MVVM_VAVM1stDI.Label = "VNC MVVM VAVM 1st (DI)";
            this.btnVNC_MVVM_VAVM1stDI.Name = "btnVNC_MVVM_VAVM1stDI";
            this.btnVNC_MVVM_VAVM1stDI.SuperTip = "ViewAViewModel 1st using DI";
            this.btnVNC_MVVM_VAVM1stDI.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnVNC_MVVM_VAVM1stDI_Click);
            // 
            // btnVNC_MVVM_VB1st
            // 
            this.btnVNC_MVVM_VB1st.Label = "VNC MVVM VB 1st (DI)";
            this.btnVNC_MVVM_VB1st.Name = "btnVNC_MVVM_VB1st";
            this.btnVNC_MVVM_VB1st.SuperTip = "ViewB has a parameterless constructor and one that takes a ViewModel and ViewMode" +
    "l is registed with DI";
            this.btnVNC_MVVM_VB1st.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnVNC_MVVM_VB1st_Click);
            // 
            // btnVNC_MVVM_VC11st
            // 
            this.btnVNC_MVVM_VC11st.Label = "VNC MVVM VC1 1st (DI)";
            this.btnVNC_MVVM_VC11st.Name = "btnVNC_MVVM_VC11st";
            this.btnVNC_MVVM_VC11st.SuperTip = "ViewC has parameterless and parameterized(ViewModel) constructors and is not regi" +
    "stered with DI";
            this.btnVNC_MVVM_VC11st.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnVNC_MVVM_VC11st_Click);
            // 
            // btnVNC_MVVM_VC21st
            // 
            this.btnVNC_MVVM_VC21st.Label = "VNC MVVM VC2 1st (DI)";
            this.btnVNC_MVVM_VC21st.Name = "btnVNC_MVVM_VC21st";
            this.btnVNC_MVVM_VC21st.SuperTip = "ViewC2 has no parameterless constructor and is not Registered with DI";
            this.btnVNC_MVVM_VC21st.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnVNC_MVVM_VC21st_Click);
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Tabs.Add(this.tabSupportTools);
            this.Tabs.Add(this.tabUILaunch);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.tabSupportTools.ResumeLayout(false);
            this.tabSupportTools.PerformLayout();
            this.grpTaskPanes.ResumeLayout(false);
            this.grpTaskPanes.PerformLayout();
            this.grpForms.ResumeLayout(false);
            this.grpForms.PerformLayout();
            this.grpHelp.ResumeLayout(false);
            this.grpHelp.PerformLayout();
            this.grpDebug.ResumeLayout(false);
            this.grpDebug.PerformLayout();
            this.tabUILaunch.ResumeLayout(false);
            this.tabUILaunch.PerformLayout();
            this.grpUILaunch.ResumeLayout(false);
            this.grpUILaunch.PerformLayout();
            this.grpWPFUI.ResumeLayout(false);
            this.grpWPFUI.PerformLayout();
            this.rgMVVMExamples.ResumeLayout(false);
            this.rgMVVMExamples.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        private Microsoft.Office.Tools.Ribbon.RibbonTab tabSupportTools;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpTaskPanes;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpDebug;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDebugWindow;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnWatchWindow;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox chkEnableAppEvents;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox chkDisplayEvents;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpHelp;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAddInInfo;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDeveloperMode;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAppUtilities;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnNetworkTraces;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLogParser;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox chkScreenUpdates;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnActiveDirectory;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRally;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSalesforce;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnExaVault;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpForms;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLoadTFSHost;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown ddTheme;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnTPDevelopment;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnExcelUtilities;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox chkDisplayXlLocationUpdates;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLoadTPHost_ActiveDirectory;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnExplore;
        private Microsoft.Office.Tools.Ribbon.RibbonTab tabUILaunch;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpUILaunch;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnThemedWindowHostModeless;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnThemedWIndowHostModal;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnWindowHostLocal;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnWindowHostVNC;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDxWindowHost;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button9;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button10;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button11;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpWPFUI;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLaunchCylon;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLaunchCylonn2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPrismRegionTest;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDxLayoutControl;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDxDockLayoutControl;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDockLayoutManagerControl;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button6;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button7;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button8;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox chkEnableTraceLogging;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup rgMVVMExamples;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnVNC_MVVM_VAVM1st;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnVNC_MVVM_VA1st;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnVNC_MVVM_VAVM1stDI;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnVNC_MVVM_VB1st;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnVNC_MVVM_VC11st;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnVNC_MVVM_VC21st;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon
        {
            get
            {
                return this.GetRibbon<Ribbon>();
            }
        }
    }
}
