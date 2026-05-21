namespace VNCExcelTools
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
            if (disposing && (components != null))
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.rtVNCExcelTools = this.Factory.CreateRibbonTab();
            this.rgDocumentActions = this.Factory.CreateRibbonGroup();
            this.btnAddTableOfContents = this.Factory.CreateRibbonButton();
            this.btnAddHeader = this.Factory.CreateRibbonButton();
            this.btnAddFooter = this.Factory.CreateRibbonButton();
            this.rgUtilities = this.Factory.CreateRibbonGroup();
            this.btnFolderMap = this.Factory.CreateRibbonButton();
            this.rgDebug = this.Factory.CreateRibbonGroup();
            this.btnDebugWindow = this.Factory.CreateRibbonButton();
            this.btnWatchWindow = this.Factory.CreateRibbonButton();
            this.btnLoggingConfiguration = this.Factory.CreateRibbonButton();
            this.rcbEnableAppEvents = this.Factory.CreateRibbonCheckBox();
            this.rcbDisplayEvents = this.Factory.CreateRibbonCheckBox();
            this.rcbDisplayChattyEvents = this.Factory.CreateRibbonCheckBox();
            this.rcbToggleDeveloperUIMode = this.Factory.CreateRibbonCheckBox();
            this.btnTestExcelLogging = this.Factory.CreateRibbonButton();
            this.rcbUILaunchApproaches = this.Factory.CreateRibbonCheckBox();
            this.grpHelp = this.Factory.CreateRibbonGroup();
            this.btnDisplayAddInInfo = this.Factory.CreateRibbonButton();
            this.btnToggleDeveloperMode = this.Factory.CreateRibbonButton();
            this.rtUILaunchApproaches = this.Factory.CreateRibbonTab();
            this.rgUILaunch = this.Factory.CreateRibbonGroup();
            this.btnThemedWindowHostModeless = this.Factory.CreateRibbonButton();
            this.btnThemedWindowHostModal = this.Factory.CreateRibbonButton();
            this.btnWindowHostLocal = this.Factory.CreateRibbonButton();
            this.btnWindowHostVNC = this.Factory.CreateRibbonButton();
            this.btnDxWindowHost = this.Factory.CreateRibbonButton();
            this.rgWPFUI = this.Factory.CreateRibbonGroup();
            this.btnLaunchCylon = this.Factory.CreateRibbonButton();
            this.btnLaunchCylon2 = this.Factory.CreateRibbonButton();
            this.btnDxDockLayoutManagerControl = this.Factory.CreateRibbonButton();
            this.btnDxLayoutControl = this.Factory.CreateRibbonButton();
            this.btnDxDockLayoutControl = this.Factory.CreateRibbonButton();
            this.btnPrismRegionTest = this.Factory.CreateRibbonButton();
            this.rgMVVMExamples = this.Factory.CreateRibbonGroup();
            this.btnVNC_MVVM_VAVM1st = this.Factory.CreateRibbonButton();
            this.btnVNC_MVVM_VA1st = this.Factory.CreateRibbonButton();
            this.btnVNC_MVVM_VAVM1stDI = this.Factory.CreateRibbonButton();
            this.btnVNC_MVVM_VB1st = this.Factory.CreateRibbonButton();
            this.btnVNC_MVVM_VC11st = this.Factory.CreateRibbonButton();
            this.btnVNC_MVVM_VC21st = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.rtVNCExcelTools.SuspendLayout();
            this.rgDocumentActions.SuspendLayout();
            this.rgUtilities.SuspendLayout();
            this.rgDebug.SuspendLayout();
            this.grpHelp.SuspendLayout();
            this.rtUILaunchApproaches.SuspendLayout();
            this.rgUILaunch.SuspendLayout();
            this.rgWPFUI.SuspendLayout();
            this.rgMVVMExamples.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Label = "group1";
            this.group1.Name = "group1";
            // 
            // rtVNCExcelTools
            // 
            this.rtVNCExcelTools.Groups.Add(this.rgDocumentActions);
            this.rtVNCExcelTools.Groups.Add(this.rgUtilities);
            this.rtVNCExcelTools.Groups.Add(this.rgDebug);
            this.rtVNCExcelTools.Groups.Add(this.grpHelp);
            this.rtVNCExcelTools.Label = "VNCExcelTools";
            this.rtVNCExcelTools.Name = "rtVNCExcelTools";
            // 
            // rgDocumentActions
            // 
            this.rgDocumentActions.Items.Add(this.btnAddTableOfContents);
            this.rgDocumentActions.Items.Add(this.btnAddHeader);
            this.rgDocumentActions.Items.Add(this.btnAddFooter);
            this.rgDocumentActions.Label = "Document Actions";
            this.rgDocumentActions.Name = "rgDocumentActions";
            // 
            // btnAddTableOfContents
            // 
            this.btnAddTableOfContents.Label = "+ TOC";
            this.btnAddTableOfContents.Name = "btnAddTableOfContents";
            this.btnAddTableOfContents.ScreenTip = "Add Table of Contents Page";
            this.btnAddTableOfContents.SuperTip = "Add Table of Contents Page containing link shapes to all pages";
            this.btnAddTableOfContents.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAddTableOfContents_Click);
            // 
            // btnAddHeader
            // 
            this.btnAddHeader.Label = "+ Header";
            this.btnAddHeader.Name = "btnAddHeader";
            this.btnAddHeader.ScreenTip = "Add Header to all Pages";
            this.btnAddHeader.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAddHeader_Click);
            // 
            // btnAddFooter
            // 
            this.btnAddFooter.Label = "+ Footer";
            this.btnAddFooter.Name = "btnAddFooter";
            this.btnAddFooter.ScreenTip = "Add Footer to all Pages";
            this.btnAddFooter.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAddFooter_Click);
            // 
            // rgUtilities
            // 
            this.rgUtilities.Items.Add(this.btnFolderMap);
            this.rgUtilities.Label = "Utilities";
            this.rgUtilities.Name = "rgUtilities";
            // 
            // btnFolderMap
            // 
            this.btnFolderMap.Label = "Folder Map";
            this.btnFolderMap.Name = "btnFolderMap";
            this.btnFolderMap.SuperTip = "Create File/Folder Map";
            this.btnFolderMap.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFolderMap_Click);
            // 
            // rgDebug
            // 
            this.rgDebug.Items.Add(this.btnDebugWindow);
            this.rgDebug.Items.Add(this.btnWatchWindow);
            this.rgDebug.Items.Add(this.btnLoggingConfiguration);
            this.rgDebug.Items.Add(this.rcbEnableAppEvents);
            this.rgDebug.Items.Add(this.rcbDisplayEvents);
            this.rgDebug.Items.Add(this.rcbDisplayChattyEvents);
            this.rgDebug.Items.Add(this.rcbToggleDeveloperUIMode);
            this.rgDebug.Items.Add(this.btnTestExcelLogging);
            this.rgDebug.Items.Add(this.rcbUILaunchApproaches);
            this.rgDebug.Label = "Debug";
            this.rgDebug.Name = "rgDebug";
            this.rgDebug.Visible = false;
            // 
            // btnDebugWindow
            // 
            this.btnDebugWindow.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnDebugWindow.Image = ((System.Drawing.Image)(resources.GetObject("btnDebugWindow.Image")));
            this.btnDebugWindow.Label = "Debug Window";
            this.btnDebugWindow.Name = "btnDebugWindow";
            this.btnDebugWindow.ShowImage = true;
            this.btnDebugWindow.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDebugWindow_Click);
            // 
            // btnWatchWindow
            // 
            this.btnWatchWindow.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnWatchWindow.Image = ((System.Drawing.Image)(resources.GetObject("btnWatchWindow.Image")));
            this.btnWatchWindow.Label = "Watch Window";
            this.btnWatchWindow.Name = "btnWatchWindow";
            this.btnWatchWindow.ShowImage = true;
            this.btnWatchWindow.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnWatchWindow_Click);
            // 
            // btnLoggingConfiguration
            // 
            this.btnLoggingConfiguration.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnLoggingConfiguration.Image = ((System.Drawing.Image)(resources.GetObject("btnLoggingConfiguration.Image")));
            this.btnLoggingConfiguration.Label = "Logging Configuration";
            this.btnLoggingConfiguration.Name = "btnLoggingConfiguration";
            this.btnLoggingConfiguration.ShowImage = true;
            this.btnLoggingConfiguration.SuperTip = "Configure VNC Logging Levels";
            this.btnLoggingConfiguration.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLoggingConfiguration_Click);
            // 
            // rcbEnableAppEvents
            // 
            this.rcbEnableAppEvents.Label = "Enable App Events";
            this.rcbEnableAppEvents.Name = "rcbEnableAppEvents";
            this.rcbEnableAppEvents.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.rcbEnableAppEvents_Click);
            // 
            // rcbDisplayEvents
            // 
            this.rcbDisplayEvents.Label = "Display Events";
            this.rcbDisplayEvents.Name = "rcbDisplayEvents";
            this.rcbDisplayEvents.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.rcbDisplayEvents_Click);
            // 
            // rcbDisplayChattyEvents
            // 
            this.rcbDisplayChattyEvents.Label = "Display Chatty Events";
            this.rcbDisplayChattyEvents.Name = "rcbDisplayChattyEvents";
            this.rcbDisplayChattyEvents.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.rcbDisplayChattyEvents_Click);
            // 
            // rcbToggleDeveloperUIMode
            // 
            this.rcbToggleDeveloperUIMode.Label = "DeveloperUIMode";
            this.rcbToggleDeveloperUIMode.Name = "rcbToggleDeveloperUIMode";
            this.rcbToggleDeveloperUIMode.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.rcbToggleDeveloperUIMode_Click);
            // 
            // btnTestExcelLogging
            // 
            this.btnTestExcelLogging.Label = "Test Excel Logging";
            this.btnTestExcelLogging.Name = "btnTestExcelLogging";
            this.btnTestExcelLogging.SuperTip = "Test Debug and Watch Logging";
            this.btnTestExcelLogging.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnTestExcelLogging_Click);
            // 
            // rcbUILaunchApproaches
            // 
            this.rcbUILaunchApproaches.Label = "UILaunchApproaches";
            this.rcbUILaunchApproaches.Name = "rcbUILaunchApproaches";
            this.rcbUILaunchApproaches.SuperTip = "Display UILaunchApproaches Ribbon Group";
            this.rcbUILaunchApproaches.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.rcbUILaunchApproaches_Click);
            // 
            // grpHelp
            // 
            this.grpHelp.Items.Add(this.btnDisplayAddInInfo);
            this.grpHelp.Items.Add(this.btnToggleDeveloperMode);
            this.grpHelp.Label = "Help";
            this.grpHelp.Name = "grpHelp";
            // 
            // btnDisplayAddInInfo
            // 
            this.btnDisplayAddInInfo.Label = "AddIn Info";
            this.btnDisplayAddInInfo.Name = "btnDisplayAddInInfo";
            this.btnDisplayAddInInfo.SuperTip = "Display AddIn Information";
            this.btnDisplayAddInInfo.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDisplayAddInInfo_Click);
            // 
            // btnToggleDeveloperMode
            // 
            this.btnToggleDeveloperMode.Label = "Developer Mode";
            this.btnToggleDeveloperMode.Name = "btnToggleDeveloperMode";
            this.btnToggleDeveloperMode.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnToggleDeveloperMode_Click);
            // 
            // rtUILaunchApproaches
            // 
            this.rtUILaunchApproaches.Groups.Add(this.rgUILaunch);
            this.rtUILaunchApproaches.Groups.Add(this.rgWPFUI);
            this.rtUILaunchApproaches.Groups.Add(this.rgMVVMExamples);
            this.rtUILaunchApproaches.Label = "UI Launch Approaches";
            this.rtUILaunchApproaches.Name = "rtUILaunchApproaches";
            // 
            // rgUILaunch
            // 
            this.rgUILaunch.Items.Add(this.btnThemedWindowHostModeless);
            this.rgUILaunch.Items.Add(this.btnThemedWindowHostModal);
            this.rgUILaunch.Items.Add(this.btnWindowHostLocal);
            this.rgUILaunch.Items.Add(this.btnWindowHostVNC);
            this.rgUILaunch.Items.Add(this.btnDxWindowHost);
            this.rgUILaunch.Label = "UI Launch";
            this.rgUILaunch.Name = "rgUILaunch";
            // 
            // btnThemedWindowHostModeless
            // 
            this.btnThemedWindowHostModeless.Label = "ThemedWindow Host (Modeless)";
            this.btnThemedWindowHostModeless.Name = "btnThemedWindowHostModeless";
            this.btnThemedWindowHostModeless.ScreenTip = "dx:ThemedWindow (Show)";
            this.btnThemedWindowHostModeless.SuperTip = "Super TIp";
            this.btnThemedWindowHostModeless.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnThemedWindowHostModeless_Click);
            // 
            // btnThemedWindowHostModal
            // 
            this.btnThemedWindowHostModal.Label = "ThemedWindow Host (Modal)";
            this.btnThemedWindowHostModal.Name = "btnThemedWindowHostModal";
            this.btnThemedWindowHostModal.ScreenTip = "dx:ThemedWindow (ShowDialog)";
            this.btnThemedWindowHostModal.SuperTip = "Super TIp";
            this.btnThemedWindowHostModal.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnThemedWindowHostModal_Click);
            // 
            // btnWindowHostLocal
            // 
            this.btnWindowHostLocal.Label = "Window Host (Local)";
            this.btnWindowHostLocal.Name = "btnWindowHostLocal";
            this.btnWindowHostLocal.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnWindowHostLocal_Click);
            // 
            // btnWindowHostVNC
            // 
            this.btnWindowHostVNC.Label = "Window Host (VNC)";
            this.btnWindowHostVNC.Name = "btnWindowHostVNC";
            this.btnWindowHostVNC.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnWindowHostVNC_Click);
            // 
            // btnDxWindowHost
            // 
            this.btnDxWindowHost.Label = "DxWindow Host";
            this.btnDxWindowHost.Name = "btnDxWindowHost";
            this.btnDxWindowHost.ScreenTip = "dx:DXWindow";
            this.btnDxWindowHost.SuperTip = "Super TIp";
            this.btnDxWindowHost.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDxWindowHost_Click);
            // 
            // rgWPFUI
            // 
            this.rgWPFUI.Items.Add(this.btnLaunchCylon);
            this.rgWPFUI.Items.Add(this.btnLaunchCylon2);
            this.rgWPFUI.Items.Add(this.btnDxDockLayoutManagerControl);
            this.rgWPFUI.Items.Add(this.btnDxLayoutControl);
            this.rgWPFUI.Items.Add(this.btnDxDockLayoutControl);
            this.rgWPFUI.Items.Add(this.btnPrismRegionTest);
            this.rgWPFUI.Label = "WPF UI";
            this.rgWPFUI.Name = "rgWPFUI";
            // 
            // btnLaunchCylon
            // 
            this.btnLaunchCylon.Label = "Launch Cylon";
            this.btnLaunchCylon.Name = "btnLaunchCylon";
            this.btnLaunchCylon.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLaunchCylon_Click);
            // 
            // btnLaunchCylon2
            // 
            this.btnLaunchCylon2.Label = "Launch Cylon 2";
            this.btnLaunchCylon2.Name = "btnLaunchCylon2";
            this.btnLaunchCylon2.ScreenTip = "Uses VNC.Core.Xaml.Presentation.WindowHost";
            this.btnLaunchCylon2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLaunchCylon2_Click);
            // 
            // btnDxDockLayoutManagerControl
            // 
            this.btnDxDockLayoutManagerControl.Label = "DxDockLayoutManagerControl";
            this.btnDxDockLayoutManagerControl.Name = "btnDxDockLayoutManagerControl";
            this.btnDxDockLayoutManagerControl.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDxDockLayoutManagerControl_Click);
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
            // btnPrismRegionTest
            // 
            this.btnPrismRegionTest.Label = "Prism Region Test";
            this.btnPrismRegionTest.Name = "btnPrismRegionTest";
            this.btnPrismRegionTest.ScreenTip = "Uses SupportTools_Visio ThemedWindowHost";
            this.btnPrismRegionTest.SuperTip = "Calls ShowUserControl";
            this.btnPrismRegionTest.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnPrismRegionTest_Click);
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
            this.Tabs.Add(this.rtVNCExcelTools);
            this.Tabs.Add(this.rtUILaunchApproaches);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.rtVNCExcelTools.ResumeLayout(false);
            this.rtVNCExcelTools.PerformLayout();
            this.rgDocumentActions.ResumeLayout(false);
            this.rgDocumentActions.PerformLayout();
            this.rgUtilities.ResumeLayout(false);
            this.rgUtilities.PerformLayout();
            this.rgDebug.ResumeLayout(false);
            this.rgDebug.PerformLayout();
            this.grpHelp.ResumeLayout(false);
            this.grpHelp.PerformLayout();
            this.rtUILaunchApproaches.ResumeLayout(false);
            this.rtUILaunchApproaches.PerformLayout();
            this.rgUILaunch.ResumeLayout(false);
            this.rgUILaunch.PerformLayout();
            this.rgWPFUI.ResumeLayout(false);
            this.rgWPFUI.PerformLayout();
            this.rgMVVMExamples.ResumeLayout(false);
            this.rgMVVMExamples.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab rtVNCExcelTools;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup rgDebug;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDebugWindow;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnWatchWindow;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox rcbEnableAppEvents;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox rcbDisplayEvents;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox rcbDisplayChattyEvents;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox rcbToggleDeveloperUIMode;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox rcbUILaunchApproaches;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpHelp;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDisplayAddInInfo;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnToggleDeveloperMode;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab rtUILaunchApproaches;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup rgUILaunch;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnThemedWindowHostModeless;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnThemedWindowHostModal;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnWindowHostLocal;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnWindowHostVNC;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDxWindowHost;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup rgWPFUI;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLaunchCylon;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLaunchCylon2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDxDockLayoutManagerControl;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDxLayoutControl;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDxDockLayoutControl;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPrismRegionTest;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup rgMVVMExamples;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnVNC_MVVM_VAVM1st;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnVNC_MVVM_VA1st;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnVNC_MVVM_VAVM1stDI;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnVNC_MVVM_VB1st;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnVNC_MVVM_VC11st;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnVNC_MVVM_VC21st;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup rgDocumentActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAddTableOfContents;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAddHeader;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAddFooter;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup rgUtilities;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFolderMap;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnTestExcelLogging;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLoggingConfiguration;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
