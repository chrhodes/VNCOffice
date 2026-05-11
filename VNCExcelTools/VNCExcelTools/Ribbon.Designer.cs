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
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.rtVNCExcelTools = this.Factory.CreateRibbonTab();
            this.rgDebug = this.Factory.CreateRibbonGroup();
            this.btnDebugWindow = this.Factory.CreateRibbonButton();
            this.btnWatchWindow = this.Factory.CreateRibbonButton();
            this.rcbEnableAppEvents = this.Factory.CreateRibbonCheckBox();
            this.rcbDisplayEvents = this.Factory.CreateRibbonCheckBox();
            this.rcbDisplayChattyEvents = this.Factory.CreateRibbonCheckBox();
            this.rcbDeveloperUIMode = this.Factory.CreateRibbonCheckBox();
            this.rcbUILaunchApproaches = this.Factory.CreateRibbonCheckBox();
            this.grpHelp = this.Factory.CreateRibbonGroup();
            this.btnAddInInfo = this.Factory.CreateRibbonButton();
            this.btnDeveloperMode = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.rtVNCExcelTools.SuspendLayout();
            this.rgDebug.SuspendLayout();
            this.grpHelp.SuspendLayout();
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
            this.rtVNCExcelTools.Groups.Add(this.rgDebug);
            this.rtVNCExcelTools.Groups.Add(this.grpHelp);
            this.rtVNCExcelTools.Label = "VNCExcelTools";
            this.rtVNCExcelTools.Name = "rtVNCExcelTools";
            // 
            // rgDebug
            // 
            this.rgDebug.Items.Add(this.btnDebugWindow);
            this.rgDebug.Items.Add(this.btnWatchWindow);
            this.rgDebug.Items.Add(this.rcbEnableAppEvents);
            this.rgDebug.Items.Add(this.rcbDisplayEvents);
            this.rgDebug.Items.Add(this.rcbDisplayChattyEvents);
            this.rgDebug.Items.Add(this.rcbDeveloperUIMode);
            this.rgDebug.Items.Add(this.rcbUILaunchApproaches);
            this.rgDebug.Label = "Debug";
            this.rgDebug.Name = "rgDebug";
            this.rgDebug.Visible = false;
            // 
            // btnDebugWindow
            // 
            this.btnDebugWindow.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnDebugWindow.Label = "Debug Window";
            this.btnDebugWindow.Name = "btnDebugWindow";
            this.btnDebugWindow.ShowImage = true;
            // 
            // btnWatchWindow
            // 
            this.btnWatchWindow.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnWatchWindow.Label = "Watch Window";
            this.btnWatchWindow.Name = "btnWatchWindow";
            this.btnWatchWindow.ShowImage = true;
            // 
            // rcbEnableAppEvents
            // 
            this.rcbEnableAppEvents.Label = "Enable App Events";
            this.rcbEnableAppEvents.Name = "rcbEnableAppEvents";
            // 
            // rcbDisplayEvents
            // 
            this.rcbDisplayEvents.Label = "Display Events";
            this.rcbDisplayEvents.Name = "rcbDisplayEvents";
            // 
            // rcbDisplayChattyEvents
            // 
            this.rcbDisplayChattyEvents.Label = "Display Chatty Events";
            this.rcbDisplayChattyEvents.Name = "rcbDisplayChattyEvents";
            // 
            // rcbDeveloperUIMode
            // 
            this.rcbDeveloperUIMode.Label = "DeveloperUIMode";
            this.rcbDeveloperUIMode.Name = "rcbDeveloperUIMode";
            // 
            // rcbUILaunchApproaches
            // 
            this.rcbUILaunchApproaches.Label = "UILaunchApproaches";
            this.rcbUILaunchApproaches.Name = "rcbUILaunchApproaches";
            this.rcbUILaunchApproaches.SuperTip = "Display UILaunchApproaches Ribbon Group";
            // 
            // grpHelp
            // 
            this.grpHelp.Items.Add(this.btnAddInInfo);
            this.grpHelp.Items.Add(this.btnDeveloperMode);
            this.grpHelp.Label = "Help";
            this.grpHelp.Name = "grpHelp";
            // 
            // btnAddInInfo
            // 
            this.btnAddInInfo.Label = "AddIn Info";
            this.btnAddInInfo.Name = "btnAddInInfo";
            // 
            // btnDeveloperMode
            // 
            this.btnDeveloperMode.Label = "Developer Mode";
            this.btnDeveloperMode.Name = "btnDeveloperMode";
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Tabs.Add(this.rtVNCExcelTools);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.rtVNCExcelTools.ResumeLayout(false);
            this.rtVNCExcelTools.PerformLayout();
            this.rgDebug.ResumeLayout(false);
            this.rgDebug.PerformLayout();
            this.grpHelp.ResumeLayout(false);
            this.grpHelp.PerformLayout();
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
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox rcbDeveloperUIMode;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox rcbUILaunchApproaches;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpHelp;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAddInInfo;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDeveloperMode;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
