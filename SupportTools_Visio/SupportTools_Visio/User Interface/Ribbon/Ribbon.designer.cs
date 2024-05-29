﻿namespace SupportTools_Visio
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
            Microsoft.Office.Tools.Ribbon.RibbonDialogLauncher ribbonDialogLauncherImpl1 = this.Factory.CreateRibbonDialogLauncher();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl3 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl4 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl5 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl6 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl7 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl8 = this.Factory.CreateRibbonDropDownItem();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.tabSupportTools = this.Factory.CreateRibbonTab();
            this.rgDocumentActions = this.Factory.CreateRibbonGroup();
            this.btnGetApplicationInfo = this.Factory.CreateRibbonButton();
            this.btnGetDocumentInfo = this.Factory.CreateRibbonButton();
            this.btnGetStencilInfo = this.Factory.CreateRibbonButton();
            this.btnAddTableOfContents = this.Factory.CreateRibbonButton();
            this.btnAddHeader = this.Factory.CreateRibbonButton();
            this.btnAddFooter = this.Factory.CreateRibbonButton();
            this.btnAddDefaultLayers = this.Factory.CreateRibbonButton();
            this.btnRemoveLayers = this.Factory.CreateRibbonButton();
            this.btnSortAllPages = this.Factory.CreateRibbonButton();
            this.btnDisplayPageNames = this.Factory.CreateRibbonButton();
            this.btnSyncPageNames = this.Factory.CreateRibbonButton();
            this.btnAutoSizePagesOn = this.Factory.CreateRibbonButton();
            this.btnAutoSizePagesOff = this.Factory.CreateRibbonButton();
            this.btnUpdatePageNameShapes = this.Factory.CreateRibbonButton();
            this.btnAddNavigationLinks = this.Factory.CreateRibbonButton();
            this.btnPrintPages = this.Factory.CreateRibbonButton();
            this.btnDeletePages = this.Factory.CreateRibbonButton();
            this.btnSavePages = this.Factory.CreateRibbonButton();
            this.rgPageActions = this.Factory.CreateRibbonGroup();
            this.btnGetPageInfo = this.Factory.CreateRibbonButton();
            this.btnUpdatePageNameShapesPage = this.Factory.CreateRibbonButton();
            this.btnAddNavLinks = this.Factory.CreateRibbonButton();
            this.btnPrintPage = this.Factory.CreateRibbonButton();
            this.btnSavePage = this.Factory.CreateRibbonButton();
            this.btnSyncPageNamesPage = this.Factory.CreateRibbonButton();
            this.btnAutoSizePageOn = this.Factory.CreateRibbonButton();
            this.btnAutoSizePageOff = this.Factory.CreateRibbonButton();
            this.rgLayerActions = this.Factory.CreateRibbonGroup();
            this.btnPageOn = this.Factory.CreateRibbonButton();
            this.btnPageOff = this.Factory.CreateRibbonButton();
            this.cmbLayers = this.Factory.CreateRibbonComboBox();
            this.btnAllPageOn = this.Factory.CreateRibbonButton();
            this.btnAllPageOff = this.Factory.CreateRibbonButton();
            this.btnLayerManager = this.Factory.CreateRibbonButton();
            this.btnLockBackground = this.Factory.CreateRibbonButton();
            this.btnUnlockBackground = this.Factory.CreateRibbonButton();
            this.btnAddDefaultLayers_Page = this.Factory.CreateRibbonButton();
            this.btnRemoveLayers_Page = this.Factory.CreateRibbonButton();
            this.rgShapeActions = this.Factory.CreateRibbonGroup();
            this.btnGetShapeInfo = this.Factory.CreateRibbonButton();
            this.btnAddTextControl = this.Factory.CreateRibbonButton();
            this.btnAddIsPageName = this.Factory.CreateRibbonButton();
            this.btnAddHyperLink = this.Factory.CreateRibbonButton();
            this.btnAddColorSupport = this.Factory.CreateRibbonButton();
            this.btnMakeLinkableMaster = this.Factory.CreateRibbonButton();
            this.btnAddIDSupport = this.Factory.CreateRibbonButton();
            this.btnAddIDAndTextSupport = this.Factory.CreateRibbonButton();
            this.btnMoveToBackgroundLayer = this.Factory.CreateRibbonButton();
            this.btn0PtMargins = this.Factory.CreateRibbonButton();
            this.btn1PtMargins = this.Factory.CreateRibbonButton();
            this.btn2PtMargins = this.Factory.CreateRibbonButton();
            this.rgDebug = this.Factory.CreateRibbonGroup();
            this.btnDebugWindow = this.Factory.CreateRibbonButton();
            this.btnWatchWindow = this.Factory.CreateRibbonButton();
            this.chkEnableAppEvents = this.Factory.CreateRibbonCheckBox();
            this.chkDisplayEvents = this.Factory.CreateRibbonCheckBox();
            this.chkDisplayChattyEvents = this.Factory.CreateRibbonCheckBox();
            this.grpHelp = this.Factory.CreateRibbonGroup();
            this.btnAddInInfo = this.Factory.CreateRibbonButton();
            this.btnDeveloperMode = this.Factory.CreateRibbonButton();
            this.tabShapeSheet = this.Factory.CreateRibbonTab();
            this.rgSSDocumentObjectSections = this.Factory.CreateRibbonGroup();
            this.btnDocumentProperties = this.Factory.CreateRibbonButton();
            this.rgSSDocumentRowSections = this.Factory.CreateRibbonGroup();
            this.btnDocumentHyperlinks = this.Factory.CreateRibbonButton();
            this.btnDocumentScratch = this.Factory.CreateRibbonButton();
            this.btnDocumentShapeData = this.Factory.CreateRibbonButton();
            this.btnDocumentUserDefinedCells = this.Factory.CreateRibbonButton();
            this.rgSSPageObjectSections = this.Factory.CreateRibbonGroup();
            this.btnPageLayout = this.Factory.CreateRibbonButton();
            this.btnPageProperties = this.Factory.CreateRibbonButton();
            this.btnPrintProperties = this.Factory.CreateRibbonButton();
            this.btnRulerAndGrid = this.Factory.CreateRibbonButton();
            this.btnPageThemeProperties = this.Factory.CreateRibbonButton();
            this.rgSSPPageRowSections = this.Factory.CreateRibbonGroup();
            this.btnLayers = this.Factory.CreateRibbonButton();
            this.btnPageActions = this.Factory.CreateRibbonButton();
            this.btnPageActionTags = this.Factory.CreateRibbonButton();
            this.btnPageHyperlinks = this.Factory.CreateRibbonButton();
            this.btnPageScratch = this.Factory.CreateRibbonButton();
            this.btnPageShapeData = this.Factory.CreateRibbonButton();
            this.btnPageUserDefinedCells = this.Factory.CreateRibbonButton();
            this.rgSSShapeObjectSections = this.Factory.CreateRibbonGroup();
            this.btn1DEndpoints = this.Factory.CreateRibbonButton();
            this.btn3DRotationProperties = this.Factory.CreateRibbonButton();
            this.btnAdditionalEffectProperties = this.Factory.CreateRibbonButton();
            this.btnBevelProperties = this.Factory.CreateRibbonButton();
            this.btnChangeShapeBehavior = this.Factory.CreateRibbonButton();
            this.btnEvents = this.Factory.CreateRibbonButton();
            this.btnFillFormat = this.Factory.CreateRibbonButton();
            this.btnGlueInfo = this.Factory.CreateRibbonButton();
            this.btnGradientProperties = this.Factory.CreateRibbonButton();
            this.btnGroupProperties = this.Factory.CreateRibbonButton();
            this.btnImageProperties = this.Factory.CreateRibbonButton();
            this.btnLayerMembership = this.Factory.CreateRibbonButton();
            this.btnLineFormat = this.Factory.CreateRibbonButton();
            this.btnMiscelleaneous = this.Factory.CreateRibbonButton();
            this.btnProtection = this.Factory.CreateRibbonButton();
            this.btnQuickStyle = this.Factory.CreateRibbonButton();
            this.btnShapeLayout = this.Factory.CreateRibbonButton();
            this.btnShapeTransform = this.Factory.CreateRibbonButton();
            this.btnTextBlockFormat = this.Factory.CreateRibbonButton();
            this.btnTextTransform = this.Factory.CreateRibbonButton();
            this.btnThemeProperties = this.Factory.CreateRibbonButton();
            this.rgSSShapeRowSections = this.Factory.CreateRibbonGroup();
            this.btnActions = this.Factory.CreateRibbonButton();
            this.btnActionTags = this.Factory.CreateRibbonButton();
            this.btnCharacter = this.Factory.CreateRibbonButton();
            this.btnConnectionPoints = this.Factory.CreateRibbonButton();
            this.btnControls = this.Factory.CreateRibbonButton();
            this.btnGeometry = this.Factory.CreateRibbonButton();
            this.btnGradientStops = this.Factory.CreateRibbonButton();
            this.btnParagraph = this.Factory.CreateRibbonButton();
            this.btnShapeHyperlinks = this.Factory.CreateRibbonButton();
            this.btnShapeScratch = this.Factory.CreateRibbonButton();
            this.btnShapeShapeData = this.Factory.CreateRibbonButton();
            this.btnTabs = this.Factory.CreateRibbonButton();
            this.btnShapeUserDefinedCells = this.Factory.CreateRibbonButton();
            this.tabCustomUI = this.Factory.CreateRibbonTab();
            this.rgCustomUI = this.Factory.CreateRibbonGroup();
            this.btnCommandCockpit = this.Factory.CreateRibbonButton();
            this.btnLinq2Excel = this.Factory.CreateRibbonButton();
            this.btnEditControlRows = this.Factory.CreateRibbonButton();
            this.btnEditParagraph = this.Factory.CreateRibbonButton();
            this.btnEditText = this.Factory.CreateRibbonButton();
            this.btnEditControlPoints = this.Factory.CreateRibbonButton();
            this.btnRenamePages = this.Factory.CreateRibbonButton();
            this.btnDuplicatePage = this.Factory.CreateRibbonButton();
            this.btnMovePages = this.Factory.CreateRibbonButton();
            this.btnCustomUI_Car = this.Factory.CreateRibbonButton();
            this.rgDocumentBasePages = this.Factory.CreateRibbonGroup();
            this.btnAddArchitectureBasePages = this.Factory.CreateRibbonButton();
            this.tabUILaunch = this.Factory.CreateRibbonTab();
            this.rgUILaunch = this.Factory.CreateRibbonGroup();
            this.btnThemedWindowHostModeless = this.Factory.CreateRibbonButton();
            this.btnThemedWindowHostModal = this.Factory.CreateRibbonButton();
            this.btnWindowHostLocal = this.Factory.CreateRibbonButton();
            this.btnWindowHostVNC = this.Factory.CreateRibbonButton();
            this.btnDxWindowHost = this.Factory.CreateRibbonButton();
            this.rgWPFUI = this.Factory.CreateRibbonGroup();
            this.btnLaunchCylon = this.Factory.CreateRibbonButton();
            this.btnLaunchCylon2 = this.Factory.CreateRibbonButton();
            this.btnDxDockLayoutManager = this.Factory.CreateRibbonButton();
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
            this.tabSupportTools.SuspendLayout();
            this.rgDocumentActions.SuspendLayout();
            this.rgPageActions.SuspendLayout();
            this.rgLayerActions.SuspendLayout();
            this.rgShapeActions.SuspendLayout();
            this.rgDebug.SuspendLayout();
            this.grpHelp.SuspendLayout();
            this.tabShapeSheet.SuspendLayout();
            this.rgSSDocumentObjectSections.SuspendLayout();
            this.rgSSDocumentRowSections.SuspendLayout();
            this.rgSSPageObjectSections.SuspendLayout();
            this.rgSSPPageRowSections.SuspendLayout();
            this.rgSSShapeObjectSections.SuspendLayout();
            this.rgSSShapeRowSections.SuspendLayout();
            this.tabCustomUI.SuspendLayout();
            this.rgCustomUI.SuspendLayout();
            this.rgDocumentBasePages.SuspendLayout();
            this.tabUILaunch.SuspendLayout();
            this.rgUILaunch.SuspendLayout();
            this.rgWPFUI.SuspendLayout();
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
            this.tabSupportTools.Groups.Add(this.rgDocumentActions);
            this.tabSupportTools.Groups.Add(this.rgPageActions);
            this.tabSupportTools.Groups.Add(this.rgLayerActions);
            this.tabSupportTools.Groups.Add(this.rgShapeActions);
            this.tabSupportTools.Groups.Add(this.rgDebug);
            this.tabSupportTools.Groups.Add(this.grpHelp);
            this.tabSupportTools.Label = "Support Tools";
            this.tabSupportTools.Name = "tabSupportTools";
            // 
            // rgDocumentActions
            // 
            this.rgDocumentActions.DialogLauncher = ribbonDialogLauncherImpl1;
            this.rgDocumentActions.Items.Add(this.btnGetApplicationInfo);
            this.rgDocumentActions.Items.Add(this.btnGetDocumentInfo);
            this.rgDocumentActions.Items.Add(this.btnGetStencilInfo);
            this.rgDocumentActions.Items.Add(this.btnAddTableOfContents);
            this.rgDocumentActions.Items.Add(this.btnAddHeader);
            this.rgDocumentActions.Items.Add(this.btnAddFooter);
            this.rgDocumentActions.Items.Add(this.btnAddDefaultLayers);
            this.rgDocumentActions.Items.Add(this.btnRemoveLayers);
            this.rgDocumentActions.Items.Add(this.btnSortAllPages);
            this.rgDocumentActions.Items.Add(this.btnDisplayPageNames);
            this.rgDocumentActions.Items.Add(this.btnSyncPageNames);
            this.rgDocumentActions.Items.Add(this.btnAutoSizePagesOn);
            this.rgDocumentActions.Items.Add(this.btnAutoSizePagesOff);
            this.rgDocumentActions.Items.Add(this.btnUpdatePageNameShapes);
            this.rgDocumentActions.Items.Add(this.btnAddNavigationLinks);
            this.rgDocumentActions.Items.Add(this.btnPrintPages);
            this.rgDocumentActions.Items.Add(this.btnDeletePages);
            this.rgDocumentActions.Items.Add(this.btnSavePages);
            this.rgDocumentActions.Label = "Document Actions";
            this.rgDocumentActions.Name = "rgDocumentActions";
            // 
            // btnGetApplicationInfo
            // 
            this.btnGetApplicationInfo.Image = global::SupportTools_Visio.Properties.Resources.Application_Info_64x64;
            this.btnGetApplicationInfo.Label = "Application Info";
            this.btnGetApplicationInfo.Name = "btnGetApplicationInfo";
            this.btnGetApplicationInfo.ScreenTip = "Get Application Info";
            this.btnGetApplicationInfo.ShowImage = true;
            this.btnGetApplicationInfo.SuperTip = "Get Informtation from Application Object";
            this.btnGetApplicationInfo.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnGetApplicationInfo_Click);
            // 
            // btnGetDocumentInfo
            // 
            this.btnGetDocumentInfo.Image = global::SupportTools_Visio.Properties.Resources.Document_Info_64x64;
            this.btnGetDocumentInfo.Label = "Document Info";
            this.btnGetDocumentInfo.Name = "btnGetDocumentInfo";
            this.btnGetDocumentInfo.ScreenTip = "Get Document Info";
            this.btnGetDocumentInfo.ShowImage = true;
            this.btnGetDocumentInfo.SuperTip = "Get Information from Document Object";
            this.btnGetDocumentInfo.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnGetDocumentInfo_Click);
            // 
            // btnGetStencilInfo
            // 
            this.btnGetStencilInfo.Image = global::SupportTools_Visio.Properties.Resources.Stencil_Info_64x64;
            this.btnGetStencilInfo.Label = "Stencil Info";
            this.btnGetStencilInfo.Name = "btnGetStencilInfo";
            this.btnGetStencilInfo.ScreenTip = "Get Stencil Info";
            this.btnGetStencilInfo.ShowImage = true;
            this.btnGetStencilInfo.SuperTip = "Get Information from Stencil Object";
            this.btnGetStencilInfo.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnGetStencilInfo_Click);
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
            // btnAddDefaultLayers
            // 
            this.btnAddDefaultLayers.Label = "+ DefaultLayers";
            this.btnAddDefaultLayers.Name = "btnAddDefaultLayers";
            this.btnAddDefaultLayers.ScreenTip = "Add Default Layers to all Pages";
            this.btnAddDefaultLayers.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAddDefaultLayers_Click);
            // 
            // btnRemoveLayers
            // 
            this.btnRemoveLayers.Label = "Remove Layers";
            this.btnRemoveLayers.Name = "btnRemoveLayers";
            this.btnRemoveLayers.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnRemoveLayers_Click);
            // 
            // btnSortAllPages
            // 
            this.btnSortAllPages.Label = "Sort All Pages";
            this.btnSortAllPages.Name = "btnSortAllPages";
            this.btnSortAllPages.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSortAllPages_Click);
            // 
            // btnDisplayPageNames
            // 
            this.btnDisplayPageNames.Label = "Display Page Names";
            this.btnDisplayPageNames.Name = "btnDisplayPageNames";
            this.btnDisplayPageNames.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDisplayPageNames_Click);
            // 
            // btnSyncPageNames
            // 
            this.btnSyncPageNames.Label = "Sync Name(U)";
            this.btnSyncPageNames.Name = "btnSyncPageNames";
            this.btnSyncPageNames.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSyncPageNames_Click);
            // 
            // btnAutoSizePagesOn
            // 
            this.btnAutoSizePagesOn.Label = "AutoSize On";
            this.btnAutoSizePagesOn.Name = "btnAutoSizePagesOn";
            this.btnAutoSizePagesOn.SuperTip = "Turn On AutoSize for All Pages";
            this.btnAutoSizePagesOn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAutoSizePagesOn_Click);
            // 
            // btnAutoSizePagesOff
            // 
            this.btnAutoSizePagesOff.Label = "AutoSize Off";
            this.btnAutoSizePagesOff.Name = "btnAutoSizePagesOff";
            this.btnAutoSizePagesOff.SuperTip = "Turn Off AutoSize for All Pages";
            this.btnAutoSizePagesOff.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAutoSizePagesOff_Click);
            // 
            // btnUpdatePageNameShapes
            // 
            this.btnUpdatePageNameShapes.Image = global::SupportTools_Visio.Properties.Resources.Update_Name_Shapes64x64;
            this.btnUpdatePageNameShapes.Label = "Update Shapes";
            this.btnUpdatePageNameShapes.Name = "btnUpdatePageNameShapes";
            this.btnUpdatePageNameShapes.ScreenTip = "Update PageName Shapes";
            this.btnUpdatePageNameShapes.ShowImage = true;
            this.btnUpdatePageNameShapes.SuperTip = "Update Page Name Shapes from Page Name text";
            this.btnUpdatePageNameShapes.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUpdatePageNameShapes_Click);
            // 
            // btnAddNavigationLinks
            // 
            this.btnAddNavigationLinks.Image = global::SupportTools_Visio.Properties.Resources.Navigation_Links_64x64;
            this.btnAddNavigationLinks.Label = "Update Nav Links";
            this.btnAddNavigationLinks.Name = "btnAddNavigationLinks";
            this.btnAddNavigationLinks.ScreenTip = "Add Navigation Links";
            this.btnAddNavigationLinks.ShowImage = true;
            this.btnAddNavigationLinks.SuperTip = "Add Navigation Links from Navigation Links Background Page";
            this.btnAddNavigationLinks.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAddNavigationLinks_Click);
            // 
            // btnPrintPages
            // 
            this.btnPrintPages.Label = "Print Pages";
            this.btnPrintPages.Name = "btnPrintPages";
            this.btnPrintPages.SuperTip = "Print all Pages listed on current Page";
            this.btnPrintPages.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnPrintPages_Click);
            // 
            // btnDeletePages
            // 
            this.btnDeletePages.Label = "Delete Pages";
            this.btnDeletePages.Name = "btnDeletePages";
            this.btnDeletePages.SuperTip = "Delete all Pages listed on current Page to Another Document";
            this.btnDeletePages.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDeletePages_Click);
            // 
            // btnSavePages
            // 
            this.btnSavePages.Label = "Save Pages";
            this.btnSavePages.Name = "btnSavePages";
            this.btnSavePages.SuperTip = "Save all Pages listed on current Page to Image";
            this.btnSavePages.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSavePages_Click);
            // 
            // rgPageActions
            // 
            this.rgPageActions.Items.Add(this.btnGetPageInfo);
            this.rgPageActions.Items.Add(this.btnUpdatePageNameShapesPage);
            this.rgPageActions.Items.Add(this.btnAddNavLinks);
            this.rgPageActions.Items.Add(this.btnPrintPage);
            this.rgPageActions.Items.Add(this.btnSavePage);
            this.rgPageActions.Items.Add(this.btnSyncPageNamesPage);
            this.rgPageActions.Items.Add(this.btnAutoSizePageOn);
            this.rgPageActions.Items.Add(this.btnAutoSizePageOff);
            this.rgPageActions.Label = "Page Actions";
            this.rgPageActions.Name = "rgPageActions";
            // 
            // btnGetPageInfo
            // 
            this.btnGetPageInfo.Image = global::SupportTools_Visio.Properties.Resources.Page_Info_64x64;
            this.btnGetPageInfo.Label = "Page Info";
            this.btnGetPageInfo.Name = "btnGetPageInfo";
            this.btnGetPageInfo.ScreenTip = "Get Page Info";
            this.btnGetPageInfo.ShowImage = true;
            this.btnGetPageInfo.SuperTip = "Get Information from Page Object";
            this.btnGetPageInfo.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnGetPageInfo_Click);
            // 
            // btnUpdatePageNameShapesPage
            // 
            this.btnUpdatePageNameShapesPage.Image = global::SupportTools_Visio.Properties.Resources.Update_Name_Shapes64x64;
            this.btnUpdatePageNameShapesPage.Label = "Update Shapes";
            this.btnUpdatePageNameShapesPage.Name = "btnUpdatePageNameShapesPage";
            this.btnUpdatePageNameShapesPage.ScreenTip = "Update PageName Shapes";
            this.btnUpdatePageNameShapesPage.ShowImage = true;
            this.btnUpdatePageNameShapesPage.SuperTip = "Update Page Name Shapes from Page Name text";
            this.btnUpdatePageNameShapesPage.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUpdatePageNameShapesPage_Click);
            // 
            // btnAddNavLinks
            // 
            this.btnAddNavLinks.Image = global::SupportTools_Visio.Properties.Resources.Navigation_Links_64x64;
            this.btnAddNavLinks.Label = "Nav Links";
            this.btnAddNavLinks.Name = "btnAddNavLinks";
            this.btnAddNavLinks.ScreenTip = "Add Navigation Links";
            this.btnAddNavLinks.ShowImage = true;
            this.btnAddNavLinks.SuperTip = "Add Navigation Links from Navigation Links Background Page";
            this.btnAddNavLinks.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAddNavLinks_Click);
            // 
            // btnPrintPage
            // 
            this.btnPrintPage.Label = "Print Page";
            this.btnPrintPage.Name = "btnPrintPage";
            this.btnPrintPage.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnPrintPage_Click);
            // 
            // btnSavePage
            // 
            this.btnSavePage.Label = "Save Page";
            this.btnSavePage.Name = "btnSavePage";
            this.btnSavePage.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSavePage_Click);
            // 
            // btnSyncPageNamesPage
            // 
            this.btnSyncPageNamesPage.Label = "Sync Name(U)";
            this.btnSyncPageNamesPage.Name = "btnSyncPageNamesPage";
            this.btnSyncPageNamesPage.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSyncPageNamesPage_Click);
            // 
            // btnAutoSizePageOn
            // 
            this.btnAutoSizePageOn.Label = "AutoSize Page On";
            this.btnAutoSizePageOn.Name = "btnAutoSizePageOn";
            this.btnAutoSizePageOn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAutoSizePageOn_Click);
            // 
            // btnAutoSizePageOff
            // 
            this.btnAutoSizePageOff.Label = "AutoSize Page Off";
            this.btnAutoSizePageOff.Name = "btnAutoSizePageOff";
            this.btnAutoSizePageOff.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAutoSizePageOff_Click);
            // 
            // rgLayerActions
            // 
            this.rgLayerActions.Items.Add(this.btnPageOn);
            this.rgLayerActions.Items.Add(this.btnPageOff);
            this.rgLayerActions.Items.Add(this.cmbLayers);
            this.rgLayerActions.Items.Add(this.btnAllPageOn);
            this.rgLayerActions.Items.Add(this.btnAllPageOff);
            this.rgLayerActions.Items.Add(this.btnLayerManager);
            this.rgLayerActions.Items.Add(this.btnLockBackground);
            this.rgLayerActions.Items.Add(this.btnUnlockBackground);
            this.rgLayerActions.Items.Add(this.btnAddDefaultLayers_Page);
            this.rgLayerActions.Items.Add(this.btnRemoveLayers_Page);
            this.rgLayerActions.Label = "Layer Actions";
            this.rgLayerActions.Name = "rgLayerActions";
            // 
            // btnPageOn
            // 
            this.btnPageOn.Label = "Page On";
            this.btnPageOn.Name = "btnPageOn";
            this.btnPageOn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnPageOn_Click);
            // 
            // btnPageOff
            // 
            this.btnPageOff.Label = "Page Off";
            this.btnPageOff.Name = "btnPageOff";
            this.btnPageOff.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnPageOff_Click);
            // 
            // cmbLayers
            // 
            ribbonDropDownItemImpl1.Label = "Navigation";
            ribbonDropDownItemImpl2.Label = "Header";
            ribbonDropDownItemImpl3.Label = "Security";
            ribbonDropDownItemImpl4.Label = "Application";
            ribbonDropDownItemImpl5.Label = "Level0";
            ribbonDropDownItemImpl6.Label = "Level1";
            ribbonDropDownItemImpl7.Label = "Level2";
            ribbonDropDownItemImpl8.Label = "Notes";
            this.cmbLayers.Items.Add(ribbonDropDownItemImpl1);
            this.cmbLayers.Items.Add(ribbonDropDownItemImpl2);
            this.cmbLayers.Items.Add(ribbonDropDownItemImpl3);
            this.cmbLayers.Items.Add(ribbonDropDownItemImpl4);
            this.cmbLayers.Items.Add(ribbonDropDownItemImpl5);
            this.cmbLayers.Items.Add(ribbonDropDownItemImpl6);
            this.cmbLayers.Items.Add(ribbonDropDownItemImpl7);
            this.cmbLayers.Items.Add(ribbonDropDownItemImpl8);
            this.cmbLayers.Label = "Layer";
            this.cmbLayers.Name = "cmbLayers";
            this.cmbLayers.Text = null;
            // 
            // btnAllPageOn
            // 
            this.btnAllPageOn.Label = "All Pages On";
            this.btnAllPageOn.Name = "btnAllPageOn";
            this.btnAllPageOn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAllPageOn_Click);
            // 
            // btnAllPageOff
            // 
            this.btnAllPageOff.Label = "All Pages Off";
            this.btnAllPageOff.Name = "btnAllPageOff";
            this.btnAllPageOff.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAllPageOff_Click);
            // 
            // btnLayerManager
            // 
            this.btnLayerManager.Label = "Layer Manager";
            this.btnLayerManager.Name = "btnLayerManager";
            this.btnLayerManager.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLoadLayers_Click);
            // 
            // btnLockBackground
            // 
            this.btnLockBackground.Label = "Lock Background";
            this.btnLockBackground.Name = "btnLockBackground";
            this.btnLockBackground.ScreenTip = "Lock Background Layer";
            this.btnLockBackground.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLockBackground_Click);
            // 
            // btnUnlockBackground
            // 
            this.btnUnlockBackground.Label = "Unlock Background";
            this.btnUnlockBackground.Name = "btnUnlockBackground";
            this.btnUnlockBackground.ScreenTip = "Unlock Background Layer";
            this.btnUnlockBackground.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUnlockBackground_Click);
            // 
            // btnAddDefaultLayers_Page
            // 
            this.btnAddDefaultLayers_Page.Label = "Add DefaultLayers";
            this.btnAddDefaultLayers_Page.Name = "btnAddDefaultLayers_Page";
            // 
            // btnRemoveLayers_Page
            // 
            this.btnRemoveLayers_Page.Label = "Remove Layers";
            this.btnRemoveLayers_Page.Name = "btnRemoveLayers_Page";
            this.btnRemoveLayers_Page.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnRemoveLayers_Page_Click);
            // 
            // rgShapeActions
            // 
            this.rgShapeActions.Items.Add(this.btnGetShapeInfo);
            this.rgShapeActions.Items.Add(this.btnAddTextControl);
            this.rgShapeActions.Items.Add(this.btnAddIsPageName);
            this.rgShapeActions.Items.Add(this.btnAddHyperLink);
            this.rgShapeActions.Items.Add(this.btnAddColorSupport);
            this.rgShapeActions.Items.Add(this.btnMakeLinkableMaster);
            this.rgShapeActions.Items.Add(this.btnAddIDSupport);
            this.rgShapeActions.Items.Add(this.btnAddIDAndTextSupport);
            this.rgShapeActions.Items.Add(this.btnMoveToBackgroundLayer);
            this.rgShapeActions.Items.Add(this.btn0PtMargins);
            this.rgShapeActions.Items.Add(this.btn1PtMargins);
            this.rgShapeActions.Items.Add(this.btn2PtMargins);
            this.rgShapeActions.Label = "Shape Actions";
            this.rgShapeActions.Name = "rgShapeActions";
            // 
            // btnGetShapeInfo
            // 
            this.btnGetShapeInfo.Image = global::SupportTools_Visio.Properties.Resources.Shape_Info_64x64;
            this.btnGetShapeInfo.Label = "Shape Info";
            this.btnGetShapeInfo.Name = "btnGetShapeInfo";
            this.btnGetShapeInfo.ScreenTip = "Get Shape Info";
            this.btnGetShapeInfo.ShowImage = true;
            this.btnGetShapeInfo.SuperTip = "Get Information from Shape Object";
            this.btnGetShapeInfo.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnGetShapeInfo_Click);
            // 
            // btnAddTextControl
            // 
            this.btnAddTextControl.Label = "+ Text Control";
            this.btnAddTextControl.Name = "btnAddTextControl";
            this.btnAddTextControl.ScreenTip = "Add Text Transform Control to Shape";
            this.btnAddTextControl.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAddTextControl_Click);
            // 
            // btnAddIsPageName
            // 
            this.btnAddIsPageName.Label = "+ IsPageName";
            this.btnAddIsPageName.Name = "btnAddIsPageName";
            this.btnAddIsPageName.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAddIsPageName_Click);
            // 
            // btnAddHyperLink
            // 
            this.btnAddHyperLink.Label = "+ HyperLink";
            this.btnAddHyperLink.Name = "btnAddHyperLink";
            this.btnAddHyperLink.ScreenTip = "Add HyperLink to Page with same name";
            this.btnAddHyperLink.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAddHyperLink_Click);
            // 
            // btnAddColorSupport
            // 
            this.btnAddColorSupport.Label = "+ Color Support";
            this.btnAddColorSupport.Name = "btnAddColorSupport";
            this.btnAddColorSupport.ScreenTip = "Add Color Support to Shape";
            this.btnAddColorSupport.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAddColorSupport_Click);
            // 
            // btnMakeLinkableMaster
            // 
            this.btnMakeLinkableMaster.Label = "Linkable Master";
            this.btnMakeLinkableMaster.Name = "btnMakeLinkableMaster";
            this.btnMakeLinkableMaster.ScreenTip = "Make Linkable Master";
            this.btnMakeLinkableMaster.SuperTip = "Make Linkable Master by adding Action Sections";
            this.btnMakeLinkableMaster.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnMakeLinkableMaster_Click);
            // 
            // btnAddIDSupport
            // 
            this.btnAddIDSupport.Label = "+ ID Support";
            this.btnAddIDSupport.Name = "btnAddIDSupport";
            this.btnAddIDSupport.ScreenTip = "Add ID Support to Shape";
            this.btnAddIDSupport.SuperTip = "Add ID Support to Shape by adding Shape Data";
            this.btnAddIDSupport.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAddIDSupport_Click);
            // 
            // btnAddIDAndTextSupport
            // 
            this.btnAddIDAndTextSupport.Label = "+ ID/Text Support";
            this.btnAddIDAndTextSupport.Name = "btnAddIDAndTextSupport";
            this.btnAddIDAndTextSupport.ScreenTip = "Add ID and Text Box suppor to shape";
            this.btnAddIDAndTextSupport.SuperTip = "Add ID and Text Box suppor to shape by adding Shape Data";
            this.btnAddIDAndTextSupport.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAddIDAndTextSupport_Click);
            // 
            // btnMoveToBackgroundLayer
            // 
            this.btnMoveToBackgroundLayer.Label = "-> Background";
            this.btnMoveToBackgroundLayer.Name = "btnMoveToBackgroundLayer";
            this.btnMoveToBackgroundLayer.ScreenTip = "Move Shape to Backgroud Layer";
            this.btnMoveToBackgroundLayer.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnMoveToBackgroundLayer_Click);
            // 
            // btn0PtMargins
            // 
            this.btn0PtMargins.Label = "0pt Margins";
            this.btn0PtMargins.Name = "btn0PtMargins";
            this.btn0PtMargins.ScreenTip = "0 Pt Text Block Margins for selected Shapes";
            this.btn0PtMargins.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn0PtMargins_Click);
            // 
            // btn1PtMargins
            // 
            this.btn1PtMargins.Label = "1 Pt Margins";
            this.btn1PtMargins.Name = "btn1PtMargins";
            this.btn1PtMargins.ScreenTip = "1 Pt Text Block Margins for selected Shapes";
            this.btn1PtMargins.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn1PtMargins_Click);
            // 
            // btn2PtMargins
            // 
            this.btn2PtMargins.Label = "2 Pt Margins";
            this.btn2PtMargins.Name = "btn2PtMargins";
            this.btn2PtMargins.ScreenTip = "2 Pt Text Block Margins for selected Shapes";
            this.btn2PtMargins.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn2PtMargins_Click);
            // 
            // rgDebug
            // 
            this.rgDebug.Items.Add(this.btnDebugWindow);
            this.rgDebug.Items.Add(this.btnWatchWindow);
            this.rgDebug.Items.Add(this.chkEnableAppEvents);
            this.rgDebug.Items.Add(this.chkDisplayEvents);
            this.rgDebug.Items.Add(this.chkDisplayChattyEvents);
            this.rgDebug.Label = "Debug";
            this.rgDebug.Name = "rgDebug";
            this.rgDebug.Visible = false;
            // 
            // btnDebugWindow
            // 
            this.btnDebugWindow.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnDebugWindow.Image = global::SupportTools_Visio.Properties.Resources.Auto_Debug_System_icon;
            this.btnDebugWindow.Label = "Debug Window";
            this.btnDebugWindow.Name = "btnDebugWindow";
            this.btnDebugWindow.ShowImage = true;
            this.btnDebugWindow.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDebugWindow_Click);
            // 
            // btnWatchWindow
            // 
            this.btnWatchWindow.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnWatchWindow.Image = global::SupportTools_Visio.Properties.Resources.WatchWindow;
            this.btnWatchWindow.Label = "Watch Window";
            this.btnWatchWindow.Name = "btnWatchWindow";
            this.btnWatchWindow.ShowImage = true;
            this.btnWatchWindow.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnWatchWindow_Click);
            // 
            // chkEnableAppEvents
            // 
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
            // chkDisplayChattyEvents
            // 
            this.chkDisplayChattyEvents.Label = "Display Chatty Events";
            this.chkDisplayChattyEvents.Name = "chkDisplayChattyEvents";
            this.chkDisplayChattyEvents.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.chkDisplayChattyEvents_Click);
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
            this.btnAddInInfo.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAddInInfo_Click);
            // 
            // btnDeveloperMode
            // 
            this.btnDeveloperMode.Label = "Developer Mode";
            this.btnDeveloperMode.Name = "btnDeveloperMode";
            this.btnDeveloperMode.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDeveloperMode_Click);
            // 
            // tabShapeSheet
            // 
            this.tabShapeSheet.Groups.Add(this.rgSSDocumentObjectSections);
            this.tabShapeSheet.Groups.Add(this.rgSSDocumentRowSections);
            this.tabShapeSheet.Groups.Add(this.rgSSPageObjectSections);
            this.tabShapeSheet.Groups.Add(this.rgSSPPageRowSections);
            this.tabShapeSheet.Groups.Add(this.rgSSShapeObjectSections);
            this.tabShapeSheet.Groups.Add(this.rgSSShapeRowSections);
            this.tabShapeSheet.Label = "ShapeSheet";
            this.tabShapeSheet.Name = "tabShapeSheet";
            // 
            // rgSSDocumentObjectSections
            // 
            this.rgSSDocumentObjectSections.Items.Add(this.btnDocumentProperties);
            this.rgSSDocumentObjectSections.Label = "Document (Object)";
            this.rgSSDocumentObjectSections.Name = "rgSSDocumentObjectSections";
            // 
            // btnDocumentProperties
            // 
            this.btnDocumentProperties.Label = "Document Properties";
            this.btnDocumentProperties.Name = "btnDocumentProperties";
            this.btnDocumentProperties.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDocumentProperties_Click);
            // 
            // rgSSDocumentRowSections
            // 
            this.rgSSDocumentRowSections.Items.Add(this.btnDocumentHyperlinks);
            this.rgSSDocumentRowSections.Items.Add(this.btnDocumentScratch);
            this.rgSSDocumentRowSections.Items.Add(this.btnDocumentShapeData);
            this.rgSSDocumentRowSections.Items.Add(this.btnDocumentUserDefinedCells);
            this.rgSSDocumentRowSections.Label = "Document (Rows)";
            this.rgSSDocumentRowSections.Name = "rgSSDocumentRowSections";
            // 
            // btnDocumentHyperlinks
            // 
            this.btnDocumentHyperlinks.Label = "Hyperlinks";
            this.btnDocumentHyperlinks.Name = "btnDocumentHyperlinks";
            this.btnDocumentHyperlinks.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDocumentHyperlinks_Click);
            // 
            // btnDocumentScratch
            // 
            this.btnDocumentScratch.Label = "Scratch";
            this.btnDocumentScratch.Name = "btnDocumentScratch";
            this.btnDocumentScratch.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDocumentScratch_Click);
            // 
            // btnDocumentShapeData
            // 
            this.btnDocumentShapeData.Label = "Shape Data";
            this.btnDocumentShapeData.Name = "btnDocumentShapeData";
            this.btnDocumentShapeData.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDocumentShapeData_Click);
            // 
            // btnDocumentUserDefinedCells
            // 
            this.btnDocumentUserDefinedCells.Label = "User-Defined Cells";
            this.btnDocumentUserDefinedCells.Name = "btnDocumentUserDefinedCells";
            this.btnDocumentUserDefinedCells.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDocumentUserDefinedCells_Click);
            // 
            // rgSSPageObjectSections
            // 
            this.rgSSPageObjectSections.Items.Add(this.btnPageLayout);
            this.rgSSPageObjectSections.Items.Add(this.btnPageProperties);
            this.rgSSPageObjectSections.Items.Add(this.btnPrintProperties);
            this.rgSSPageObjectSections.Items.Add(this.btnRulerAndGrid);
            this.rgSSPageObjectSections.Items.Add(this.btnPageThemeProperties);
            this.rgSSPageObjectSections.Label = "Page (Object)";
            this.rgSSPageObjectSections.Name = "rgSSPageObjectSections";
            // 
            // btnPageLayout
            // 
            this.btnPageLayout.Label = "Page Layout";
            this.btnPageLayout.Name = "btnPageLayout";
            this.btnPageLayout.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnPageLayout_Click);
            // 
            // btnPageProperties
            // 
            this.btnPageProperties.Label = "Page Properties";
            this.btnPageProperties.Name = "btnPageProperties";
            this.btnPageProperties.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnPageProperties_Click);
            // 
            // btnPrintProperties
            // 
            this.btnPrintProperties.Label = "Print Properties";
            this.btnPrintProperties.Name = "btnPrintProperties";
            this.btnPrintProperties.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnPrintProperties_Click);
            // 
            // btnRulerAndGrid
            // 
            this.btnRulerAndGrid.Label = "Ruler and Grid";
            this.btnRulerAndGrid.Name = "btnRulerAndGrid";
            this.btnRulerAndGrid.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnRulerAndGrid_Click);
            // 
            // btnPageThemeProperties
            // 
            this.btnPageThemeProperties.Label = "Theme Properties";
            this.btnPageThemeProperties.Name = "btnPageThemeProperties";
            this.btnPageThemeProperties.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnPageThemeProperties_Click);
            // 
            // rgSSPPageRowSections
            // 
            this.rgSSPPageRowSections.Items.Add(this.btnLayers);
            this.rgSSPPageRowSections.Items.Add(this.btnPageActions);
            this.rgSSPPageRowSections.Items.Add(this.btnPageActionTags);
            this.rgSSPPageRowSections.Items.Add(this.btnPageHyperlinks);
            this.rgSSPPageRowSections.Items.Add(this.btnPageScratch);
            this.rgSSPPageRowSections.Items.Add(this.btnPageShapeData);
            this.rgSSPPageRowSections.Items.Add(this.btnPageUserDefinedCells);
            this.rgSSPPageRowSections.Label = "Page (Rows)";
            this.rgSSPPageRowSections.Name = "rgSSPPageRowSections";
            // 
            // btnLayers
            // 
            this.btnLayers.Label = "Layers";
            this.btnLayers.Name = "btnLayers";
            this.btnLayers.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLayers_Click);
            // 
            // btnPageActions
            // 
            this.btnPageActions.Label = "Actions";
            this.btnPageActions.Name = "btnPageActions";
            this.btnPageActions.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnActionsPage_Click);
            // 
            // btnPageActionTags
            // 
            this.btnPageActionTags.Label = "Action Tags";
            this.btnPageActionTags.Name = "btnPageActionTags";
            this.btnPageActionTags.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnActionTagsPage_Click);
            // 
            // btnPageHyperlinks
            // 
            this.btnPageHyperlinks.Label = "Hyperlinks";
            this.btnPageHyperlinks.Name = "btnPageHyperlinks";
            this.btnPageHyperlinks.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnPageHyperlinks_Click);
            // 
            // btnPageScratch
            // 
            this.btnPageScratch.Label = "Scratch";
            this.btnPageScratch.Name = "btnPageScratch";
            this.btnPageScratch.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnPageScratch_Click);
            // 
            // btnPageShapeData
            // 
            this.btnPageShapeData.Label = "Shape Data";
            this.btnPageShapeData.Name = "btnPageShapeData";
            this.btnPageShapeData.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnPageShapeData_Click);
            // 
            // btnPageUserDefinedCells
            // 
            this.btnPageUserDefinedCells.Label = "User-Defined Cells";
            this.btnPageUserDefinedCells.Name = "btnPageUserDefinedCells";
            this.btnPageUserDefinedCells.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnPageUserDefinedCells_Click);
            // 
            // rgSSShapeObjectSections
            // 
            this.rgSSShapeObjectSections.Items.Add(this.btn1DEndpoints);
            this.rgSSShapeObjectSections.Items.Add(this.btn3DRotationProperties);
            this.rgSSShapeObjectSections.Items.Add(this.btnAdditionalEffectProperties);
            this.rgSSShapeObjectSections.Items.Add(this.btnBevelProperties);
            this.rgSSShapeObjectSections.Items.Add(this.btnChangeShapeBehavior);
            this.rgSSShapeObjectSections.Items.Add(this.btnEvents);
            this.rgSSShapeObjectSections.Items.Add(this.btnFillFormat);
            this.rgSSShapeObjectSections.Items.Add(this.btnGlueInfo);
            this.rgSSShapeObjectSections.Items.Add(this.btnGradientProperties);
            this.rgSSShapeObjectSections.Items.Add(this.btnGroupProperties);
            this.rgSSShapeObjectSections.Items.Add(this.btnImageProperties);
            this.rgSSShapeObjectSections.Items.Add(this.btnLayerMembership);
            this.rgSSShapeObjectSections.Items.Add(this.btnLineFormat);
            this.rgSSShapeObjectSections.Items.Add(this.btnMiscelleaneous);
            this.rgSSShapeObjectSections.Items.Add(this.btnProtection);
            this.rgSSShapeObjectSections.Items.Add(this.btnQuickStyle);
            this.rgSSShapeObjectSections.Items.Add(this.btnShapeLayout);
            this.rgSSShapeObjectSections.Items.Add(this.btnShapeTransform);
            this.rgSSShapeObjectSections.Items.Add(this.btnTextBlockFormat);
            this.rgSSShapeObjectSections.Items.Add(this.btnTextTransform);
            this.rgSSShapeObjectSections.Items.Add(this.btnThemeProperties);
            this.rgSSShapeObjectSections.Label = "Shape (Object)";
            this.rgSSShapeObjectSections.Name = "rgSSShapeObjectSections";
            // 
            // btn1DEndpoints
            // 
            this.btn1DEndpoints.Label = "1D Endpoints";
            this.btn1DEndpoints.Name = "btn1DEndpoints";
            this.btn1DEndpoints.ScreenTip = "1D Endpoints tip";
            this.btn1DEndpoints.SuperTip = "What does this section do";
            this.btn1DEndpoints.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn1DEndpoints_Click);
            // 
            // btn3DRotationProperties
            // 
            this.btn3DRotationProperties.Label = "3D Rotation Properties";
            this.btn3DRotationProperties.Name = "btn3DRotationProperties";
            this.btn3DRotationProperties.ScreenTip = "3D Rotation Properties tip";
            this.btn3DRotationProperties.SuperTip = "What does this section do";
            this.btn3DRotationProperties.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn3DRotationProperties_Click);
            // 
            // btnAdditionalEffectProperties
            // 
            this.btnAdditionalEffectProperties.Label = "Additional Effect Properties";
            this.btnAdditionalEffectProperties.Name = "btnAdditionalEffectProperties";
            this.btnAdditionalEffectProperties.ScreenTip = "Edit Text";
            this.btnAdditionalEffectProperties.SuperTip = "Launch the Super Duper Edit Text UI";
            this.btnAdditionalEffectProperties.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAdditionalEffectProperties_Click);
            // 
            // btnBevelProperties
            // 
            this.btnBevelProperties.Label = "Bevel Properties";
            this.btnBevelProperties.Name = "btnBevelProperties";
            this.btnBevelProperties.ScreenTip = "Edit Text";
            this.btnBevelProperties.SuperTip = "Launch the Super Duper Edit Text UI";
            this.btnBevelProperties.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnBevelProperties_Click);
            // 
            // btnChangeShapeBehavior
            // 
            this.btnChangeShapeBehavior.Label = "Change Shape Behavior";
            this.btnChangeShapeBehavior.Name = "btnChangeShapeBehavior";
            this.btnChangeShapeBehavior.ScreenTip = "Edit Text";
            this.btnChangeShapeBehavior.SuperTip = "Launch the Super Duper Edit Text UI";
            this.btnChangeShapeBehavior.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnChangeShapeBehavior_Click);
            // 
            // btnEvents
            // 
            this.btnEvents.Label = "Events";
            this.btnEvents.Name = "btnEvents";
            this.btnEvents.ScreenTip = "Edit Text";
            this.btnEvents.SuperTip = "Launch the Super Duper Edit Text UI";
            this.btnEvents.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnEvents_Click);
            // 
            // btnFillFormat
            // 
            this.btnFillFormat.Label = "Fill Format";
            this.btnFillFormat.Name = "btnFillFormat";
            this.btnFillFormat.ScreenTip = "Edit Text";
            this.btnFillFormat.SuperTip = "Launch the Super Duper Edit Text UI";
            this.btnFillFormat.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFillFormat_Click);
            // 
            // btnGlueInfo
            // 
            this.btnGlueInfo.Label = "Glue Info";
            this.btnGlueInfo.Name = "btnGlueInfo";
            this.btnGlueInfo.ScreenTip = "Edit Text";
            this.btnGlueInfo.SuperTip = "Launch the Super Duper Edit Text UI";
            this.btnGlueInfo.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnGlueInfo_Click);
            // 
            // btnGradientProperties
            // 
            this.btnGradientProperties.Label = "Gradient Properties";
            this.btnGradientProperties.Name = "btnGradientProperties";
            this.btnGradientProperties.ScreenTip = "Edit Text";
            this.btnGradientProperties.SuperTip = "Launch the Super Duper Edit Text UI";
            this.btnGradientProperties.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnGradientProperties_Click);
            // 
            // btnGroupProperties
            // 
            this.btnGroupProperties.Label = "Group Properties";
            this.btnGroupProperties.Name = "btnGroupProperties";
            this.btnGroupProperties.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnGroupProperties_Click);
            // 
            // btnImageProperties
            // 
            this.btnImageProperties.Label = "Image Properties";
            this.btnImageProperties.Name = "btnImageProperties";
            this.btnImageProperties.ScreenTip = "Edit Text";
            this.btnImageProperties.SuperTip = "Launch the Super Duper Edit Text UI";
            this.btnImageProperties.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnImageProperties_Click);
            // 
            // btnLayerMembership
            // 
            this.btnLayerMembership.Label = "Layer Membership";
            this.btnLayerMembership.Name = "btnLayerMembership";
            this.btnLayerMembership.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLayerMembership_Click);
            // 
            // btnLineFormat
            // 
            this.btnLineFormat.Label = "Line Format";
            this.btnLineFormat.Name = "btnLineFormat";
            this.btnLineFormat.ScreenTip = "Edit Text";
            this.btnLineFormat.SuperTip = "Launch the Super Duper Edit Text UI";
            this.btnLineFormat.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLineFormat_Click);
            // 
            // btnMiscelleaneous
            // 
            this.btnMiscelleaneous.Label = "Miscellaneous";
            this.btnMiscelleaneous.Name = "btnMiscelleaneous";
            this.btnMiscelleaneous.ScreenTip = "Edit Text";
            this.btnMiscelleaneous.SuperTip = "Launch the Super Duper Edit Text UI";
            this.btnMiscelleaneous.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnMiscelleaneous_Click);
            // 
            // btnProtection
            // 
            this.btnProtection.Label = "Protection";
            this.btnProtection.Name = "btnProtection";
            this.btnProtection.ScreenTip = "Edit Text";
            this.btnProtection.SuperTip = "Launch the Super Duper Edit Text UI";
            this.btnProtection.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnProtection_Click);
            // 
            // btnQuickStyle
            // 
            this.btnQuickStyle.Label = "Quick Style";
            this.btnQuickStyle.Name = "btnQuickStyle";
            this.btnQuickStyle.ScreenTip = "Edit Text";
            this.btnQuickStyle.SuperTip = "Launch the Super Duper Edit Text UI";
            this.btnQuickStyle.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnQuickStyle_Click);
            // 
            // btnShapeLayout
            // 
            this.btnShapeLayout.Label = "Shape Layout";
            this.btnShapeLayout.Name = "btnShapeLayout";
            this.btnShapeLayout.ScreenTip = "Edit Text";
            this.btnShapeLayout.SuperTip = "Launch the Super Duper Edit Text UI";
            this.btnShapeLayout.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnShapeLayout_Click);
            // 
            // btnShapeTransform
            // 
            this.btnShapeTransform.Label = "Shape Transform";
            this.btnShapeTransform.Name = "btnShapeTransform";
            this.btnShapeTransform.ScreenTip = "Edit Text";
            this.btnShapeTransform.SuperTip = "Launch the Super Duper Edit Text UI";
            this.btnShapeTransform.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnShapeTransform_Click);
            // 
            // btnTextBlockFormat
            // 
            this.btnTextBlockFormat.Label = "Text Block Format";
            this.btnTextBlockFormat.Name = "btnTextBlockFormat";
            this.btnTextBlockFormat.ScreenTip = "Edit Text";
            this.btnTextBlockFormat.SuperTip = "Launch the Super Duper Edit Text UI";
            this.btnTextBlockFormat.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnTextBlockFormat_Click);
            // 
            // btnTextTransform
            // 
            this.btnTextTransform.Label = "Text Transform";
            this.btnTextTransform.Name = "btnTextTransform";
            this.btnTextTransform.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnTextTransform_Click);
            // 
            // btnThemeProperties
            // 
            this.btnThemeProperties.Label = "Theme Properties";
            this.btnThemeProperties.Name = "btnThemeProperties";
            this.btnThemeProperties.ScreenTip = "Edit Text";
            this.btnThemeProperties.SuperTip = "Launch the Super Duper Edit Text UI";
            this.btnThemeProperties.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnShapeThemeProperties_Click);
            // 
            // rgSSShapeRowSections
            // 
            this.rgSSShapeRowSections.Items.Add(this.btnActions);
            this.rgSSShapeRowSections.Items.Add(this.btnActionTags);
            this.rgSSShapeRowSections.Items.Add(this.btnCharacter);
            this.rgSSShapeRowSections.Items.Add(this.btnConnectionPoints);
            this.rgSSShapeRowSections.Items.Add(this.btnControls);
            this.rgSSShapeRowSections.Items.Add(this.btnGeometry);
            this.rgSSShapeRowSections.Items.Add(this.btnGradientStops);
            this.rgSSShapeRowSections.Items.Add(this.btnParagraph);
            this.rgSSShapeRowSections.Items.Add(this.btnShapeHyperlinks);
            this.rgSSShapeRowSections.Items.Add(this.btnShapeScratch);
            this.rgSSShapeRowSections.Items.Add(this.btnShapeShapeData);
            this.rgSSShapeRowSections.Items.Add(this.btnTabs);
            this.rgSSShapeRowSections.Items.Add(this.btnShapeUserDefinedCells);
            this.rgSSShapeRowSections.Label = "Shape (Rows)";
            this.rgSSShapeRowSections.Name = "rgSSShapeRowSections";
            // 
            // btnActions
            // 
            this.btnActions.Label = "Actions";
            this.btnActions.Name = "btnActions";
            this.btnActions.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnActionsShape_Click);
            // 
            // btnActionTags
            // 
            this.btnActionTags.Label = "Action Tags";
            this.btnActionTags.Name = "btnActionTags";
            this.btnActionTags.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnActionTagsShape_Click);
            // 
            // btnCharacter
            // 
            this.btnCharacter.Label = "X Character";
            this.btnCharacter.Name = "btnCharacter";
            this.btnCharacter.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCharacter_Click);
            // 
            // btnConnectionPoints
            // 
            this.btnConnectionPoints.Label = "Connection Points";
            this.btnConnectionPoints.Name = "btnConnectionPoints";
            this.btnConnectionPoints.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnConnectionPoints_Click);
            // 
            // btnControls
            // 
            this.btnControls.Label = "Controls";
            this.btnControls.Name = "btnControls";
            this.btnControls.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnControls_Click);
            // 
            // btnGeometry
            // 
            this.btnGeometry.Label = "X Geometry";
            this.btnGeometry.Name = "btnGeometry";
            this.btnGeometry.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnGeometry_Click);
            // 
            // btnGradientStops
            // 
            this.btnGradientStops.Label = "X Gradient Stops";
            this.btnGradientStops.Name = "btnGradientStops";
            this.btnGradientStops.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnGradientStops_Click);
            // 
            // btnParagraph
            // 
            this.btnParagraph.Label = "X Paragraph";
            this.btnParagraph.Name = "btnParagraph";
            this.btnParagraph.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnParagraph_Click);
            // 
            // btnShapeHyperlinks
            // 
            this.btnShapeHyperlinks.Label = "Hyperlinks";
            this.btnShapeHyperlinks.Name = "btnShapeHyperlinks";
            this.btnShapeHyperlinks.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnShapeHyperlinks_Click);
            // 
            // btnShapeScratch
            // 
            this.btnShapeScratch.Label = "X Scratch";
            this.btnShapeScratch.Name = "btnShapeScratch";
            this.btnShapeScratch.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnShapeScratch_Click);
            // 
            // btnShapeShapeData
            // 
            this.btnShapeShapeData.Label = "Shape Data";
            this.btnShapeShapeData.Name = "btnShapeShapeData";
            this.btnShapeShapeData.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnShapeShapeData_Click);
            // 
            // btnTabs
            // 
            this.btnTabs.Label = "X Tabs";
            this.btnTabs.Name = "btnTabs";
            this.btnTabs.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnTabs_Click);
            // 
            // btnShapeUserDefinedCells
            // 
            this.btnShapeUserDefinedCells.Label = "User-Defined Cells";
            this.btnShapeUserDefinedCells.Name = "btnShapeUserDefinedCells";
            this.btnShapeUserDefinedCells.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnShapeUserDefinedCells_Click);
            // 
            // tabCustomUI
            // 
            this.tabCustomUI.Groups.Add(this.rgCustomUI);
            this.tabCustomUI.Groups.Add(this.rgDocumentBasePages);
            this.tabCustomUI.Label = "Custom UI";
            this.tabCustomUI.Name = "tabCustomUI";
            // 
            // rgCustomUI
            // 
            this.rgCustomUI.Items.Add(this.btnCommandCockpit);
            this.rgCustomUI.Items.Add(this.btnLinq2Excel);
            this.rgCustomUI.Items.Add(this.btnEditControlRows);
            this.rgCustomUI.Items.Add(this.btnEditParagraph);
            this.rgCustomUI.Items.Add(this.btnEditText);
            this.rgCustomUI.Items.Add(this.btnEditControlPoints);
            this.rgCustomUI.Items.Add(this.btnRenamePages);
            this.rgCustomUI.Items.Add(this.btnDuplicatePage);
            this.rgCustomUI.Items.Add(this.btnMovePages);
            this.rgCustomUI.Items.Add(this.btnCustomUI_Car);
            this.rgCustomUI.Label = "Custom UI";
            this.rgCustomUI.Name = "rgCustomUI";
            // 
            // btnCommandCockpit
            // 
            this.btnCommandCockpit.Label = "Command Cockpit";
            this.btnCommandCockpit.Name = "btnCommandCockpit";
            this.btnCommandCockpit.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCommandCockpit_Click);
            // 
            // btnLinq2Excel
            // 
            this.btnLinq2Excel.Label = "Linq to Excel";
            this.btnLinq2Excel.Name = "btnLinq2Excel";
            this.btnLinq2Excel.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLinqToExcel_Click);
            // 
            // btnEditControlRows
            // 
            this.btnEditControlRows.Label = "Edit Control Rows";
            this.btnEditControlRows.Name = "btnEditControlRows";
            this.btnEditControlRows.ScreenTip = "Edit Text";
            this.btnEditControlRows.SuperTip = "Launch the Super Duper Edit Text UI";
            this.btnEditControlRows.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnEditControlRows_Click);
            // 
            // btnEditParagraph
            // 
            this.btnEditParagraph.Label = "Edit Paragraph";
            this.btnEditParagraph.Name = "btnEditParagraph";
            this.btnEditParagraph.ScreenTip = "Edit Text";
            this.btnEditParagraph.SuperTip = "Launch the Super Duper Edit Text UI";
            this.btnEditParagraph.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnEditParagraph_Click);
            // 
            // btnEditText
            // 
            this.btnEditText.Label = "Edit Text";
            this.btnEditText.Name = "btnEditText";
            this.btnEditText.ScreenTip = "Edit Text";
            this.btnEditText.SuperTip = "Launch the Super Duper Edit Text UI";
            this.btnEditText.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnEditText_Click);
            // 
            // btnEditControlPoints
            // 
            this.btnEditControlPoints.Label = "EditControlPoints";
            this.btnEditControlPoints.Name = "btnEditControlPoints";
            this.btnEditControlPoints.ScreenTip = "EditControlPoints";
            this.btnEditControlPoints.SuperTip = "Launch the Super Duper Edit Control Points";
            this.btnEditControlPoints.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnEditControlPoints_Click);
            // 
            // btnRenamePages
            // 
            this.btnRenamePages.Label = "Rename Pages";
            this.btnRenamePages.Name = "btnRenamePages";
            this.btnRenamePages.SuperTip = "Rename Pages using RegEx";
            this.btnRenamePages.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnRenamePages_Click);
            // 
            // btnDuplicatePage
            // 
            this.btnDuplicatePage.Label = "Duplicate Page";
            this.btnDuplicatePage.Name = "btnDuplicatePage";
            this.btnDuplicatePage.SuperTip = "Duplicate Page";
            this.btnDuplicatePage.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDuplicatePage_Click);
            // 
            // btnMovePages
            // 
            this.btnMovePages.Label = "Move Pages";
            this.btnMovePages.Name = "btnMovePages";
            this.btnMovePages.SuperTip = "Move all Pages listed on current Page to Another Document";
            this.btnMovePages.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnMovePages_Click);
            // 
            // btnCustomUI_Car
            // 
            this.btnCustomUI_Car.Label = "Car";
            this.btnCustomUI_Car.Name = "btnCustomUI_Car";
            this.btnCustomUI_Car.SuperTip = "Load CarMain from Explore module";
            this.btnCustomUI_Car.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCustomUI_Car_Click);
            // 
            // rgDocumentBasePages
            // 
            this.rgDocumentBasePages.Items.Add(this.btnAddArchitectureBasePages);
            this.rgDocumentBasePages.Label = "Document Base Pages";
            this.rgDocumentBasePages.Name = "rgDocumentBasePages";
            // 
            // btnAddArchitectureBasePages
            // 
            this.btnAddArchitectureBasePages.Label = "+ Architecture Base Pages";
            this.btnAddArchitectureBasePages.Name = "btnAddArchitectureBasePages";
            this.btnAddArchitectureBasePages.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAddArchitectureBasePages_Click);
            // 
            // tabUILaunch
            // 
            this.tabUILaunch.Groups.Add(this.rgUILaunch);
            this.tabUILaunch.Groups.Add(this.rgWPFUI);
            this.tabUILaunch.Groups.Add(this.rgMVVMExamples);
            this.tabUILaunch.Label = "UI Launch Approaches";
            this.tabUILaunch.Name = "tabUILaunch";
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
            this.rgWPFUI.Items.Add(this.btnDxDockLayoutManager);
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
            // btnDxDockLayoutManager
            // 
            this.btnDxDockLayoutManager.Label = "DxDockLayoutManager";
            this.btnDxDockLayoutManager.Name = "btnDxDockLayoutManager";
            this.btnDxDockLayoutManager.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDxDockLayoutManagerControl_Click);
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
            this.RibbonType = "Microsoft.Visio.Drawing";
            this.Tabs.Add(this.tab1);
            this.Tabs.Add(this.tabSupportTools);
            this.Tabs.Add(this.tabShapeSheet);
            this.Tabs.Add(this.tabCustomUI);
            this.Tabs.Add(this.tabUILaunch);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.tabSupportTools.ResumeLayout(false);
            this.tabSupportTools.PerformLayout();
            this.rgDocumentActions.ResumeLayout(false);
            this.rgDocumentActions.PerformLayout();
            this.rgPageActions.ResumeLayout(false);
            this.rgPageActions.PerformLayout();
            this.rgLayerActions.ResumeLayout(false);
            this.rgLayerActions.PerformLayout();
            this.rgShapeActions.ResumeLayout(false);
            this.rgShapeActions.PerformLayout();
            this.rgDebug.ResumeLayout(false);
            this.rgDebug.PerformLayout();
            this.grpHelp.ResumeLayout(false);
            this.grpHelp.PerformLayout();
            this.tabShapeSheet.ResumeLayout(false);
            this.tabShapeSheet.PerformLayout();
            this.rgSSDocumentObjectSections.ResumeLayout(false);
            this.rgSSDocumentObjectSections.PerformLayout();
            this.rgSSDocumentRowSections.ResumeLayout(false);
            this.rgSSDocumentRowSections.PerformLayout();
            this.rgSSPageObjectSections.ResumeLayout(false);
            this.rgSSPageObjectSections.PerformLayout();
            this.rgSSPPageRowSections.ResumeLayout(false);
            this.rgSSPPageRowSections.PerformLayout();
            this.rgSSShapeObjectSections.ResumeLayout(false);
            this.rgSSShapeObjectSections.PerformLayout();
            this.rgSSShapeRowSections.ResumeLayout(false);
            this.rgSSShapeRowSections.PerformLayout();
            this.tabCustomUI.ResumeLayout(false);
            this.tabCustomUI.PerformLayout();
            this.rgCustomUI.ResumeLayout(false);
            this.rgCustomUI.PerformLayout();
            this.rgDocumentBasePages.ResumeLayout(false);
            this.rgDocumentBasePages.PerformLayout();
            this.tabUILaunch.ResumeLayout(false);
            this.tabUILaunch.PerformLayout();
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
        private Microsoft.Office.Tools.Ribbon.RibbonTab tabSupportTools;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup rgDebug;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDebugWindow;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnWatchWindow;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox chkEnableAppEvents;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox chkDisplayEvents;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpHelp;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAddInInfo;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDeveloperMode;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox chkDisplayChattyEvents;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup rgPageActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAddTableOfContents;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGetDocumentInfo;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGetStencilInfo;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup rgDocumentActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGetShapeInfo;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGetPageInfo;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAddFooter;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAddDefaultLayers;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRemoveLayers;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAddIsPageName;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUpdatePageNameShapes;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAddNavigationLinks;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGetApplicationInfo;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup rgShapeActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAddNavLinks;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAddHyperLink;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPrintPage;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSavePage;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup rgLayerActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox cmbLayers;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPageOn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPageOff;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAllPageOn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAllPageOff;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLayerManager;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAddTextControl;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn0PtMargins;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnMakeLinkableMaster;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAddColorSupport;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAddIDSupport;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAddIDAndTextSupport;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSortAllPages;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDisplayPageNames;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSyncPageNames;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSyncPageNamesPage;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUpdatePageNameShapesPage;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn1PtMargins;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn2PtMargins;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnMoveToBackgroundLayer;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLockBackground;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUnlockBackground;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAddHeader;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAutoSizePagesOn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAutoSizePagesOff;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAutoSizePageOn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAutoSizePageOff;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPrintPages;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDeletePages;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSavePages;
        private Microsoft.Office.Tools.Ribbon.RibbonTab tabUILaunch;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup rgUILaunch;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnWindowHostLocal;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup rgWPFUI;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLaunchCylon;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnThemedWindowHostModeless;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLaunchCylon2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPrismRegionTest;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDxWindowHost;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnWindowHostVNC;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDxDockLayoutManager;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnThemedWindowHostModal;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDxLayoutControl;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDxDockLayoutControl;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnShapeTransform;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnProtection;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnMiscelleaneous;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLineFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFillFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnTextBlockFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnEvents;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnImageProperties;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGlueInfo;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnShapeLayout;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGradientProperties;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAdditionalEffectProperties;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnBevelProperties;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnThemeProperties;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnQuickStyle;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnChangeShapeBehavior;
        private Microsoft.Office.Tools.Ribbon.RibbonTab tabShapeSheet;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup rgSSShapeObjectSections;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn3DRotationProperties;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup rgSSShapeRowSections;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn1DEndpoints;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnShapeUserDefinedCells;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLayerMembership;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGroupProperties;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnTextTransform;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup rgSSDocumentRowSections;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup rgSSDocumentObjectSections;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup rgSSPPageRowSections;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup rgSSPageObjectSections;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDocumentUserDefinedCells;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDocumentHyperlinks;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPageUserDefinedCells;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPageHyperlinks;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnShapeHyperlinks;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDocumentScratch;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDocumentShapeData;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDocumentProperties;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPageScratch;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPageShapeData;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPageProperties;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLayers;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPageLayout;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPrintProperties;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRulerAndGrid;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnShapeScratch;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnShapeShapeData;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnActionTags;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCharacter;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnConnectionPoints;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnControls;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGeometry;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGradientStops;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnParagraph;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnTabs;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPageActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPageActionTags;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPageThemeProperties;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup rgMVVMExamples;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnVNC_MVVM_VAVM1st;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabCustomUI;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup rgCustomUI;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLinq2Excel;

        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnEditControlRows;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnEditParagraph;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnEditText;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAddDefaultLayers_Page;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRemoveLayers_Page;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnEditControlPoints;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnVNC_MVVM_VA1st;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnVNC_MVVM_VAVM1stDI;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCommandCockpit;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRenamePages;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDuplicatePage;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnMovePages;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnVNC_MVVM_VB1st;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnVNC_MVVM_VC11st;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnVNC_MVVM_VC21st;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup rgDocumentBasePages;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAddArchitectureBasePages;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCustomUI_Car;
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
