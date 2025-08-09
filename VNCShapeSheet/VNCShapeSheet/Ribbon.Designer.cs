namespace VNCShapeSheet
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
            this.rtVisioAddInTemplate = this.Factory.CreateRibbonTab();
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
            this.btnAlignment = this.Factory.CreateRibbonButton();
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
            this.rgDebug = this.Factory.CreateRibbonGroup();
            this.btnDebugWindow = this.Factory.CreateRibbonButton();
            this.btnWatchWindow = this.Factory.CreateRibbonButton();
            this.rcbEnableAppEvents = this.Factory.CreateRibbonCheckBox();
            this.rcbDisplayEvents = this.Factory.CreateRibbonCheckBox();
            this.rcbDisplayChattyEvents = this.Factory.CreateRibbonCheckBox();
            this.rcbDeveloperUIMode = this.Factory.CreateRibbonCheckBox();
            this.grpHelp = this.Factory.CreateRibbonGroup();
            this.btnAddInInfo = this.Factory.CreateRibbonButton();
            this.btnDeveloperMode = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.rtVisioAddInTemplate.SuspendLayout();
            this.rgSSDocumentObjectSections.SuspendLayout();
            this.rgSSDocumentRowSections.SuspendLayout();
            this.rgSSPageObjectSections.SuspendLayout();
            this.rgSSPPageRowSections.SuspendLayout();
            this.rgSSShapeObjectSections.SuspendLayout();
            this.rgSSShapeRowSections.SuspendLayout();
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
            // rtVisioAddInTemplate
            // 
            this.rtVisioAddInTemplate.Groups.Add(this.rgSSDocumentObjectSections);
            this.rtVisioAddInTemplate.Groups.Add(this.rgSSDocumentRowSections);
            this.rtVisioAddInTemplate.Groups.Add(this.rgSSPageObjectSections);
            this.rtVisioAddInTemplate.Groups.Add(this.rgSSPPageRowSections);
            this.rtVisioAddInTemplate.Groups.Add(this.rgSSShapeObjectSections);
            this.rtVisioAddInTemplate.Groups.Add(this.rgSSShapeRowSections);
            this.rtVisioAddInTemplate.Groups.Add(this.rgDebug);
            this.rtVisioAddInTemplate.Groups.Add(this.grpHelp);
            this.rtVisioAddInTemplate.Label = "VNCShapeSheet";
            this.rtVisioAddInTemplate.Name = "rtVisioAddInTemplate";
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
            this.rgSSPPageRowSections.Items.Add(this.btnPageActions);
            this.rgSSPPageRowSections.Items.Add(this.btnPageActionTags);
            this.rgSSPPageRowSections.Items.Add(this.btnPageHyperlinks);
            this.rgSSPPageRowSections.Items.Add(this.btnLayers);
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
            this.btnPageActions.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnPageActions_Click);
            // 
            // btnPageActionTags
            // 
            this.btnPageActionTags.Label = "Action Tags";
            this.btnPageActionTags.Name = "btnPageActionTags";
            this.btnPageActionTags.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnPageActionTags_Click);
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
            this.rgSSShapeObjectSections.Items.Add(this.btnAlignment);
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
            // btnAlignment
            // 
            this.btnAlignment.Label = "Alignment";
            this.btnAlignment.Name = "btnAlignment";
            this.btnAlignment.ScreenTip = "Edit Text";
            this.btnAlignment.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAlignment_Click);
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
            this.btnThemeProperties.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnThemeProperties_Click);
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
            this.btnActions.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnActions_Click);
            // 
            // btnActionTags
            // 
            this.btnActionTags.Label = "Action Tags";
            this.btnActionTags.Name = "btnActionTags";
            this.btnActionTags.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnActionTags_Click);
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
            // rgDebug
            // 
            this.rgDebug.Items.Add(this.btnDebugWindow);
            this.rgDebug.Items.Add(this.btnWatchWindow);
            this.rgDebug.Items.Add(this.rcbEnableAppEvents);
            this.rgDebug.Items.Add(this.rcbDisplayEvents);
            this.rgDebug.Items.Add(this.rcbDisplayChattyEvents);
            this.rgDebug.Items.Add(this.rcbDeveloperUIMode);
            this.rgDebug.Label = "Debug";
            this.rgDebug.Name = "rgDebug";
            this.rgDebug.Visible = false;
            // 
            // btnDebugWindow
            // 
            this.btnDebugWindow.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnDebugWindow.Image = global::VNCShapeSheet.Properties.Resources.DebugWindow;
            this.btnDebugWindow.Label = "Debug Window";
            this.btnDebugWindow.Name = "btnDebugWindow";
            this.btnDebugWindow.ShowImage = true;
            this.btnDebugWindow.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDebugWindow_Click);
            // 
            // btnWatchWindow
            // 
            this.btnWatchWindow.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnWatchWindow.Image = global::VNCShapeSheet.Properties.Resources.WatchWindow;
            this.btnWatchWindow.Label = "Watch Window";
            this.btnWatchWindow.Name = "btnWatchWindow";
            this.btnWatchWindow.ShowImage = true;
            this.btnWatchWindow.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnWatchWindow_Click);
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
            // rcbDeveloperUIMode
            // 
            this.rcbDeveloperUIMode.Label = "DeveloperUIMode";
            this.rcbDeveloperUIMode.Name = "rcbDeveloperUIMode";
            this.rcbDeveloperUIMode.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnToggleDeveloperUIMode_Click);
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
            this.btnAddInInfo.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDisplayAddInInfo_Click);
            // 
            // btnDeveloperMode
            // 
            this.btnDeveloperMode.Label = "Developer Mode";
            this.btnDeveloperMode.Name = "btnDeveloperMode";
            this.btnDeveloperMode.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnToggleDeveloperMode_Click);
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.Visio.Drawing";
            this.Tabs.Add(this.tab1);
            this.Tabs.Add(this.rtVisioAddInTemplate);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.rtVisioAddInTemplate.ResumeLayout(false);
            this.rtVisioAddInTemplate.PerformLayout();
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
            this.rgDebug.ResumeLayout(false);
            this.rgDebug.PerformLayout();
            this.grpHelp.ResumeLayout(false);
            this.grpHelp.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab rtVisioAddInTemplate;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup rgDebug;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDebugWindow;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnWatchWindow;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox rcbEnableAppEvents;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox rcbDisplayEvents;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox rcbDisplayChattyEvents;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox rcbDeveloperUIMode;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpHelp;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAddInInfo;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDeveloperMode;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup rgSSDocumentObjectSections;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDocumentProperties;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup rgSSDocumentRowSections;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDocumentHyperlinks;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDocumentScratch;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDocumentShapeData;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDocumentUserDefinedCells;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup rgSSPageObjectSections;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPageLayout;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPageProperties;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPrintProperties;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRulerAndGrid;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPageThemeProperties;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup rgSSPPageRowSections;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLayers;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPageActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPageActionTags;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPageHyperlinks;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPageScratch;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPageShapeData;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPageUserDefinedCells;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup rgSSShapeObjectSections;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn1DEndpoints;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn3DRotationProperties;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAdditionalEffectProperties;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAlignment;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnBevelProperties;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnChangeShapeBehavior;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnEvents;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFillFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGlueInfo;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGradientProperties;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGroupProperties;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnImageProperties;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLayerMembership;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLineFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnMiscelleaneous;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnProtection;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnQuickStyle;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnShapeLayout;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnShapeTransform;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnTextBlockFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnTextTransform;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnThemeProperties;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup rgSSShapeRowSections;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnActions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnActionTags;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCharacter;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnConnectionPoints;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnControls;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGeometry;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGradientStops;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnParagraph;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnShapeHyperlinks;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnShapeScratch;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnShapeShapeData;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnTabs;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnShapeUserDefinedCells;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
