namespace UrduLanguageTools
{
    partial class LanguageToolsRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public LanguageToolsRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

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
            this.tabMain = this.Factory.CreateRibbonTab();
            this.grpTools = this.Factory.CreateRibbonGroup();
            this.btnRefresh = this.Factory.CreateRibbonButton();
            this.btnRemoveMultipleSpaces = this.Factory.CreateRibbonButton();
            this.grpGhazal = this.Factory.CreateRibbonGroup();
            this.btnGhazalPaste = this.Factory.CreateRibbonButton();
            this.btnGhazalFormat = this.Factory.CreateRibbonButton();
            this.drpGhazalStyle = this.Factory.CreateRibbonDropDown();
            this.cbGhazalAddToTableOfContents = this.Factory.CreateRibbonCheckBox();
            this.drpGhazalLinesPerVerse = this.Factory.CreateRibbonDropDown();
            this.tabMain.SuspendLayout();
            this.grpTools.SuspendLayout();
            this.grpGhazal.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabMain
            // 
            this.tabMain.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabMain.Groups.Add(this.grpTools);
            this.tabMain.Groups.Add(this.grpGhazal);
            this.tabMain.Label = "TabAddIns";
            this.tabMain.Name = "tabMain";
            // 
            // grpTools
            // 
            this.grpTools.Items.Add(this.btnRefresh);
            this.grpTools.Items.Add(this.btnRemoveMultipleSpaces);
            this.grpTools.Label = "Tools";
            this.grpTools.Name = "grpTools";
            // 
            // btnRefresh
            // 
            this.btnRefresh.Label = "Refresh";
            this.btnRefresh.Name = "btnRefresh";
            this.btnRefresh.OfficeImageId = "Repeat";
            this.btnRefresh.ShowImage = true;
            // 
            // btnRemoveMultipleSpaces
            // 
            this.btnRemoveMultipleSpaces.Label = "Remove Multiple Spaces";
            this.btnRemoveMultipleSpaces.Name = "btnRemoveMultipleSpaces";
            this.btnRemoveMultipleSpaces.OfficeImageId = "Clear";
            this.btnRemoveMultipleSpaces.ShowImage = true;
            this.btnRemoveMultipleSpaces.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnRemoveMultipleSpaces_Click);
            // 
            // grpGhazal
            // 
            this.grpGhazal.Items.Add(this.btnGhazalPaste);
            this.grpGhazal.Items.Add(this.btnGhazalFormat);
            this.grpGhazal.Items.Add(this.drpGhazalStyle);
            this.grpGhazal.Items.Add(this.drpGhazalLinesPerVerse);
            this.grpGhazal.Items.Add(this.cbGhazalAddToTableOfContents);
            this.grpGhazal.Label = "Ghazal";
            this.grpGhazal.Name = "grpGhazal";
            // 
            // btnGhazalPaste
            // 
            this.btnGhazalPaste.Label = "Paste";
            this.btnGhazalPaste.Name = "btnGhazalPaste";
            this.btnGhazalPaste.OfficeImageId = "Paste";
            this.btnGhazalPaste.ShowImage = true;
            this.btnGhazalPaste.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnGhazalPaste_Click);
            // 
            // btnGhazalFormat
            // 
            this.btnGhazalFormat.Label = "Format";
            this.btnGhazalFormat.Name = "btnGhazalFormat";
            this.btnGhazalFormat.OfficeImageId = "FormatPainter";
            this.btnGhazalFormat.ShowImage = true;
            this.btnGhazalFormat.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnGhazalFormat_Click);
            // 
            // drpGhazalStyle
            // 
            this.drpGhazalStyle.Label = "Text Style";
            this.drpGhazalStyle.Name = "drpGhazalStyle";
            this.drpGhazalStyle.OfficeImageId = "QuickStylesGallery";
            this.drpGhazalStyle.ShowImage = true;
            this.drpGhazalStyle.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.drpGhazalParagraphStyle_SelectionChanged);
            // 
            // cbGhazalAddToTableOfContents
            // 
            this.cbGhazalAddToTableOfContents.Label = "Add To Table of Contents";
            this.cbGhazalAddToTableOfContents.Name = "cbGhazalAddToTableOfContents";
            // 
            // drpGhazalLinesPerVerse
            // 
            this.drpGhazalLinesPerVerse.Label = "Lines Per Verse";
            this.drpGhazalLinesPerVerse.Name = "drpGhazalLinesPerVerse";
            this.drpGhazalLinesPerVerse.OfficeImageId = "Numbering";
            this.drpGhazalLinesPerVerse.ShowImage = true;
            // 
            // LanguageToolsRibbon
            // 
            this.Name = "LanguageToolsRibbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tabMain);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.LanguageToolsRibbon_Load);
            this.tabMain.ResumeLayout(false);
            this.tabMain.PerformLayout();
            this.grpTools.ResumeLayout(false);
            this.grpTools.PerformLayout();
            this.grpGhazal.ResumeLayout(false);
            this.grpGhazal.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabMain;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpGhazal;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGhazalPaste;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown drpGhazalStyle;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpTools;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRefresh;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGhazalFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRemoveMultipleSpaces;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbGhazalAddToTableOfContents;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown drpGhazalLinesPerVerse;
    }

    partial class ThisRibbonCollection
    {
        internal LanguageToolsRibbon LanguageToolsRibbon
        {
            get { return this.GetRibbon<LanguageToolsRibbon>(); }
        }
    }
}
