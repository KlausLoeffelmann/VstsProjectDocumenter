namespace VstsProjectDocumenter
{
    partial class ProjectDocumenterRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public ProjectDocumenterRibbon()
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
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.documentProjectbutton = this.Factory.CreateRibbonButton();
            this.button1 = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
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
            this.group1.Items.Add(this.documentProjectbutton);
            this.group1.Items.Add(this.button1);
            this.group1.Label = "Projects";
            this.group1.Name = "group1";
            // 
            // documentProjectbutton
            // 
            this.documentProjectbutton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.documentProjectbutton.Image = global::VstsProjectDocumenter.Properties.Resources.Reports_collapsed_12995;
            this.documentProjectbutton.Label = "Document Project";
            this.documentProjectbutton.Name = "documentProjectbutton";
            this.documentProjectbutton.ShowImage = true;
            this.documentProjectbutton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DocumentProjectbutton_Click);
            // 
            // button1
            // 
            this.button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button1.Image = global::VstsProjectDocumenter.Properties.Resources.Arrow_RedoRetry_16xLG_color;
            this.button1.Label = "Reset Connection";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
            // 
            // ProjectDocumenterRibbon
            // 
            this.Name = "ProjectDocumenterRibbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton documentProjectbutton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
    }

    partial class ThisRibbonCollection
    {
        internal ProjectDocumenterRibbon ProjectDocumenterRibbon
        {
            get { return this.GetRibbon<ProjectDocumenterRibbon>(); }
        }
    }
}
