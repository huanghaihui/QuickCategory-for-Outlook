namespace TagCloud4
{
    partial class TagCloud : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;
        
        public TagCloud()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(TagCloud));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.show_tagCloud = this.Factory.CreateRibbonGroup();
            this.showTagCloud = this.Factory.CreateRibbonButton();
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.tab1.SuspendLayout();
            this.show_tagCloud.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.show_tagCloud);
            this.tab1.Label = "AddIns";
            this.tab1.Name = "tab1";
            // 
            // show_tagCloud
            // 
            this.show_tagCloud.Items.Add(this.showTagCloud);
            this.show_tagCloud.Label = "CatagoriesManager";
            this.show_tagCloud.Name = "show_tagCloud";
            // 
            // showTagCloud
            // 
            this.showTagCloud.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.showTagCloud.Image = ((System.Drawing.Image)(resources.GetObject("showTagCloud.Image")));
            this.showTagCloud.Label = "Hide/Display Catagory";
            this.showTagCloud.Name = "showTagCloud";
            this.showTagCloud.ShowImage = true;
            this.showTagCloud.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // TagCloud
            // 
            this.Name = "TagCloud";
            this.RibbonType = "Microsoft.Outlook.Explorer, Microsoft.Outlook.Mail.Read";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.TagCloud_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.show_tagCloud.ResumeLayout(false);
            this.show_tagCloud.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup show_tagCloud;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton showTagCloud;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
    }

    partial class ThisRibbonCollection
    {
        internal TagCloud TagCloud
        {
            get { return this.GetRibbon<TagCloud>(); }
        }
    }
}
