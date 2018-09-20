namespace CrmOutlookAddIn2013
{
    partial class CrmSyncRibbonDesign : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public CrmSyncRibbonDesign()
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
            this.tabCrmSync = this.Factory.CreateRibbonTab();
            this.groupSynch = this.Factory.CreateRibbonGroup();
            this.btnDoSync = this.Factory.CreateRibbonButton();
            this.tabCrmSync.SuspendLayout();
            this.groupSynch.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabCrmSync
            // 
            this.tabCrmSync.Groups.Add(this.groupSynch);
            this.tabCrmSync.Label = "CRM";
            this.tabCrmSync.Name = "tabCrmSync";
            // 
            // groupSynch
            // 
            this.groupSynch.Items.Add(this.btnDoSync);
            this.groupSynch.Label = "Integration";
            this.groupSynch.Name = "groupSynch";
            // 
            // btnDoSync
            // 
            this.btnDoSync.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnDoSync.Image = global::CrmOutlookAddIn2013.Properties.Resources.sync3;
            this.btnDoSync.Label = "Synchronize appointments with AX";
            this.btnDoSync.Name = "btnDoSync";
            this.btnDoSync.ScreenTip = "Start synchronizing appointments with CRM activities.";
            this.btnDoSync.ShowImage = true;
            this.btnDoSync.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDoSync_Click);
            // 
            // CrmSyncRibbonDesign
            // 
            this.Name = "CrmSyncRibbonDesign";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.tabCrmSync);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.CrmSyncRibbonDesign_Load);
            this.tabCrmSync.ResumeLayout(false);
            this.tabCrmSync.PerformLayout();
            this.groupSynch.ResumeLayout(false);
            this.groupSynch.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabCrmSync;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupSynch;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDoSync;
    }

    partial class ThisRibbonCollection
    {
        internal CrmSyncRibbonDesign CrmSyncRibbonDesign
        {
            get { return this.GetRibbon<CrmSyncRibbonDesign>(); }
        }
    }
}
