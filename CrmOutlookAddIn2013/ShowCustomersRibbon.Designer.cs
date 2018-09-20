namespace CrmOutlookAddIn2013
{
    partial class ShowCustomersRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public ShowCustomersRibbon()
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
            this.Customers = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.ShowCustomers = this.Factory.CreateRibbonButton();
            this.Customers.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // Customers
            // 
            this.Customers.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.Customers.ControlId.OfficeId = "TabAppointment";
            this.Customers.Groups.Add(this.group1);
            this.Customers.Label = "TabAppointment";
            this.Customers.Name = "Customers";
            // 
            // group1
            // 
            this.group1.Items.Add(this.ShowCustomers);
            this.group1.Label = "SelectGrp";
            this.group1.Name = "group1";
            // 
            // ShowCustomers
            // 
            this.ShowCustomers.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ShowCustomers.Image = global::CrmOutlookAddIn2013.Properties.Resources.add_contact;
            this.ShowCustomers.Label = "Add Business Relation";
            this.ShowCustomers.Name = "ShowCustomers";
            this.ShowCustomers.ScreenTip = "Select business relation to be added to appointment.";
            this.ShowCustomers.ShowImage = true;
            this.ShowCustomers.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ShowCustomers_Click);
            // 
            // ShowCustomersRibbon
            // 
            this.Name = "ShowCustomersRibbon";
            this.RibbonType = "Microsoft.Outlook.Appointment";
            this.Tabs.Add(this.Customers);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.ShowCustomersRibbon_Load);
            this.Customers.ResumeLayout(false);
            this.Customers.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab Customers;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ShowCustomers;
    }

    partial class ThisRibbonCollection
    {
        internal ShowCustomersRibbon ShowCustomersRibbon
        {
            get { return this.GetRibbon<ShowCustomersRibbon>(); }
        }
    }
}
