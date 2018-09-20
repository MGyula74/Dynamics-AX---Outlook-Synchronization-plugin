namespace CrmOutlookAddIn2013
{
    partial class OutlookCrmSyncForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btnClose = new System.Windows.Forms.Button();
            this.progressToCrm = new System.Windows.Forms.ProgressBar();
            this.lblOutlook2Crm = new System.Windows.Forms.Label();
            this.btnStart = new System.Windows.Forms.Button();
            this.progressFromCrm = new System.Windows.Forms.ProgressBar();
            this.lblCrm2Outlook = new System.Windows.Forms.Label();
            this.SyncLogBox = new System.Windows.Forms.GroupBox();
            this.syncLogText = new System.Windows.Forms.RichTextBox();
            this.groupBox1.SuspendLayout();
            this.SyncLogBox.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btnClose);
            this.groupBox1.Controls.Add(this.progressToCrm);
            this.groupBox1.Controls.Add(this.lblOutlook2Crm);
            this.groupBox1.Controls.Add(this.btnStart);
            this.groupBox1.Controls.Add(this.progressFromCrm);
            this.groupBox1.Controls.Add(this.lblCrm2Outlook);
            this.groupBox1.Location = new System.Drawing.Point(12, 2);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(568, 85);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(487, 49);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 5;
            this.btnClose.Text = "Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // progressToCrm
            // 
            this.progressToCrm.Location = new System.Drawing.Point(111, 48);
            this.progressToCrm.Name = "progressToCrm";
            this.progressToCrm.Size = new System.Drawing.Size(262, 23);
            this.progressToCrm.Style = System.Windows.Forms.ProgressBarStyle.Continuous;
            this.progressToCrm.TabIndex = 1;
            // 
            // lblOutlook2Crm
            // 
            this.lblOutlook2Crm.AutoSize = true;
            this.lblOutlook2Crm.Location = new System.Drawing.Point(6, 54);
            this.lblOutlook2Crm.Name = "lblOutlook2Crm";
            this.lblOutlook2Crm.Size = new System.Drawing.Size(89, 13);
            this.lblOutlook2Crm.TabIndex = 5;
            this.lblOutlook2Crm.Text = "Outlook ---> CRM";
            // 
            // btnStart
            // 
            this.btnStart.Location = new System.Drawing.Point(487, 16);
            this.btnStart.Name = "btnStart";
            this.btnStart.Size = new System.Drawing.Size(75, 23);
            this.btnStart.TabIndex = 0;
            this.btnStart.Text = "Start";
            this.btnStart.UseVisualStyleBackColor = true;
            this.btnStart.Click += new System.EventHandler(this.btnStart_Click);
            // 
            // progressFromCrm
            // 
            this.progressFromCrm.Location = new System.Drawing.Point(111, 16);
            this.progressFromCrm.Name = "progressFromCrm";
            this.progressFromCrm.Size = new System.Drawing.Size(262, 23);
            this.progressFromCrm.Style = System.Windows.Forms.ProgressBarStyle.Continuous;
            this.progressFromCrm.TabIndex = 4;
            // 
            // lblCrm2Outlook
            // 
            this.lblCrm2Outlook.AutoSize = true;
            this.lblCrm2Outlook.Location = new System.Drawing.Point(6, 21);
            this.lblCrm2Outlook.Name = "lblCrm2Outlook";
            this.lblCrm2Outlook.Size = new System.Drawing.Size(89, 13);
            this.lblCrm2Outlook.TabIndex = 3;
            this.lblCrm2Outlook.Text = "CRM ---> Outlook";
            // 
            // SyncLogBox
            // 
            this.SyncLogBox.Controls.Add(this.syncLogText);
            this.SyncLogBox.Location = new System.Drawing.Point(12, 93);
            this.SyncLogBox.Name = "SyncLogBox";
            this.SyncLogBox.Size = new System.Drawing.Size(568, 180);
            this.SyncLogBox.TabIndex = 1;
            this.SyncLogBox.TabStop = false;
            this.SyncLogBox.Text = "Log";
            // 
            // syncLogText
            // 
            this.syncLogText.Location = new System.Drawing.Point(9, 20);
            this.syncLogText.Name = "syncLogText";
            this.syncLogText.ReadOnly = true;
            this.syncLogText.Size = new System.Drawing.Size(553, 154);
            this.syncLogText.TabIndex = 0;
            this.syncLogText.Text = "";
            // 
            // OutlookCrmSyncForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(594, 285);
            this.Controls.Add(this.SyncLogBox);
            this.Controls.Add(this.groupBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximumSize = new System.Drawing.Size(600, 310);
            this.MinimumSize = new System.Drawing.Size(600, 310);
            this.Name = "OutlookCrmSyncForm";
            this.Text = "Outlook - CRM Synchronization v1.0.0.0";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.SyncLogBox.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox SyncLogBox;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.ProgressBar progressToCrm;
        private System.Windows.Forms.Label lblOutlook2Crm;
        private System.Windows.Forms.Button btnStart;
        private System.Windows.Forms.ProgressBar progressFromCrm;
        private System.Windows.Forms.Label lblCrm2Outlook;
        private System.Windows.Forms.RichTextBox syncLogText;
    }
}