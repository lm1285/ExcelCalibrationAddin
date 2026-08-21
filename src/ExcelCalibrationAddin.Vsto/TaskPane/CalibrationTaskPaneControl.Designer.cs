namespace ExcelCalibrationAddin.Vsto.TaskPane
{
    partial class CalibrationTaskPaneControl
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.Panel panelRoot;
        private System.Windows.Forms.Panel panelHeader;
        private System.Windows.Forms.Label lblHeaderTitle;
        private System.Windows.Forms.Label lblHeaderSubtitle;
        private System.Windows.Forms.Panel panelBody;
        private System.Windows.Forms.Panel panelOverview;
        private System.Windows.Forms.Label lblRemoteCaption;
        private System.Windows.Forms.Label lblRemoteValue;
        private System.Windows.Forms.Panel panelRules;
        private System.Windows.Forms.Label lblRulesTitle;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }

            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.panelRoot = new System.Windows.Forms.Panel();
            this.panelBody = new System.Windows.Forms.Panel();
            this.panelRules = new System.Windows.Forms.Panel();
            this.lblRulesTitle = new System.Windows.Forms.Label();
            this.panelOverview = new System.Windows.Forms.Panel();
            this.lblRemoteValue = new System.Windows.Forms.Label();
            this.lblRemoteCaption = new System.Windows.Forms.Label();
            this.panelHeader = new System.Windows.Forms.Panel();
            this.lblHeaderSubtitle = new System.Windows.Forms.Label();
            this.lblHeaderTitle = new System.Windows.Forms.Label();
            this.panelRoot.SuspendLayout();
            this.panelBody.SuspendLayout();
            this.panelRules.SuspendLayout();
            this.panelOverview.SuspendLayout();
            this.panelHeader.SuspendLayout();
            this.SuspendLayout();
            // 
            // panelRoot
            // 
            this.panelRoot.Controls.Add(this.panelBody);
            this.panelRoot.Controls.Add(this.panelHeader);
            this.panelRoot.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelRoot.Location = new System.Drawing.Point(0, 0);
            this.panelRoot.Name = "panelRoot";
            this.panelRoot.Size = new System.Drawing.Size(420, 780);
            this.panelRoot.TabIndex = 0;
            // 
            // panelBody
            // 
            this.panelBody.AutoScroll = true;
            this.panelBody.Controls.Add(this.panelRules);
            this.panelBody.Controls.Add(this.panelOverview);
            this.panelBody.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelBody.Location = new System.Drawing.Point(0, 86);
            this.panelBody.Name = "panelBody";
            this.panelBody.Size = new System.Drawing.Size(420, 694);
            this.panelBody.TabIndex = 1;
            // 
            // panelRules
            // 
            this.panelRules.Controls.Add(this.lblRulesTitle);
            this.panelRules.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelRules.Location = new System.Drawing.Point(0, 116);
            this.panelRules.Name = "panelRules";
            this.panelRules.Size = new System.Drawing.Size(420, 390);
            this.panelRules.TabIndex = 3;
            // 
            // lblRulesTitle
            // 
            this.lblRulesTitle.AutoSize = false;
            this.lblRulesTitle.Location = new System.Drawing.Point(16, 14);
            this.lblRulesTitle.Name = "lblRulesTitle";
            this.lblRulesTitle.Size = new System.Drawing.Size(380, 24);
            this.lblRulesTitle.TabIndex = 0;
            // 
            // panelOverview
            // 
            this.panelOverview.Controls.Add(this.lblRemoteValue);
            this.panelOverview.Controls.Add(this.lblRemoteCaption);
            this.panelOverview.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelOverview.Location = new System.Drawing.Point(0, 0);
            this.panelOverview.Name = "panelOverview";
            this.panelOverview.Size = new System.Drawing.Size(420, 116);
            this.panelOverview.TabIndex = 1;
            // 
            // overview labels
            // 
            this.lblRemoteValue.Location = new System.Drawing.Point(16, 144);
            this.lblRemoteValue.Name = "lblRemoteValue";
            this.lblRemoteValue.Size = new System.Drawing.Size(380, 24);
            this.lblRemoteValue.TabIndex = 5;
            this.lblRemoteCaption.Location = new System.Drawing.Point(16, 124);
            this.lblRemoteCaption.Name = "lblRemoteCaption";
            this.lblRemoteCaption.Size = new System.Drawing.Size(380, 18);
            this.lblRemoteCaption.TabIndex = 4;
            // 
            // panelHeader
            // 
            this.panelHeader.Controls.Add(this.lblHeaderSubtitle);
            this.panelHeader.Controls.Add(this.lblHeaderTitle);
            this.panelHeader.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelHeader.Location = new System.Drawing.Point(0, 0);
            this.panelHeader.Name = "panelHeader";
            this.panelHeader.Size = new System.Drawing.Size(420, 86);
            this.panelHeader.TabIndex = 0;
            // 
            // header labels
            // 
            this.lblHeaderSubtitle.Location = new System.Drawing.Point(16, 50);
            this.lblHeaderSubtitle.Name = "lblHeaderSubtitle";
            this.lblHeaderSubtitle.Size = new System.Drawing.Size(380, 20);
            this.lblHeaderSubtitle.TabIndex = 1;
            this.lblHeaderTitle.Location = new System.Drawing.Point(16, 16);
            this.lblHeaderTitle.Name = "lblHeaderTitle";
            this.lblHeaderTitle.Size = new System.Drawing.Size(380, 30);
            this.lblHeaderTitle.TabIndex = 0;
            // 
            // CalibrationTaskPaneControl
            // 
            this.Controls.Add(this.panelRoot);
            this.Name = "CalibrationTaskPaneControl";
            this.Size = new System.Drawing.Size(420, 780);
            this.panelRoot.ResumeLayout(false);
            this.panelBody.ResumeLayout(false);
            this.panelBody.PerformLayout();
            this.panelRules.ResumeLayout(false);
            this.panelOverview.ResumeLayout(false);
            this.panelHeader.ResumeLayout(false);
            this.ResumeLayout(false);
        }
    }
}
