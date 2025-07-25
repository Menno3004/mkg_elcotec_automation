namespace Mkg_Elcotec_Automation.Forms
{
    partial class Elcotec
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.runToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.mkgTestToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.mkgMenuStrip = new System.Windows.Forms.MenuStrip();
            this.MkgMenuButton = new System.Windows.Forms.ToolStripMenuItem();
            this.tabControl = new System.Windows.Forms.TabControl();
            this.tabEmailImport = new System.Windows.Forms.TabPage();
            this.tabMkgResults = new System.Windows.Forms.TabPage();
            this.tabFailedInjections = new System.Windows.Forms.TabPage();
            this.lblStatus = new System.Windows.Forms.Label();
            this.mkgMenuStrip.SuspendLayout();
            this.tabControl.SuspendLayout();
            this.SuspendLayout();
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(32, 19);
            // 
            // runToolStripMenuItem
            // 
            this.runToolStripMenuItem.Name = "runToolStripMenuItem";
            this.runToolStripMenuItem.Size = new System.Drawing.Size(32, 19);
            // 
            // mkgTestToolStripMenuItem
            // 
            this.mkgTestToolStripMenuItem.Name = "mkgTestToolStripMenuItem";
            this.mkgTestToolStripMenuItem.Size = new System.Drawing.Size(32, 19);
            // 
            // mkgMenuStrip
            // 
            this.mkgMenuStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.MkgMenuButton});
            this.mkgMenuStrip.Location = new System.Drawing.Point(0, 0);
            this.mkgMenuStrip.Name = "mkgMenuStrip";
            this.mkgMenuStrip.Size = new System.Drawing.Size(884, 24);
            this.mkgMenuStrip.TabIndex = 0;
            this.mkgMenuStrip.Text = "menuStrip1";
            // 
            // MkgMenuButton
            // 
            this.MkgMenuButton.Name = "MkgMenuButton";
            this.MkgMenuButton.Size = new System.Drawing.Size(44, 20);
            this.MkgMenuButton.Text = "MKG";
            // 
            // tabControl
            // 
            this.tabControl.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
            | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControl.Controls.Add(this.tabEmailImport);
            this.tabControl.Controls.Add(this.tabMkgResults);
            this.tabControl.Controls.Add(this.tabFailedInjections);
            this.tabControl.Location = new System.Drawing.Point(12, 35);
            this.tabControl.Name = "tabControl";
            this.tabControl.SelectedIndex = 0;
            this.tabControl.Size = new System.Drawing.Size(860, 500);
            this.tabControl.TabIndex = 0;
            // 
            // tabEmailImport
            // 
            this.tabEmailImport.Location = new System.Drawing.Point(4, 24);
            this.tabEmailImport.Name = "tabEmailImport";
            this.tabEmailImport.Padding = new System.Windows.Forms.Padding(3);
            this.tabEmailImport.Size = new System.Drawing.Size(852, 472);
            this.tabEmailImport.TabIndex = 0;
            this.tabEmailImport.Text = "Email Import";
            this.tabEmailImport.UseVisualStyleBackColor = true;
            // 
            // tabMkgResults
            // 
            this.tabMkgResults.Location = new System.Drawing.Point(4, 24);
            this.tabMkgResults.Name = "tabMkgResults";
            this.tabMkgResults.Padding = new System.Windows.Forms.Padding(3);
            this.tabMkgResults.Size = new System.Drawing.Size(852, 472);
            this.tabMkgResults.TabIndex = 1;
            this.tabMkgResults.Text = "MKG Results";
            this.tabMkgResults.UseVisualStyleBackColor = true;
            // 
            // tabFailedInjections
            // 
            this.tabFailedInjections.Location = new System.Drawing.Point(4, 24);
            this.tabFailedInjections.Name = "tabFailedInjections";
            this.tabFailedInjections.Padding = new System.Windows.Forms.Padding(3);
            this.tabFailedInjections.Size = new System.Drawing.Size(852, 472);
            this.tabFailedInjections.TabIndex = 2;
            this.tabFailedInjections.Text = "Failed Injections";
            this.tabFailedInjections.UseVisualStyleBackColor = true;
            // 
            // lblStatus
            // 
            this.lblStatus.AutoSize = true;
            this.lblStatus.Location = new System.Drawing.Point(12, 550);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(0, 15);
            this.lblStatus.TabIndex = 1;
            // 
            // Elcotec
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(884, 591);
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.tabControl);
            this.Controls.Add(this.mkgMenuStrip);
            this.MainMenuStrip = this.mkgMenuStrip;
            this.Name = "Elcotec";
            this.Text = "MKG Elcotec Automation";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.mkgMenuStrip.ResumeLayout(false);
            this.mkgMenuStrip.PerformLayout();
            this.tabControl.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem runToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem mkgTestToolStripMenuItem;
        private System.Windows.Forms.MenuStrip mkgMenuStrip;
        private System.Windows.Forms.ToolStripMenuItem MkgMenuButton;
        private System.Windows.Forms.TabControl tabControl;
        private System.Windows.Forms.TabPage tabEmailImport;
        private System.Windows.Forms.TabPage tabMkgResults;
        private System.Windows.Forms.TabPage tabFailedInjections;
        private System.Windows.Forms.Label lblStatus;
    }
}