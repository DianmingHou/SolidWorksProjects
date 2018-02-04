namespace ExportExcelAddin
{
    partial class FormStatusBar
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
            this.lblMessage = new System.Windows.Forms.Label();
            this.pgbRunningStatus = new System.Windows.Forms.ProgressBar();
            this.SuspendLayout();
            // 
            // lblMessage
            // 
            this.lblMessage.Location = new System.Drawing.Point(12, 9);
            this.lblMessage.Name = "lblMessage";
            this.lblMessage.Size = new System.Drawing.Size(398, 23);
            this.lblMessage.TabIndex = 0;
            this.lblMessage.Text = "text";
            this.lblMessage.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // pgbRunningStatus
            // 
            this.pgbRunningStatus.Location = new System.Drawing.Point(12, 35);
            this.pgbRunningStatus.Name = "pgbRunningStatus";
            this.pgbRunningStatus.Size = new System.Drawing.Size(398, 23);
            this.pgbRunningStatus.Style = System.Windows.Forms.ProgressBarStyle.Marquee;
            this.pgbRunningStatus.TabIndex = 1;
            // 
            // FormStatusBar
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(417, 94);
            this.Controls.Add(this.pgbRunningStatus);
            this.Controls.Add(this.lblMessage);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FormStatusBar";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "明细表导出";
            this.TopMost = true;
            this.Load += new System.EventHandler(this.FormStatusBar_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label lblMessage;
        private System.Windows.Forms.ProgressBar pgbRunningStatus;
    }
}