namespace CardVars_CSharp
{
    partial class Form1
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.VaultsLabel = new System.Windows.Forms.Label();
            this.VaultsComboBox = new System.Windows.Forms.ComboBox();
            this.BrowseButton = new System.Windows.Forms.Button();
            this.ListBox = new System.Windows.Forms.ListBox();
            this.GetVars = new System.Windows.Forms.Button();
            this.OpenFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.SuspendLayout();
            // 
            // VaultsLabel
            // 
            this.VaultsLabel.AutoSize = true;
            this.VaultsLabel.Location = new System.Drawing.Point(13, 26);
            this.VaultsLabel.Name = "VaultsLabel";
            this.VaultsLabel.Size = new System.Drawing.Size(94, 13);
            this.VaultsLabel.TabIndex = 0;
            this.VaultsLabel.Text = " Select vault view:";
            // 
            // VaultsComboBox
            // 
            this.VaultsComboBox.FormattingEnabled = true;
            this.VaultsComboBox.Location = new System.Drawing.Point(16, 42);
            this.VaultsComboBox.Name = "VaultsComboBox";
            this.VaultsComboBox.Size = new System.Drawing.Size(121, 21);
            this.VaultsComboBox.TabIndex = 1;
            // 
            // BrowseButton
            // 
            this.BrowseButton.Location = new System.Drawing.Point(16, 85);
            this.BrowseButton.Name = "BrowseButton";
            this.BrowseButton.Size = new System.Drawing.Size(98, 23);
            this.BrowseButton.TabIndex = 3;
            this.BrowseButton.Text = "Browse...";
            this.BrowseButton.UseVisualStyleBackColor = true;
            this.BrowseButton.Click += new System.EventHandler(this.BrowseButton_Click);
            // 
            // ListBox
            // 
            this.ListBox.FormattingEnabled = true;
            this.ListBox.HorizontalScrollbar = true;
            this.ListBox.Location = new System.Drawing.Point(16, 114);
            this.ListBox.Name = "ListBox";
            this.ListBox.Size = new System.Drawing.Size(259, 43);
            this.ListBox.TabIndex = 4;
            // 
            // GetVars
            // 
            this.GetVars.Location = new System.Drawing.Point(84, 178);
            this.GetVars.Name = "GetVars";
            this.GetVars.Size = new System.Drawing.Size(98, 23);
            this.GetVars.TabIndex = 5;
            this.GetVars.Text = "Get variables";
            this.GetVars.UseVisualStyleBackColor = true;
            this.GetVars.Click += new System.EventHandler(GetVars_Click);
            // 
            // OpenFileDialog
            // 
            this.OpenFileDialog.Title = "Open";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(317, 215);
            this.Controls.Add(this.GetVars);
            this.Controls.Add(this.ListBox);
            this.Controls.Add(this.BrowseButton);
            this.Controls.Add(this.VaultsComboBox);
            this.Controls.Add(this.VaultsLabel);
            this.Name = "Form1";
            this.Text = "Access file variables";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }


        #endregion
        internal System.Windows.Forms.Label VaultsLabel;
        internal System.Windows.Forms.ComboBox VaultsComboBox;
        internal System.Windows.Forms.Button BrowseButton;
        internal System.Windows.Forms.ListBox ListBox;
        internal System.Windows.Forms.Button GetVars;
        internal System.Windows.Forms.OpenFileDialog OpenFileDialog;
    }
}

