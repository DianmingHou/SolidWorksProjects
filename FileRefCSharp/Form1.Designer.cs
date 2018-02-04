namespace FileRefCSharp
{
    partial class Form1
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
            this.VaultsLabel = new System.Windows.Forms.Label();
            this.VaultsComboBox = new System.Windows.Forms.ComboBox();
            this.BatchRefLabel = new System.Windows.Forms.Label();
            this.BrowseButton = new System.Windows.Forms.Button();
            this.BatchRefListBox = new System.Windows.Forms.ListBox();
            this.WriteXMLButton = new System.Windows.Forms.Button();
            this.BatchRefOpenFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.SuspendLayout();
            // 
            // VaultsLabel
            // 
            this.VaultsLabel.AutoSize = true;
            this.VaultsLabel.Location = new System.Drawing.Point(13, 26);
            this.VaultsLabel.Name = "VaultsLabel";
            this.VaultsLabel.Size = new System.Drawing.Size(91, 13);
            this.VaultsLabel.TabIndex = 0;
            this.VaultsLabel.Text = "Select vault view:";
            // 
            // VaultsComboBox
            // 
            this.VaultsComboBox.FormattingEnabled = true;
            this.VaultsComboBox.Location = new System.Drawing.Point(16, 59);
            this.VaultsComboBox.Name = "VaultsComboBox";
            this.VaultsComboBox.Size = new System.Drawing.Size(121, 21);
            this.VaultsComboBox.TabIndex = 1;
            // 
            // BatchRefLabel
            // 
            this.BatchRefLabel.AutoSize = true;
            this.BatchRefLabel.Location = new System.Drawing.Point(16, 116);
            this.BatchRefLabel.Name = "BatchRefLabel";
            this.BatchRefLabel.Size = new System.Drawing.Size(160, 13);
            this.BatchRefLabel.TabIndex = 2;
            this.BatchRefLabel.Text = "Files for which to get references:";
            // 
            // BrowseButton
            // 
            this.BrowseButton.Location = new System.Drawing.Point(197, 106);
            this.BrowseButton.Name = "BrowseButton";
            this.BrowseButton.Size = new System.Drawing.Size(75, 23);
            this.BrowseButton.TabIndex = 3;
            this.BrowseButton.Text = "Browse...";
            this.BrowseButton.UseVisualStyleBackColor = true;
            this.BrowseButton.Click += new System.EventHandler(BrowseButton_Click);
            // 
            // BatchRefListBox
            // 
            this.BatchRefListBox.FormattingEnabled = true;
            this.BatchRefListBox.Location = new System.Drawing.Point(13, 150);
            this.BatchRefListBox.Name = "BatchRefListBox";
            this.BatchRefListBox.Size = new System.Drawing.Size(259, 95);
            this.BatchRefListBox.TabIndex = 4;
            // 
            // WriteXMLButton
            // 
            this.WriteXMLButton.Location = new System.Drawing.Point(13, 270);
            this.WriteXMLButton.Name = "WriteXMLButton";
            this.WriteXMLButton.Size = new System.Drawing.Size(259, 23);
            this.WriteXMLButton.TabIndex = 5;
            this.WriteXMLButton.Text = "Write file references to an XML file";
            this.WriteXMLButton.UseVisualStyleBackColor = true;
            this.WriteXMLButton.Click += new System.EventHandler(this.WriteXmlButton_Click);

            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 308);
            this.Controls.Add(this.WriteXMLButton);
            this.Controls.Add(this.BatchRefListBox);
            this.Controls.Add(this.BrowseButton);
            this.Controls.Add(this.BatchRefLabel);
            this.Controls.Add(this.VaultsComboBox);
            this.Controls.Add(this.VaultsLabel);
            this.Name = "Form1";
            this.Text = "Get file references";
            this.Load += new System.EventHandler(this.FileReferencesCSharp_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label VaultsLabel;
        private System.Windows.Forms.ComboBox VaultsComboBox;
        private System.Windows.Forms.Label BatchRefLabel;
        private System.Windows.Forms.Button BrowseButton;
        private System.Windows.Forms.ListBox BatchRefListBox;
        private System.Windows.Forms.Button WriteXMLButton;
        private System.Windows.Forms.OpenFileDialog BatchRefOpenFileDialog;
    }
}

