namespace EdmBOM_CSharp
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
            this.SelectFiles = new System.Windows.Forms.Button();
            this.File1List = new System.Windows.Forms.ListBox();
            this.GetBOM = new System.Windows.Forms.Button();
            this.OpenFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.SuspendLayout();
            //
            //VaultsLabel
            //
            this.VaultsLabel.AutoSize = true;
            this.VaultsLabel.Location = new System.Drawing.Point(36, 24);
            this.VaultsLabel.Name = "VaultsLabel";
            this.VaultsLabel.Size = new System.Drawing.Size(91, 13);
            this.VaultsLabel.TabIndex = 0;
            this.VaultsLabel.Text = "Select vault view:";
            //
            //VaultsComboBox
            //
            this.VaultsComboBox.FormattingEnabled = true;
            this.VaultsComboBox.Location = new System.Drawing.Point(39, 40);
            this.VaultsComboBox.Name = "VaultsComboBox";
            this.VaultsComboBox.Size = new System.Drawing.Size(121, 21);
            this.VaultsComboBox.TabIndex = 1;
            //
            //SelectFiles
            //
            this.SelectFiles.Location = new System.Drawing.Point(39, 85);
            this.SelectFiles.Name = "SelectFiles";
            this.SelectFiles.Size = new System.Drawing.Size(191, 23);
            this.SelectFiles.TabIndex = 2;
            this.SelectFiles.Text = "Select file...";
            this.SelectFiles.UseVisualStyleBackColor = true;
            this.SelectFiles.Click += new System.EventHandler(SelectFiles_Click);
            //
            //File1List
            //
            this.File1List.FormattingEnabled = true;
            this.File1List.HorizontalScrollbar = true;
            this.File1List.Location = new System.Drawing.Point(40, 114);
            this.File1List.Name = "File1List";
            this.File1List.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.File1List.Size = new System.Drawing.Size(220, 43);
            this.File1List.TabIndex = 4;
            //
            //GetBOM
            //
            this.GetBOM.Location = new System.Drawing.Point(40, 183);
            this.GetBOM.Name = "GetBOM";
            this.GetBOM.Size = new System.Drawing.Size(157, 23);
            this.GetBOM.TabIndex = 6;
            this.GetBOM.Text = "Get BOM";
            this.GetBOM.UseVisualStyleBackColor = true;
            this.GetBOM.Click += new System.EventHandler(this.GetBOM_Click);
            //
            //OpenFileDialog1
            //
            this.OpenFileDialog1.FileName = "OpenFileDialog1";
            this.OpenFileDialog1.Title = "Select File";
            //
            //Form1
            //
            this.AutoScaleDimensions = new System.Drawing.SizeF(6f, 13f);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 235);
            this.Controls.Add(this.GetBOM);
            this.Controls.Add(this.File1List);
            this.Controls.Add(this.SelectFiles);
            this.Controls.Add(this.VaultsComboBox);
            this.Controls.Add(this.VaultsLabel);
            this.Name = "Form1";
            this.Text = "Bill of Materials";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        internal System.Windows.Forms.Label VaultsLabel;
        internal System.Windows.Forms.ComboBox VaultsComboBox;
        internal System.Windows.Forms.Button SelectFiles;
        internal System.Windows.Forms.ListBox File1List;
        internal System.Windows.Forms.Button GetBOM;
        internal System.Windows.Forms.OpenFileDialog OpenFileDialog1;

        #endregion
    }
}

