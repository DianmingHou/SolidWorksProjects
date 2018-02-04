namespace ExportExcelAddin
{
    partial class ExportExcelForm
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
            this.labelSelectFile = new System.Windows.Forms.Label();
            this.textBoxSelectFile = new System.Windows.Forms.TextBox();
            this.labelSaveFile = new System.Windows.Forms.Label();
            this.textBoxSaveFilePath = new System.Windows.Forms.TextBox();
            this.buttonFilePath = new System.Windows.Forms.Button();
            this.buttonOK = new System.Windows.Forms.Button();
            this.buttonClose = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // labelSelectFile
            // 
            this.labelSelectFile.AutoSize = true;
            this.labelSelectFile.Location = new System.Drawing.Point(13, 33);
            this.labelSelectFile.Name = "labelSelectFile";
            this.labelSelectFile.Size = new System.Drawing.Size(77, 12);
            this.labelSelectFile.TabIndex = 0;
            this.labelSelectFile.Text = "选中文件名：";
            // 
            // textBoxSelectFile
            // 
            this.textBoxSelectFile.Location = new System.Drawing.Point(96, 33);
            this.textBoxSelectFile.Name = "textBoxSelectFile";
            this.textBoxSelectFile.ReadOnly = true;
            this.textBoxSelectFile.Size = new System.Drawing.Size(252, 21);
            this.textBoxSelectFile.TabIndex = 1;
            // 
            // labelSaveFile
            // 
            this.labelSaveFile.AutoSize = true;
            this.labelSaveFile.Location = new System.Drawing.Point(15, 74);
            this.labelSaveFile.Name = "labelSaveFile";
            this.labelSaveFile.Size = new System.Drawing.Size(77, 12);
            this.labelSaveFile.TabIndex = 2;
            this.labelSaveFile.Text = "保存报表名：";
            // 
            // textBoxSaveFilePath
            // 
            this.textBoxSaveFilePath.Location = new System.Drawing.Point(96, 74);
            this.textBoxSaveFilePath.Name = "textBoxSaveFilePath";
            this.textBoxSaveFilePath.Size = new System.Drawing.Size(221, 21);
            this.textBoxSaveFilePath.TabIndex = 3;
            // 
            // buttonFilePath
            // 
            this.buttonFilePath.Location = new System.Drawing.Point(323, 74);
            this.buttonFilePath.Name = "buttonFilePath";
            this.buttonFilePath.Size = new System.Drawing.Size(25, 23);
            this.buttonFilePath.TabIndex = 4;
            this.buttonFilePath.Text = "...";
            this.buttonFilePath.UseVisualStyleBackColor = true;
            this.buttonFilePath.Click += new System.EventHandler(this.buttonFilePath_Click);
            // 
            // buttonOK
            // 
            this.buttonOK.Location = new System.Drawing.Point(45, 121);
            this.buttonOK.Name = "buttonOK";
            this.buttonOK.Size = new System.Drawing.Size(102, 23);
            this.buttonOK.TabIndex = 5;
            this.buttonOK.Text = "确定";
            this.buttonOK.UseVisualStyleBackColor = true;
            this.buttonOK.Click += new System.EventHandler(this.buttonOK_Click);
            // 
            // buttonClose
            // 
            this.buttonClose.Location = new System.Drawing.Point(217, 121);
            this.buttonClose.Name = "buttonClose";
            this.buttonClose.Size = new System.Drawing.Size(115, 23);
            this.buttonClose.TabIndex = 6;
            this.buttonClose.Text = "取消";
            this.buttonClose.UseVisualStyleBackColor = true;
            this.buttonClose.Click += new System.EventHandler(this.buttonClose_Click);
            // 
            // ExportExcelForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(361, 163);
            this.Controls.Add(this.buttonClose);
            this.Controls.Add(this.buttonOK);
            this.Controls.Add(this.buttonFilePath);
            this.Controls.Add(this.textBoxSaveFilePath);
            this.Controls.Add(this.labelSaveFile);
            this.Controls.Add(this.textBoxSelectFile);
            this.Controls.Add(this.labelSelectFile);
            this.Name = "ExportExcelForm";
            this.Text = "ExportExcelForm";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label labelSelectFile;
        private System.Windows.Forms.TextBox textBoxSelectFile;
        private System.Windows.Forms.Label labelSaveFile;
        private System.Windows.Forms.TextBox textBoxSaveFilePath;
        private System.Windows.Forms.Button buttonFilePath;
        private System.Windows.Forms.Button buttonOK;
        private System.Windows.Forms.Button buttonClose;

    }
}