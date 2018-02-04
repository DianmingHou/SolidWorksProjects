using Microsoft.VisualBasic;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;

namespace SerialNumbersCSharp
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
            this.Button1 = new System.Windows.Forms.Button();
            this.FileListBox = new System.Windows.Forms.ListBox();
            this.OpenFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.BrowseButton = new System.Windows.Forms.Button();
            this.NewButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            //
            //VaultsLabel
            //
            this.VaultsLabel.AutoSize = true;
            this.VaultsLabel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.VaultsLabel.Location = new System.Drawing.Point(25, 28);
            this.VaultsLabel.Name = "VaultsLabel";
            this.VaultsLabel.Size = new System.Drawing.Size(91, 13);
            this.VaultsLabel.TabIndex = 0;
            this.VaultsLabel.Text = "Select vault view:";
            //
            //VaultsComboBox
            //
            this.VaultsComboBox.FormattingEnabled = true;
            this.VaultsComboBox.Location = new System.Drawing.Point(28, 44);
            this.VaultsComboBox.Name = "VaultsComboBox";
            this.VaultsComboBox.Size = new System.Drawing.Size(190, 21);
            this.VaultsComboBox.TabIndex = 1;
            //
            //Button1
            //
            this.Button1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.Button1.Location = new System.Drawing.Point(28, 84);
            this.Button1.Name = "Button1";
            this.Button1.Size = new System.Drawing.Size(190, 23);
            this.Button1.TabIndex = 2;
            this.Button1.Text = "Get Vault Serial Number Names";
            this.Button1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.Button1.UseVisualStyleBackColor = true;
            this.Button1.Click += new System.EventHandler(this.Button1_Click);
            //
            //FileListBox
            //
            this.FileListBox.FormattingEnabled = true;
            this.FileListBox.Location = new System.Drawing.Point(28, 141);
            this.FileListBox.Name = "FileListBox";
            this.FileListBox.Size = new System.Drawing.Size(190, 17);
            this.FileListBox.TabIndex = 3;
            //
            //OpenFileDialog1
            //
            this.OpenFileDialog1.FileName = "OpenFileDialog1";
            this.OpenFileDialog1.Multiselect = true;
            this.OpenFileDialog1.Title = "Select a file";
            //
            //BrowseButton
            //
            this.BrowseButton.Location = new System.Drawing.Point(239, 135);
            this.BrowseButton.Name = "BrowseButton";
            this.BrowseButton.Size = new System.Drawing.Size(56, 23);
            this.BrowseButton.TabIndex = 4;
            this.BrowseButton.Text = "Browse...";
            this.BrowseButton.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.BrowseButton.UseVisualStyleBackColor = true;
            this.BrowseButton.Click += new System.EventHandler(this.BrowseButton_Click);
            //
            //NewButton
            //
            this.NewButton.Location = new System.Drawing.Point(28, 182);
            this.NewButton.Name = "NewButton";
            this.NewButton.Size = new System.Drawing.Size(190, 23);
            this.NewButton.TabIndex = 5;
            this.NewButton.Text = "Set New Serial Number";
            this.NewButton.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.NewButton.UseVisualStyleBackColor = true;
            this.NewButton.Click += new System.EventHandler(this.NewButton_Click);
            //
            //Form1
            //
            this.AutoScaleDimensions = new System.Drawing.SizeF(6f, 13f);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(326, 262);
            this.Controls.Add(this.NewButton);
            this.Controls.Add(this.BrowseButton);
            this.Controls.Add(this.FileListBox);
            this.Controls.Add(this.Button1);
            this.Controls.Add(this.VaultsComboBox);
            this.Controls.Add(this.VaultsLabel);
            this.Name = "Form1";
            this.Text = "Serial Numbers";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        internal System.Windows.Forms.Label VaultsLabel;
        internal System.Windows.Forms.ComboBox VaultsComboBox;
        internal System.Windows.Forms.Button Button1;
        internal System.Windows.Forms.ListBox FileListBox;
        internal System.Windows.Forms.Button BrowseButton;
        internal System.Windows.Forms.Button NewButton;
        internal System.Windows.Forms.OpenFileDialog OpenFileDialog1;

    }
}

