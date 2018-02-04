using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using EdmLib;
//using EPDM.Interop.epdm;

namespace ExportExcelAddin
{
    public partial class ExportExcelForm : Form
    {
        public ExportExcelForm()
        {
            InitializeComponent();
        }

        private void buttonFilePath_Click(object sender, EventArgs e)
        {
            //SaveFileDialog
            SaveFileDialog openFile = new SaveFileDialog();
            openFile.Filter = "Worksheet Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*||";
            openFile.RestoreDirectory = true;
            openFile.AddExtension = true;
            openFile.FilterIndex = 1;
            if (openFile.ShowDialog() == DialogResult.OK)
            {
                string fName = openFile.FileName;
                textBoxSaveFilePath.Text = fName;
            }
        }

        private void buttonOK_Click(object sender, EventArgs e)
        {
            //处理窗口
            this.DialogResult = DialogResult.OK;  
            this.Close(); 
        }

        private void buttonClose_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close(); 
        }


        private IEdmVault5 m_poVault;
        public IEdmVault5 vault
        {
            get { return m_poVault; }
            set { m_poVault = value; }
        }
        public string selectFile
        {
            get { return textBoxSelectFile.Text; }
            set { textBoxSelectFile.Text = value; }
        }
        public string saveFilePath
        {
            get { return textBoxSaveFilePath.Text; }
            set { textBoxSaveFilePath.Text = value; }
        }
    }
}
