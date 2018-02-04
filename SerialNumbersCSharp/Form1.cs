using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using EPDM.Interop.epdm;

namespace SerialNumbersCSharp
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private IEdmVault5 vault1 = null;
        string aSerialNbrName;
        string aFileName;
        IEdmFile5 aFile;
        string aFolder;
        IEdmSerNoGen7 serialNbrs;

        private void Form1_Load(System.Object sender, System.EventArgs e)
        {

            try
            {
                vault1 = new EdmVault5();
                IEdmVault8 vault = (IEdmVault8)vault1;
                EdmViewInfo[] Views = null;

                vault.GetVaultViews(out Views, false);
                VaultsComboBox.Items.Clear();
                foreach (EdmViewInfo View in Views)
                {
                    VaultsComboBox.Items.Add(View.mbsVaultName);
                }
                if (VaultsComboBox.Items.Count > 0)
                {
                    VaultsComboBox.Text = (string)VaultsComboBox.Items[0];
                }

            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                MessageBox.Show("HRESULT = 0x" + ex.ErrorCode.ToString("X") + " " + ex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void Button1_Click(System.Object sender, System.EventArgs e)
        {
            try
            {
                //Only create a new vault object
                //if one hasn't been created yet
                IEdmVault11 vault2 = null;
                if (vault1 == null)
                {
                    vault1 = new EdmVault5();
                }
                vault2 = (IEdmVault11)vault1;
                if (!vault1.IsLoggedIn)
                {
                    //Log into selected vault as the current user
                    vault1.LoginAuto(VaultsComboBox.Text, this.Handle.ToInt32());
                }

                serialNbrs = (IEdmSerNoGen7)vault2.CreateUtility(EdmUtility.EdmUtil_SerNoGen);
                string[] names = { };
                serialNbrs.GetSerialNumberNames(out names);
                string myMessage = "";
                myMessage = "Serial number names in selected vault: " + "\n ";
                int idx = 0;
                idx = (names.GetLowerBound(0));
                while (idx <= (names.GetUpperBound(0)))
                {
                    myMessage = myMessage + names[idx] + "\n";
                    idx = idx + 1;
                }

                // Use this serial number generator
                aSerialNbrName = names[0];

                MessageBox.Show(myMessage);

            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                MessageBox.Show("HRESULT = 0x" + ex.ErrorCode.ToString("X") + " " + ex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void BrowseButton_Click(System.Object sender, System.EventArgs e)
        {
            try
            {

                //Set the initial directory in the Select a file dialog
                OpenFileDialog1.InitialDirectory = vault1.RootFolderPath;
                //Show the Select a file dialog
                System.Windows.Forms.DialogResult DialogResult;
                DialogResult = OpenFileDialog1.ShowDialog();

                if (!(DialogResult == System.Windows.Forms.DialogResult.OK))
                {
                    // Do nothing
                }
                else
                {
                    // Browse for a file whose serial number to set
                    // File's data card must have a Part Number associated
                    // with a serial number generator and must be checked out
                    string fileName = OpenFileDialog1.FileName;
                    FileListBox.Items.Add(fileName);
                    IEdmFolder5 retFolder = default(IEdmFolder5);
                    aFile = vault1.GetFileFromPath(fileName, out retFolder);

                }

            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                MessageBox.Show("HRESULT = 0x" + ex.ErrorCode.ToString("X") + " " + ex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void NewButton_Click(System.Object sender, System.EventArgs e)
        {
            try
            {

                IEdmSerNoValue serialNbrValue = default(IEdmSerNoValue);
                serialNbrValue = serialNbrs.AllocSerNoValue(aSerialNbrName, this.Handle.ToInt32(), " ", 0, 0, 0, 0);
                dynamic serialNbrValueValue = serialNbrValue.Value;
                IEdmEnumeratorVariable5 enumVariable = default(IEdmEnumeratorVariable5);
                enumVariable = aFile.GetEnumeratorVariable(aFileName);
                // Set the Part Number of the selected file
                enumVariable.SetVar("Part Number", "@", serialNbrValueValue);
                IEdmEnumeratorVariable8 enumVariable8 = (IEdmEnumeratorVariable8)enumVariable;
                enumVariable8.CloseFile(true);

                MessageBox.Show("Part Number set to " + serialNbrValueValue.ToString() + "." + " ");


            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                MessageBox.Show("HRESULT = 0x" + ex.ErrorCode.ToString("X") + " " + ex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }

}
