using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using EPDM.Interop.epdm;
using Microsoft.VisualBasic;

namespace CardVars_CSharp
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private IEdmVault5 vault1 = null;

        public void Form1_Load(System.Object sender, System.EventArgs e)
        {
            try
            {
                IEdmVault5 vault1 = new EdmVault5();
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

        public void BrowseButton_Click(System.Object sender, System.EventArgs e)
        {
            try
            {
                ListBox.Items.Clear();

                if (vault1 == null)
                {
                    vault1 = new EdmVault5();
                }
                if (!vault1.IsLoggedIn)
                {
                    //Log into selected vault as the current user
                    vault1.LoginAuto(VaultsComboBox.Text, this.Handle.ToInt32());
                }

                //Set the initial directory in the Open dialog
                OpenFileDialog.InitialDirectory = vault1.RootFolderPath;
                //Show the Open dialog
                System.Windows.Forms.DialogResult DialogResult;
                DialogResult = OpenFileDialog.ShowDialog();
                //If the user didn't click Open, exit
                if (!(DialogResult == System.Windows.Forms.DialogResult.OK))
                {
                    return;
                }

                foreach (string FileName in OpenFileDialog.FileNames)
                {
                    ListBox.Items.Add(FileName);
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


        public void GetVars_Click(System.Object sender, System.EventArgs e)
        {

            try
            {
                IEdmVault7 vault2 = null;
                if (vault1 == null)
                {
                    vault1 = new EdmVault5();
                }
                vault2 = (IEdmVault7)vault1;
                if (!vault1.IsLoggedIn)
                {
                    //Log into selected vault as the current user
                    vault1.LoginAuto(VaultsComboBox.Text, this.Handle.ToInt32());
                }

                IEdmFile5 aFile = default(IEdmFile5);
                IEdmFolder5 ppoRetParentFolder = null;
                aFile = vault2.GetFileFromPath(ListBox.Items[0].ToString(), out ppoRetParentFolder);

                //Get card variables only from a file checked into the vault
                if (!aFile.IsLocked)
                {
                    aFile.LockFile(vault2.RootFolderID, this.Handle.ToInt32(), (int)EdmLockFlag.EdmLock_Simple);

                    IEdmEnumeratorVariable5 varEnum = default(IEdmEnumeratorVariable5);
                    varEnum = aFile.GetEnumeratorVariable();

                    object[] valueList = null;
                    varEnum.GetUpdateVars(aFile.LockedInFolderID, out valueList);

                    int idx = 0;
                    idx = Information.LBound(valueList);
                    int upper = 0;
                    upper = Information.UBound(valueList);

                    string msg = null;
                    msg = "Card variables for " + aFile.Name + " in configuration, @:" + Constants.vbLf + Constants.vbLf;

                    IEdmVariableMgr5 varMgr = default(IEdmVariableMgr5);
                    varMgr = (IEdmVariableMgr5)aFile.Vault;

                    IEdmVariable5 var = default(IEdmVariable5);
                    IEdmVariableValue6 value = default(IEdmVariableValue6);
                    while (idx <= upper)
                    {
                        value = (IEdmVariableValue6)valueList[idx];
                        idx = idx + 1;
                        var = varMgr.GetVariable(value.VariableID);
                        msg = msg + value.VariableName + " = > " + value.GetValue("@").ToString() + Constants.vbLf;
                        msg = msg + "EdmVariableFlags: " + var.Flags + ", EdmVariableType: " + var.VariableType + Constants.vbLf + Constants.vbLf;
                        //msg = msg & "EdmVariableFlags: " & value.VariableFlags & ", EdmVariableType: " & value.VariableType & vbLf & vbLf
                    }

                    MessageBox.Show(msg);

                    aFile.UndoLockFile(this.Handle.ToInt32());
                }
                else
                {
                    //User selected a checked-out file
                    MessageBox.Show("Please select a checked-in file.");
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
    }
}


