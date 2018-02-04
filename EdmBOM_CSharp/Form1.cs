using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Xml.Serialization;
using System.Windows.Forms;
using System.ComponentModel;
using EPDM.Interop.epdm;
using Microsoft.VisualBasic;

namespace EdmBOM_CSharp
{
    public partial class Form1 : Form
    {
        private IEdmVault5 vault1 = null;
        IEdmFile7 aFile;
        IEdmBom bom;
        IEdmBomMgr bomMgr;
        IEdmBomView bomView;

        public Form1()
        {
            InitializeComponent();
        }
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

        public void SelectFiles_Click(System.Object sender, System.EventArgs e)
        {
            try
            {
                File1List.Items.Clear();

                IEdmVault7 vault2 = null;
                if (vault1 == null)
                {
                    vault1 = new EdmVault5();
                }
                vault2 = (IEdmVault7)vault1;

                if (!vault1.IsLoggedIn)
                {
                    vault1.LoginAuto(VaultsComboBox.Text, this.Handle.ToInt32());
                }

                //Set the initial directory in the Select File dialog
                OpenFileDialog1.InitialDirectory = vault1.RootFolderPath;

                //Show the Select File dialog
                System.Windows.Forms.DialogResult DialogResult;
                DialogResult = OpenFileDialog1.ShowDialog();

                if (!(DialogResult == System.Windows.Forms.DialogResult.OK))
                {
                    // do nothing
                }
                else
                {
                    IEdmFolder5 ppoRetParentFolder;
                    foreach (string FileName in OpenFileDialog1.FileNames)
                    {
                        File1List.Items.Add(FileName);
                        aFile = (IEdmFile7)vault1.GetFileFromPath(FileName, out ppoRetParentFolder);
                    }
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

        public void GetBOM_Click(System.Object sender, System.EventArgs e)
        {

            try
            {
                IEdmVault7 vault2 = null;
                if (vault1 == null)
                {
                    vault1 = new EdmVault5();
                }
                vault2 = (IEdmVault9)vault1;
                if (!vault1.IsLoggedIn)
                {
                    vault1.LoginAuto(VaultsComboBox.Text, this.Handle.ToInt32());
                }

                if (aFile != null)
                {
                    // Get named BOMs and their versions for the selected file
                    EdmBomInfo[] derivedBOMs = null;
                    aFile.GetDerivedBOMs(out derivedBOMs);

                    int arrSize = 0;
                    EdmBomVersion[] ppoVersions = null;
                    int i = 0;
                    int j = 0;
                    int id = 0;
                    string str = "";
                    string verstr = "";
                    int verArrSize = 0;
                    arrSize = derivedBOMs.Length;

                    while (i < arrSize)
                    {
                        id = derivedBOMs[i].mlBomID;
                        bom = (IEdmBom)vault2.GetObject(EdmObjectType.EdmObject_BOM, id);
                        str = "Named BOM: " + derivedBOMs[i].mbsBomName + Constants.vbLf + "Check-out user: " + bom.CheckOutUserID + Constants.vbLf + "Current state: " + bom.CurrentState.Name + Constants.vbLf + "Current version: " + bom.CurrentVersion + Constants.vbLf + "ID: " + bom.FileID + Constants.vbLf + "Is checked out: " + bom.IsCheckedOut;
                        MessageBox.Show(str);
                        bom.GetVersions(out ppoVersions);
                        verArrSize = ppoVersions.Length;
                        while (j < verArrSize)
                        {
                            verstr = "BOM version: " + Constants.vbLf + "Type as defined in EdmBomVersionType: " + ppoVersions[j].meType + Constants.vbLf + "Version number: " + ppoVersions[j].mlVersion + Constants.vbLf + "Date: " + ppoVersions[j].moDate + Constants.vbLf + "Label: " + ppoVersions[j].mbsTag + Constants.vbLf + "Comment: " + ppoVersions[j].mbsComment;
                            MessageBox.Show(verstr);
                            j = j + 1;
                        }
                        i = i + 1;
                    }

                    // Get a BOM view with the specified layout
                    bomMgr = (IEdmBomMgr)vault2.CreateUtility(EdmUtility.EdmUtil_BomMgr);
                    EdmBomLayout[] ppoRetLayouts = null;
                    EdmBomLayout ppoRetLayout = default(EdmBomLayout);
                    bomMgr.GetBomLayouts(out ppoRetLayouts);
                    i = 0;
                    arrSize = ppoRetLayouts.Length;
                    str = "";
                    while (i < arrSize)
                    {
                        ppoRetLayout = ppoRetLayouts[i];
                        str = "BOM Layout " + i + ": " + ppoRetLayout.mbsLayoutName + Constants.vbLf + "ID: " + ppoRetLayout.mlLayoutID;
                        if (ppoRetLayout.mbsLayoutName == "BOM")
                        {
                            bomView = aFile.GetComputedBOM(ppoRetLayout.mbsLayoutName, 0, "@", (int)EdmBomFlag.EdmBf_AsBuilt + (int)EdmBomFlag.EdmBf_ShowSelected);
                        }
                        MessageBox.Show(str);
                        i = i + 1;
                    }

                    // Display BOM view row and column information
                    object[] ppoRows = null;
                    IEdmBomCell ppoRow = default(IEdmBomCell);
                    bomView.GetRows(out ppoRows);
                    i = 0;
                    arrSize = ppoRows.Length;
                    str = "";
                    EdmBomColumn[] ppoColumns = null;
                    bomView.GetColumns(out ppoColumns);
                    while (i < arrSize)
                    {
                        ppoRow = (IEdmBomCell)ppoRows[i];
                        str = "BOM Row " + i + ": " + Constants.vbLf + "Item ID: " + ppoRow.GetItemID() + Constants.vbLf + "Path name: " + ppoRow.GetPathName() + Constants.vbLf + "Tree level: " + ppoRow.GetTreeLevel() + Constants.vbLf + " Is locked? " + ppoRow.IsLocked();
                        MessageBox.Show(str);

                        int js = 0;
                        int ararSize = ppoColumns.Length;
                        while (js < ararSize)
                        {
                            object poValue = null;
                            object poComputedValue = null;
                            string pbsConfiguration = "";
                            bool pbReadOnly = false;
                            int plFocusNode = 0;
                            ppoRow.GetVar(ppoColumns[js].mlVariableID, ppoColumns[js].meType, out poValue, out poComputedValue, out pbsConfiguration, out pbReadOnly);
                            str = "BOM Column " + js + ": " + Constants.vbLf + "Header: " + ppoColumns[js].mbsCaption +
                                Constants.vbLf + "Value " + poValue.ToString() +
                                Constants.vbLf + "ComputeValue " + poComputedValue.ToString() +
                                Constants.vbLf + "Config: " + pbsConfiguration +
                                Constants.vbLf + "ReadOnly: " + pbReadOnly.ToString();
                            MessageBox.Show(str);
                            js = js + 1;
                        }
                        i = i + 1;
                    }

                   
                   
                    
                    //bomView.InsertRow((EdmBomCell)ppoRows[0], EdmBomInsertRowOption.EdmBomInsertRowOption_BelowRow, out ppoNewRow);
                   // ppoNewRow.SetVar(ppoColumns[0].mlVariableID, ppoColumns[0].meType, "vault_name\\temp", "", EdmBomSetVarOption.EdmBomSetVarOption_Both, out errMsg);
                    


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