using System.IO;
using System.Xml.Serialization;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Windows.Forms;
using System.ComponentModel;
using EPDM.Interop.epdm;


namespace FileRefCSharp
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        IEdmVault5 vault1 = null;

        public void FileReferencesCSharp_Load(System.Object sender, System.EventArgs e)
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
                BatchRefListBox.Items.Clear();

                //Only create a new vault object
                //if one hasn't been created yet
                if (vault1 == null)
                    vault1 = new EdmVault5();
                if (!vault1.IsLoggedIn)
                {
                    //Log into selected vault as the current user
                    vault1.LoginAuto(VaultsComboBox.Text, this.Handle.ToInt32());
                }

                //Set the initial directory in the File Open dialog
                BatchRefOpenFileDialog.InitialDirectory = vault1.RootFolderPath;
                //Show the File Open dialog
                System.Windows.Forms.DialogResult DialogResult = default(System.Windows.Forms.DialogResult);
                DialogResult = BatchRefOpenFileDialog.ShowDialog();
                //If the user didn't click Open, exit the sub
                if (!(DialogResult == System.Windows.Forms.DialogResult.OK))
                    return;

                foreach (string FileName in BatchRefOpenFileDialog.FileNames)
                {
                    BatchRefListBox.Items.Add(FileName);
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


        public void GetReferencedFiles(IEdmReference10 Reference, string FilePath, int Level, string ProjectName, ref Dictionary<string, string> RefFilesDictionary)
        {
            try
            {
                bool Top = false;
                if (Reference == null)
                {
                    //This is the first time this function is called for this 
                    //reference tree; i.e., this is the root
                    Top = true;
                    //Add the top-level file path to the dictionary
                    RefFilesDictionary.Add(FilePath, Level.ToString());
                    IEdmFile5 File = null;
                    IEdmFolder5 ParentFolder = null;
                    File = vault1.GetFileFromPath(FilePath, out ParentFolder);
                    //Get the reference tree for this file
                    Reference = (IEdmReference10)File.GetReferenceTree(ParentFolder.ID);
                    GetReferencedFiles(Reference, "", Level + 1, ProjectName, ref RefFilesDictionary);
                }
                else
                {
                    //Execute this code when this function is called recursively; 
                    //i.e., this is not the top-level IEdmReference in the tree

                    //Recursively traverse the references
                    IEdmPos5 pos = default(IEdmPos5);
                    IEdmReference10 Reference2 = (IEdmReference10)Reference;
                    pos = Reference2.GetFirstChildPosition3(ProjectName, Top, true, (int)EdmRefFlags.EdmRef_File, "", 0);
                    IEdmReference10 @ref = default(IEdmReference10);
                    while ((!pos.IsNull))
                    {
                        @ref = (IEdmReference10)Reference.GetNextChild(pos);
                        RefFilesDictionary.Add(@ref.FoundPath, Level.ToString());
                        GetReferencedFiles(@ref, "", Level + 1, ProjectName, ref RefFilesDictionary);
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


        public void WriteXmlButton_Click(System.Object sender, System.EventArgs e)
        {
            try
            {
                //Only create a new vault object
                //if one hasn't been created yet
                IEdmVault7 vault2 = null;
                if (vault1 == null)
                    vault1 = new EdmVault5();
                vault2 = (IEdmVault7)vault1;
                if (!vault1.IsLoggedIn)
                {
                    //Log into selected vault as the current user
                    vault1.LoginAuto(VaultsComboBox.Text, this.Handle.ToInt32());
                }

                //Get the file paths of all of the referenced files
                //and store them in RefFilesDictionary as keys;
                //the levels where they are found in the file hierarchy 
                //are stored as values
                Dictionary<string, string> RefFilesDictionary = new Dictionary<string, string>();
                foreach (string FileName in BatchRefListBox.Items)
                {
                    GetReferencedFiles(null, FileName, 0, "A", ref RefFilesDictionary);
                }

                //Because selecting a file in the Open File dialog 
                //adds the file and its references to the local cache, 
                //clear the local cache to demonstrate that the 
                //IEdmBatchListing methods don't add the files to 
                //the local cache 
                //Declare and create the IEdmClearLocalCache3 object
                IEdmClearLocalCache3 ClearLocalCache = default(IEdmClearLocalCache3);
                ClearLocalCache = (IEdmClearLocalCache3)vault2.CreateUtility(EdmUtility.EdmUtil_ClearLocalCache);
                ClearLocalCache.IgnoreToolboxFiles = true;
                ClearLocalCache.UseAutoClearCacheOption = true;

                //Declare and create the IEdmBatchListing object
                IEdmBatchListing BatchListing = default(IEdmBatchListing);
                BatchListing = (IEdmBatchListing)vault2.CreateUtility(EdmUtility.EdmUtil_BatchList);

                //Add all of the reference file paths to the 
                //IEdmClearLocalCache object to clear from the 
                //local cache and to the IEdmBatchListing object
                //to get the file information in batch mode
                foreach (KeyValuePair<string, string> KvPair in RefFilesDictionary)
                {
                    ClearLocalCache.AddFileByPath(KvPair.Key);
                    ((IEdmBatchListing2)BatchListing).AddFileCfg(KvPair.Key, DateTime.Now, (Convert.ToInt32(KvPair.Value)), "@", Convert.ToInt32(EdmListFileFlags.EdmList_Nothing));
                }
                //Clear the local cache of the reference files
                ClearLocalCache.CommitClear();

                //Create the batch file listing along with the file
                //card variables Description and Number
                EdmListCol[] BatchListCols = null;
                ((IEdmBatchListing2)BatchListing).CreateListEx("\n\nDescription\nNumber", Convert.ToInt32(EdmCreateListExFlags.Edmclef_MayReadFiles), ref BatchListCols, null);
                //Get the generated file information
                EdmListFile[] BatchListFiles = null;
                BatchListing.GetFiles(ref BatchListFiles);

                //Create the list where to store all the file information
                List<FileRef> FileRefs = new List<FileRef>();

                //Recursively add the file information to the list
                int j = 0;
                AddFileRef(ref BatchListFiles, ref j, 0, ref FileRefs);

                if (!Directory.Exists("C:\\temp"))
                {
                    MessageBox.Show("Directory \"c:\temp\" does " + "not exist; please create it and try again.");
                }
                else
                {
                    XmlSerializer XmlSer = new XmlSerializer(typeof(List<FileRef>));
                    StreamWriter StrWriter = new StreamWriter("C:\\temp\\BatchFileRefInfo.xml");
                    XmlSer.Serialize(StrWriter, FileRefs);
                    StrWriter.Close();
                    MessageBox.Show("File references successfully exported to an XML file.");
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


        public void AddFileRef(ref EdmListFile[] BatchListFiles, ref int curIndex, int curLevel, ref List<FileRef> FileRefs)

        {
            try
            {

                while (curIndex < BatchListFiles.Length)
                {
                    EdmListFile curListFile = BatchListFiles[curIndex];
                    //If the depth level of this listfile is <
                    //the current depth level then...
                    if (curListFile.mlParam > curLevel)
                    {
                        //Create a new FileRefs list
                        FileRefs[FileRefs.Count - 1].FileRefs = new List<FileRef>();
                        List<FileRef> fRef = FileRefs[FileRefs.Count - 1].FileRefs;
                        //Recurse using a new FileRefs list              
                        AddFileRef(ref BatchListFiles, ref curIndex, curListFile.mlParam, ref fRef);
                    }
                    else
                    {
                        //Create a new FileRef object to hold
                        //the file information
                        FileRef FileRef = new FileRef();
                        //Assign the FileRef properties
                        FileRef.CheckedOutBy = curListFile.mbsLockUser;
                        FileRef.CurrentState = curListFile.moCurrentState.mbsStateName;
                        object[] columnData = (object[])curListFile.moColumnData;
                        FileRef.Description = (string)columnData[0];
                        IEdmFile5 File = default(IEdmFile5);
                        File = (IEdmFile5)vault1.GetObject(EdmObjectType.EdmObject_File, curListFile.mlFileID);
                        FileRef.FileName = File.Name;
                        FileRef.LatestRevision = curListFile.mbsRevisionName;
                        FileRef.LatestVersion = Convert.ToString(curListFile.mlLatestVersion);
                        FileRef.Number = (string)columnData[1];
                        FileRef.FileRefs = null;
                        //Add the FileRef to this level's list
                        FileRefs.Add(FileRef);
                        curIndex += 1;
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

    }
}
