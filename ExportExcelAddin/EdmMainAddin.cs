using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
//using EPDM.Interop.epdm;
using EdmLib;
using Microsoft.Office.Interop.Excel;
using SldWorks;
using SwConst;

namespace ExportExcelAddin
{
    public class EdmMainAddin : IEdmAddIn5
    {
        public void GetAddInInfo(ref EdmAddInInfo poInfo, IEdmVault5 poVault, IEdmCmdMgr5 poCmdMgr)
        {
            //Specify information to display in the add-in's Properties dialog box
            poInfo.mbsAddInName = "ExportExcelAddin";
            poInfo.mbsCompany = "WISE Corp.";
            poInfo.mbsDescription = "Export a slddrawing BOM.";
            poInfo.mlAddInVersion = 3;

            //Specify the minimum required version of SolidWorks PDM Professional
            poInfo.mlRequiredVersionMajor = 6;
            poInfo.mlRequiredVersionMinor = 4;

            // Register a menu command
            poCmdMgr.AddCmd(1007, "明细表导出-科德", (int)(EdmMenuFlags.EdmMenu_OnlySingleSelection | EdmMenuFlags.EdmMenu_ContextMenuItem | EdmMenuFlags.EdmMenu_OnlyFiles), "导出明细表到制定文件");
            //poCmdMgr.AddCmd(1007, "明细表导出-科德", (int)EdmMenuFlags.EdmMenu_Nothing);
            poCmdMgr.AddCmd(1, "明细表导出-科德", (int)EdmMenuFlags.EdmMenu_Nothing , "导出明细表到制定文件");
            

        }
        public void OnCmd(ref EdmCmd poCmd, ref System.Array ppoData)
        {

        //public void OnCmd(ref EdmCmd poCmd, ref EdmCmdData[] ppoData)
        //{
            EdmVault5 vault = (EdmVault5)poCmd.mpoVault;
            if (poCmd.meCmdType == EdmCmdType.EdmCmd_Menu)
            {
                if (poCmd.mlCmdID == 1)
                {
                    if (ppoData.Length < 0)
                    {
                        System.Windows.Forms.MessageBox.Show("请选择一个文件");
                        return;
                    }
                    if (ppoData.Length > 1)
                    {
                        System.Windows.Forms.MessageBox.Show("只能选择一个文件");
                        return ;
                    }
                    int fileId = 0;
                    int parentFolderId = 0;
                    foreach (EdmCmdData AffectedFile in ppoData)
                    {
                        fileId = AffectedFile.mlObjectID1;
                        parentFolderId = AffectedFile.mlObjectID3;
                    }
                    if (fileId == 0)
                    {
                        System.Windows.Forms.MessageBox.Show("请选择一个文件");
                        return;
                    }
                    ExportExcelForm form = new ExportExcelForm();
                    
                    IEdmFile5 file= (IEdmFile5)vault.GetObject(EdmObjectType.EdmObject_File, fileId);
                    IEdmFolder5 folder = (IEdmFolder5)vault.GetObject(EdmObjectType.EdmObject_Folder, parentFolderId);
                    string filePath = folder.LocalPath + "\\" + file.Name;
                    
                    
                    form.selectFile = filePath;
                    if (form.ShowDialog() == DialogResult.OK)
                    {
                        string outFile = form.saveFilePath;
                        FormStatusBar statusBar = new FormStatusBar("明细表导出", "正在导出，请稍后....");
                        statusBar.UserCustomEvent += (obj) =>
                            {
                                exportExcel(vault, form.selectFile, form.saveFilePath);
                            };
                        statusBar.ShowDialog();
                    }

                    GC.Collect();
                    
                }
            }
        }
        public bool exportExcel(IEdmVault5 poVault, string selectFile, string saveFile)
        {
            WiseUtil wiseUtil = new WiseUtil();
            IEdmFolder5 templateExcelFolder = null;
            string finalFilePath = poVault.RootFolderPath + WiseUtil.templateExcelFilePath;
            IEdmFile5 templateExcelFile = poVault.GetFileFromPath(finalFilePath, out templateExcelFolder);
            if (templateExcelFile == null)
            {
                System.Windows.Forms.MessageBox.Show("在PDM中找不到模板EXCEL：\n" + finalFilePath);
                return false;
            }
            
            //string installLocation = wiseUtil.GetLocalMachineRegistryValue("SOFTWARE\\solidworks\\Applications\\PDMWorks Enterprise\\", "Location");
            //System.Console.WriteLine(installLocation);
            //if (installLocation == "")
            //{
            //   installLocation = "C:\\";
            //}
            //string excelTemplate = installLocation + "ExportTemplate\\template-list.xlsx";
            Workbook templateWb = null;
            try { 
                templateWb = wiseUtil.OpenExcel(finalFilePath);
            }
            catch (Exception e) {
                System.Windows.Forms.MessageBox.Show("打开模板表格失败，请查看位置" + templateExcelFile+"的文件状态");
                return false;
            }
            if (templateWb == null)
            {
                System.Windows.Forms.MessageBox.Show("打开模板表格失败");
                return false;
            }
            SldWorks.SldWorks swApp = new SldWorks.SldWorks();
            int longstatus = 0;
            int longwarnings = 0;
            ModelDoc2 modelDoc = swApp.OpenDoc6(selectFile, (int)swDocumentTypes_e.swDocDRAWING,
                (int)(swOpenDocOptions_e.swOpenDocOptions_ReadOnly | swOpenDocOptions_e.swOpenDocOptions_Silent), 
                "", ref longstatus, ref longwarnings);
            if (modelDoc == null || longstatus > 0)
            {
                System.Windows.Forms.MessageBox.Show("打开二维图失败");
                swApp.ExitApp();
                templateWb.Close();
                return false;
            }
            BomTableAnnotation bomTableAnno = null;
            string configName = "";
            string topFileName = "";
            try
            {
                wiseUtil.GetDrawingDocBOMTable(modelDoc, out bomTableAnno, out configName, out topFileName);
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("获取明细表信息出错"+e.StackTrace);
            }
            if (bomTableAnno == null)
            {
                System.Windows.Forms.MessageBox.Show("未找到明细表信息");
                swApp.ExitApp();
                templateWb.Close();
                return false;
            }
            //modelDoc.Close();
            SldAsm asmPrd = wiseUtil.GetAsmIndoFromFile(poVault, topFileName);// new SldAsm();
            if (asmPrd == null)
            {
                System.Windows.Forms.MessageBox.Show("获取材料明细表关联产品失败");
                swApp.ExitApp();
                templateWb.Close();
                return false;
            }
            //asmPrd.bzr = "hou";
            //asmPrd.bzsj = "2016/1/1";
            try
            {
                wiseUtil.ProcessTableAnn(poVault, bomTableAnno, configName, asmPrd);
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("解析明细表信息出错" + e.StackTrace);
            }
            //if (WiseUtil.errors > 0)
            //{
            //    System.Windows.Forms.MessageBox.Show("解析明细表信息出错" + WiseUtil.errorStr);
            //}
            wiseUtil.SaveBuyToWorkbook(templateWb, asmPrd);
            wiseUtil.SaveStdToWorkbook(templateWb, asmPrd);
            wiseUtil.SavePrtToWorkbook(templateWb, asmPrd);
            wiseUtil.SaveBspToWorkbook(templateWb, asmPrd);
            templateWb.SaveAs(saveFile);
            templateWb.Close();
            swApp.ExitApp();
            swApp = null;
            return true;
        }
    }
}
