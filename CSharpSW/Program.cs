using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SldWorks;
using SwConst;
using System.Diagnostics;
using Microsoft.Office.Interop.Excel;
using EPDM.Interop.epdm;


namespace CSharpSW
{
    class Program
    {
        static void Main(string[] args)
        {
            WiseUtil wiseUtil = new WiseUtil();
            string installLocation = wiseUtil.GetLocalMachineRegistryValue("SOFTWARE\\solidworks\\Applications\\PDMWorks Enterprise\\", "Location");
            System.Console.WriteLine(installLocation);
            if (installLocation == "")
            {
                installLocation = "C:\\";
            }
            string excelTemplate = installLocation + "ExportTemplate\\template-list.xlsx";
            Workbook templateWb = wiseUtil.OpenExcel(excelTemplate);
            SldWorks.SldWorks swApp;

            //IEdmVault5 vault = new EdmVault5();
            //vault.LoginAuto("科德研发部");
            swApp = new SldWorks.SldWorks();
            int longstatus = 0;
            int longwarnings = 0;
            //E:\\wisevault0\\draw\\G20-A01P.SLDDRW
            //F:\\科德研发部\\01-机床整机产品\\01-铣车复合机\\KMC800 系列机型\\02-裸机图纸\\KMC800\\710.14 滑枕组件\\710.1401 滑枕组件装配图基础 A2.SLDDRW
            ModelDoc2 modelDoc = swApp.OpenDoc6("E:\\wisevault0\\draw\\G20-A01P.SLDDRW", (int)swDocumentTypes_e.swDocDRAWING, (int)swOpenDocOptions_e.swOpenDocOptions_ReadOnly, "", ref longstatus, ref longwarnings);
            BomTableAnnotation bomTableAnno;
            string configName = "";
            string topFileName = "";
            wiseUtil.GetDrawingDocBOMTable(modelDoc, out bomTableAnno, out configName, out topFileName);

            //modelDoc.Close();
            IEdmVault5 poVault = null;
            SldAsm asmPrd = wiseUtil.GetAsmIndoFromFile(poVault, topFileName);// new SldAsm();
            if (asmPrd == null)
            {
                return false;
            }
            //asmPrd.bzr = "hou";
            //asmPrd.bzsj = "2016/1/1";
            wiseUtil.ProcessTableAnn(poVault, bomTableAnno, configName, asmPrd);

            asmPrd.bzr = "hou";
            asmPrd.bzsj = "2016/1/1";
            wiseUtil.ProcessTableAnn(bomTableAnno, configName,asmPrd);

            wiseUtil.SaveBuyToWorkbook(templateWb, asmPrd);
            wiseUtil.SaveStdToWorkbook(templateWb, asmPrd);
            wiseUtil.SavePrtToWorkbook(templateWb, asmPrd);
            templateWb.SaveAs("D:\\a.xlsx");
            templateWb.Close();
            swApp.ExitApp();

            swApp = null;

        }

       

        
    }
}
