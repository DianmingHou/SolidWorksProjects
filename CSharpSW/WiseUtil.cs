using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using Microsoft.Win32;
using Microsoft.Office.Interop.Excel;
using SldWorks;
using SwConst;
using EPDM.Interop.epdm;
namespace CSharpSW
{
    class SldBsp
    {
        public string path = "";
        public string parentFile = "";
        public string materialNumber = "";
        public string number = "";
        public string name = "";
        public int amout = 0;
        public string material = "";
        public string remark = "";
        public double weight = 0.0D;
        public double totalWeight = 0.0D;
        public int type = 0;
    }
    class SldPrt:SldBsp
    {
        public double rawWeight = 0.0D;
        public string route = "";
    }
    class SldBuy : SldBsp { }
    class SldStd : SldBsp { }
    class SldAsm : SldBsp
    {
        public string ztdm = "";//总图代码，代号第一位
        public string zjdm = "";//组件代码，第二位代号
        public string bzr = "";//编制人
        public string bzsj = "";//编制时间
        public string pzr = "";//批准人
        public string pzsj = "";//批准时间
        public string jdbj = "";//阶段标记
        public string sbxh = "";//设备型号
        public string projName = "";//项目名称
        public List<SldPrt> sldPrtList = new List<SldPrt>();
        public List<SldStd> sldStdList = new List<SldStd>();
        public List<SldBuy> sldBuyList = new List<SldBuy>();
    }
    /// <summary>Class <c>WiseUtil</c> 公共方法，提供表格读取，注册表，文档读取等功能
    /// .</summary>
    class WiseUtil
    {
        static string stdStr = "标准件目录";
        static string prtStr = "基本件目录";
        static string buyStr = "外购件目录";
        static public int stdPageSize = 32;
        static public int buyPageSize = 32;
        static public int prtPageSize = 10;//19
        public Workbook OpenExcel(string fileName)
        {
            Application app = new Application();
            Workbooks wbks = app.Workbooks;
            return wbks.Open(fileName);
        }
        public string GetLocalMachineRegistryValue(string iSubKey,string iValue)
        {
            RegistryKey registryKey = Registry.LocalMachine.OpenSubKey(iSubKey);
            return registryKey.GetValue(iValue).ToString();
        }
        ///<summary> method <c>GetDrawingDocBOMTable</c>
        ///从SolidWorks 的Drawing文档获取到其材料明细表，一般一个文档一个，多余不获取
        ///</summary>
        public void GetDrawingDocBOMTable(ModelDoc2 modelDoc, out BomTableAnnotation swBOMTable, out string configName,out string topFileName)
        {
            Feature feature = modelDoc.IFirstFeature();

            swBOMTable = null;
            configName = "";
            BomFeature swBomFeat = null;
            while (feature != null)
            {
                Type type = feature.GetType();
                if ("BomFeat" == feature.GetTypeName())
                {
                    Console.WriteLine("******************************");
                    Console.WriteLine("Feature Name : " + feature.Name);

                    swBomFeat = (BomFeature)feature.GetSpecificFeature2();
                    break;
                }

                string name = feature.Name;
                feature = feature.IGetNextFeature();
            }
            if (swBomFeat == null||swBomFeat.GetTableAnnotationCount()<=0)
            {
                return;
            }
            swBOMTable = swBomFeat.IGetTableAnnotations(1);// default(TableAnnotation);
            topFileName = swBomFeat.GetReferencedModelName();
            configName = "";
            if (configCount > 0)
            {
                bool visibility = true;
                configName = swBomFeat.IGetConfigurations(true, 0, ref visibility);
            }
            //swBOMTable.Iget
            return;
        }

        public void ProcessTableAnn(BomTableAnnotation swBOMTableAnn, string ConfigName, SldAsm asmPrd)
        {
            int nNumRow = 0;
            string ItemNumber = null;
            string PartNumber = null;
            TableAnnotation swTableAnn = (TableAnnotation)swBOMTableAnn;

            Console.WriteLine("   Table Title        " + swTableAnn.Title);
            IEdmVault5 vault = new EdmVault5();
            vault.LoginAuto("科德研发部", 0);
            if (!vault.IsLoggedIn)
            {
                Console.WriteLine("登录PDM失败");
                return;
            }
            nNumRow = swTableAnn.RowCount;
            swBOMTableAnn = (BomTableAnnotation)swTableAnn;
            for (int j = 0; j < swTableAnn.RowCount; j++)
            {
                //
                //for (int i = 0; i < swTableAnn.ColumnCount;i++ )
                //    Console.WriteLine(swTableAnn.get_Text(j, i));
                //获取类别和特有信息
                //if (j == swTableAnn.RowCount - 1)//最后一行为标题栏，跳过
               //     continue;
                string filePath = "";
                string compName = "";
                string compConfig = "";
                int compCount = swBOMTableAnn.GetComponentsCount2(j, ConfigName, out ItemNumber, out PartNumber);
                for (int i = 0; i < compCount; i++)
                {
                    Component2 comp2 = swBOMTableAnn.IGetComponents2(j, ConfigName, i);
                    if (comp2 != null)
                    {
                        filePath = comp2.GetPathName();
                        compName = comp2.Name2;
                        compConfig =  comp2.ReferencedConfiguration;
                        Console.WriteLine("           Component Name :" + comp2.Name2 + "      Configuration Name : " + comp2.ReferencedConfiguration);
                        Console.WriteLine("           Component Path :" + comp2.GetPathName());
                    }
                }
                if (filePath == "")
                    continue;
                SldBsp bsp  = GetSldPrdInfoFromFile(vault, filePath);
                if (bsp == null)
                    continue;
                bsp.path = filePath;
                //SldBsp bsp = new SldBsp();
                swTableAnn.get_Text(j, 0);//序号
                bsp.number = swTableAnn.get_Text(j, 1);//代号
                bsp.name = swTableAnn.get_Text(j, 2);//名称
                string amout = swTableAnn.get_Text(j, 3);
                bsp.amout = amout==""?0:int.Parse(amout);//数量
                bsp.material = swTableAnn.get_Text(j, 4);//材料
                string weight = swTableAnn.get_Text(j, 5);
                bsp.weight = weight == "" ? 0 : int.Parse(weight);//单重
                bsp.totalWeight = bsp.weight * bsp.amout;// swTableAnn.get_Text(j, 6);//总重
                bsp.remark = swTableAnn.get_Text(j, 7);//备注
                string number = swTableAnn.get_Text(j, 8);//测试
                if (bsp.number == "") bsp.number = number;
                if (bsp is SldPrt)
                {
                    asmPrd.sldPrtList.Add((SldPrt)bsp);
                }
                else if (bsp is SldStd)
                {
                    asmPrd.sldStdList.Add((SldStd)bsp);
                }else if(bsp is SldBuy){
                    asmPrd.sldBuyList.Add((SldBuy)bsp);
                }
                
            }
            return;
        }
        public bool GetAsmIndoFromFile(IEdmVault5 poVault, string filePath, SldAsm asmPrd)
        {
            IEdmFolder5 opParentFolder;
            IEdmFile5 poFile = poVault.GetFileFromPath(filePath, out opParentFolder);
            IEdmEnumeratorVariable5 enumVar = poFile.GetEnumeratorVariable();
            if (filePath == "")
                return false;
            asmPrd.path = filePath;
            //
            object tmpVar;
            if (enumVar.GetVar("代号", "@", out tmpVar))
            {
                asmPrd.number = tmpVar.ToString();
                asmPrd.ztdm = asmPrd.number.Substring(0,asmPrd.number.IndexOf("."));
                asmPrd.zjdm = asmPrd.number.Substring(asmPrd.number.IndexOf(".") + 1);
            }
            if (enumVar.GetVar("名称", "@", out tmpVar))
                asmPrd.name = tmpVar.ToString();
            if (enumVar.GetVar("设计", "@", out tmpVar))
                asmPrd.bzr = tmpVar.ToString();
            if (enumVar.GetVar("设计日期", "@", out tmpVar))
            {
                asmPrd.bzsj = ((DateTime)tmpVar).ToString("%y/%m/%d");
            }
            if (enumVar.GetVar("批准", "@", out tmpVar))
                asmPrd.pzr = tmpVar.ToString();
            if (enumVar.GetVar("批准日期", "@", out tmpVar))
            {
                asmPrd.pzsj = ((DateTime)tmpVar).ToString("%y/%m/%d");
            }
            if (enumVar.GetVar("阶段标记", "@", out tmpVar))
                asmPrd.jdbj = tmpVar.ToString();

            if (enumVar.GetVar("设备型号", "@", out tmpVar))
                asmPrd.sbxh = tmpVar.ToString();
            if (enumVar.GetVar("Project Name", "@", out tmpVar))
                asmPrd.projName = tmpVar.ToString();
            return true;
        }
        public SldBsp GetSldPrdInfoFromFile(IEdmVault5 poVault, string filePath)
        {
            SldBsp bspPrt = null;
            IEdmFolder5 opParentFolder;
            //bspPrt = new SldBuy();
            IEdmFile5 poFile = poVault.GetFileFromPath(filePath, out opParentFolder);
            if (poFile == null)
            {
                bspPrt = new SldPrt();
                bspPrt.type = 1;
                return bspPrt;
            }
            IEdmEnumeratorVariable5 enumVar = poFile.GetEnumeratorVariable();
            if (enumVar == null) return null;
            
            object tmpVar;
            string partType = "";
            if (enumVar.GetVar("Part Type", "@", out tmpVar))
            {
                partType = tmpVar.ToString();
            }
            if (partType == "自制件") {

                bspPrt = new SldPrt();
                bspPrt.type = 1;
                SldPrt prt = (SldPrt)bspPrt;
                if (enumVar.GetVar("工艺路线", "@", out tmpVar))
                {
                    prt.route = tmpVar.ToString();
                }

            }
            else if (partType == "标准件")
            {
                bspPrt = new SldStd();
                bspPrt.type = 2;
            }
            else
            {
                bspPrt = new SldBuy();
                bspPrt.type = 3;
            }
            bspPrt.path = filePath;
            return bspPrt;

        }

        public bool SavePrtToWorkbook(Workbook workbook, SldAsm asmPrd)
        {
            if(asmPrd.sldPrtList.Count<=0&&asmPrd.sldStdList.Count<=0&&asmPrd.sldBuyList.Count<=0)
                return true;
            int count = asmPrd.sldPrtList.Count;
            int perPageSize = WiseUtil.prtPageSize;//
            int pageSize = count / perPageSize + (count % perPageSize > 0 ? 1 : 0);

            bool insHead = true;
            //空行
            if (insHead)
            {
                SldPrt prt = new SldPrt();
                asmPrd.sldPrtList.Insert(0, prt);
                count++;
            }
            if (insHead && asmPrd.sldStdList.Count > 0)
            {
                SldPrt prt = new SldPrt();
                prt.name = WiseUtil.stdStr;
                int thisCount = asmPrd.sldStdList.Count;
                prt.amout = thisCount / WiseUtil.stdPageSize + (thisCount % WiseUtil.stdPageSize > 0 ? 1 : 0);
                asmPrd.sldPrtList.Insert(0, prt);
                count++;
            }
            if (insHead && asmPrd.sldBuyList.Count>0)
            {
                SldPrt prt = new SldPrt();
                prt.name = WiseUtil.buyStr;
                int thisCount = asmPrd.sldBuyList.Count;
                prt.amout = thisCount / WiseUtil.buyPageSize + (thisCount % WiseUtil.buyPageSize > 0 ? 1 : 0);
                asmPrd.sldPrtList.Insert(0, prt);
                count++;
            }
            //最初count
            if (insHead && asmPrd.sldPrtList.Count > 0)
            {
                SldPrt prt = new SldPrt();
                prt.name = WiseUtil.prtStr;
                
                count++;//先加一个

                pageSize = count / WiseUtil.prtPageSize + (count % prtPageSize > 0 ? 1 : 0);
                prt.amout = pageSize;
                asmPrd.sldPrtList.Insert(0, prt);
            }


            pageSize = count / perPageSize + (count % perPageSize > 0 ? 1 : 0);


            int rowStartIndex = 4;

            Sheets sheets = workbook.Worksheets;
            Worksheet worksheet = (Worksheet)sheets.get_Item(WiseUtil.prtStr);
            if (count <= 0)
            {
                worksheet.Delete();
                return false;
            }
            List<SldPrt>.Enumerator iEnum = asmPrd.sldPrtList.GetEnumerator();
            Worksheet nextWorksheet = null;
            for (int pageIndex = 0; pageIndex < pageSize; pageIndex++)
            {
                if (pageIndex < pageSize - 1)//最后一个不能再复制了,写sheet前复制，保证模板未被写
                {
                    worksheet.Copy(Type.Missing, worksheet);
                    nextWorksheet = worksheet.Next;
                }
                Range dyn = worksheet.Cells.get_Item("31", "A");// worksheet.Cells.get_Item("25", "A").set_Value(asmPrd.ztdm);
                worksheet.Range["A25"].Value = asmPrd.ztdm;
                worksheet.Range["A31"].Value = asmPrd.zjdm;
                worksheet.Range["E30"].Value = asmPrd.bzr;
                worksheet.Range["H30"].Value = asmPrd.bzsj;
                worksheet.Range["E31"].Value = asmPrd.pzr;
                worksheet.Range["H31"].Value = asmPrd.pzsj;
                if (asmPrd.jdbj == "S")
                    worksheet.Range["K29"].Value = asmPrd.jdbj;
                else if (asmPrd.jdbj == "A")
                    worksheet.Range["M29"].Value = asmPrd.jdbj;
                else if (asmPrd.jdbj=="B")
                    worksheet.Range["O29"].Value = asmPrd.jdbj;
                else
                    worksheet.Range["R29"].Value = asmPrd.jdbj;
                worksheet.Range["K31"].Value = "共 " + pageSize + " 页";
                worksheet.Range["S31"].Value = "第 " + (pageIndex+1) + " 页";
                worksheet.Range["W31"].Value = asmPrd.sbxh;
                int row = 0;
                while (iEnum.MoveNext())
                {
                    
                    SldPrt prt = iEnum.Current;
                    int totalIndex = row+rowStartIndex;
                    worksheet.Range["B"+totalIndex].Value = prt.materialNumber;
                    worksheet.Range["E" + totalIndex].Value = prt.number;
                    worksheet.Range["J" + totalIndex].Value = prt.name;
                    worksheet.Range["O" + totalIndex].Value = prt.amout == 0 ? "" : prt.amout.ToString();
                    worksheet.Range["Q" + totalIndex].Value = prt.material;
                    worksheet.Range["S" + totalIndex].Value = prt.weight == 0 ? "" : prt.weight.ToString();
                    worksheet.Range["T" + totalIndex].Value = prt.totalWeight == 0 ? "" : prt.totalWeight.ToString();
                    worksheet.Range["U" + totalIndex].Value = prt.rawWeight == 0 ? "" : prt.rawWeight.ToString();
                    worksheet.Range["V" + totalIndex].Value = prt.route;
                    worksheet.Range["AC" + totalIndex].Value = prt.remark;

                    row++;
                    if (row >= WiseUtil.prtPageSize)
                        break;
                }

                worksheet = nextWorksheet;
            }
            return true;
        }
        public bool SaveStdToWorkbook(Workbook workbook, SldAsm asmPrd)
        {
            int count = asmPrd.sldStdList.Count;
            int perPageSize = WiseUtil.stdPageSize;//
            string pageName = WiseUtil.stdStr;

            int pageSize = count / perPageSize + (count % perPageSize > 0 ? 1 : 0);
            int rowStartIndex = 4;
            Sheets sheets = workbook.Worksheets;
            Worksheet worksheet = (Worksheet)sheets.get_Item(pageName);
            if (count <= 0)
            {
                worksheet.Delete();
                return false;
            }
            List<SldStd>.Enumerator iEnum = asmPrd.sldStdList.GetEnumerator();
            Worksheet nextWorksheet = null;
            for (int pageIndex = 0; pageIndex < pageSize; pageIndex++)
            {
                if (pageIndex < pageSize - 1)//最后一个不能再复制了
                {
                    worksheet.Copy(Type.Missing, worksheet);
                    nextWorksheet = worksheet.Next;
                }
                Range dyn = worksheet.Cells.get_Item("31", "A");// worksheet.Cells.get_Item("25", "A").set_Value(asmPrd.ztdm);
                worksheet.Range["A38"].Value = asmPrd.ztdm;
                worksheet.Range["A44"].Value = asmPrd.zjdm;
                worksheet.Range["E43"].Value = asmPrd.bzr;
                worksheet.Range["H43"].Value = asmPrd.bzsj;
                worksheet.Range["E44"].Value = asmPrd.pzr;
                worksheet.Range["H44"].Value = asmPrd.pzsj;
                if (asmPrd.jdbj == "S")
                    worksheet.Range["K42"].Value = asmPrd.jdbj;
                else if (asmPrd.jdbj == "A")
                    worksheet.Range["L42"].Value = asmPrd.jdbj;
                else if (asmPrd.jdbj == "B")
                    worksheet.Range["M42"].Value = asmPrd.jdbj;
                else
                    worksheet.Range["N42"].Value = asmPrd.jdbj;
                worksheet.Range["K44"].Value = "共 " + pageSize + " 页";
                worksheet.Range["P44"].Value = "第 " + (pageIndex + 1) + " 页";
                worksheet.Range["S44"].Value = asmPrd.sbxh;
                int row = 0;
                while (iEnum.MoveNext())
                {

                    SldStd prt = iEnum.Current;
                    int totalIndex = row + rowStartIndex;
                    worksheet.Range["B" + totalIndex].Value = prt.materialNumber;
                    worksheet.Range["E" + totalIndex].Value = prt.number;
                    worksheet.Range["J" + totalIndex].Value = prt.name;
                    worksheet.Range["O" + totalIndex].Value = prt.amout == 0 ? "" : prt.amout.ToString();
                    worksheet.Range["Q" + totalIndex].Value = prt.material;
                    worksheet.Range["S" + totalIndex].Value = prt.weight == 0 ? "" : prt.weight.ToString();
                    worksheet.Range["T" + totalIndex].Value = prt.totalWeight == 0 ? "" : prt.totalWeight.ToString();
                    worksheet.Range["U" + totalIndex].Value = prt.remark;

                    row++;
                    if (row >= perPageSize)
                        break;
                } 
                worksheet = nextWorksheet;
            }
            return true;
        }
        public bool SaveBuyToWorkbook(Workbook workbook, SldAsm asmPrd)
        {
            int count = asmPrd.sldBuyList.Count;
            int perPageSize = WiseUtil.buyPageSize;//
            string pageName = WiseUtil.buyStr;

            int pageSize = count / perPageSize + (count % perPageSize > 0 ? 1 : 0);
            int rowStartIndex = 4;
            Sheets sheets = workbook.Worksheets;
            Worksheet worksheet = (Worksheet)sheets.get_Item(pageName);
            if (count <= 0)
            {
                worksheet.Delete();
                return true;
            }
            List<SldBuy>.Enumerator iEnum = asmPrd.sldBuyList.GetEnumerator();
            Worksheet nextWorksheet = null;
            for (int pageIndex = 0; pageIndex < pageSize; pageIndex++)
            {
                if (pageIndex < pageSize - 1)//最后一个不能再复制了
                {
                    worksheet.Copy(Type.Missing, worksheet);
                    nextWorksheet = worksheet.Next;
                }
                Range dyn = worksheet.Cells.get_Item("31", "A");// worksheet.Cells.get_Item("25", "A").set_Value(asmPrd.ztdm);
                worksheet.Range["A38"].Value = asmPrd.ztdm;
                worksheet.Range["A44"].Value = asmPrd.zjdm;
                worksheet.Range["E43"].Value = asmPrd.bzr;
                worksheet.Range["H43"].Value = asmPrd.bzsj;
                worksheet.Range["E44"].Value = asmPrd.pzr;
                worksheet.Range["H44"].Value = asmPrd.pzsj;
                if (asmPrd.jdbj == "S")
                    worksheet.Range["K42"].Value = asmPrd.jdbj;
                else if (asmPrd.jdbj == "A")
                    worksheet.Range["L42"].Value = asmPrd.jdbj;
                else if (asmPrd.jdbj == "B")
                    worksheet.Range["M42"].Value = asmPrd.jdbj;
                else
                    worksheet.Range["N42"].Value = asmPrd.jdbj;
                worksheet.Range["K44"].Value = "共 " + pageSize + " 页";
                worksheet.Range["P44"].Value = "第 " + (pageIndex + 1) + " 页";
                worksheet.Range["S44"].Value = asmPrd.sbxh;
                int row = 0;
                while (iEnum.MoveNext())
                {

                    SldBuy prt = iEnum.Current;
                    int totalIndex = row + rowStartIndex;
                    worksheet.Range["B" + totalIndex].Value = prt.materialNumber;
                    worksheet.Range["E" + totalIndex].Value = prt.number;
                    worksheet.Range["J" + totalIndex].Value = prt.name;
                    worksheet.Range["O" + totalIndex].Value = prt.amout == 0 ? "" : prt.amout.ToString();
                    worksheet.Range["Q" + totalIndex].Value = prt.material;
                    worksheet.Range["S" + totalIndex].Value = prt.weight == 0 ? "" : prt.weight.ToString();
                    worksheet.Range["T" + totalIndex].Value = prt.totalWeight == 0 ? "" : prt.totalWeight.ToString();
                    worksheet.Range["U" + totalIndex].Value = prt.remark;

                    row++;
                    if (row >= perPageSize)
                        break;
                }
                worksheet = nextWorksheet;
            }
            return true;
        }
        
    }
    
}
