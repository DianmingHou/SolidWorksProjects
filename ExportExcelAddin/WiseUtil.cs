using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using Microsoft.Win32;
using Microsoft.Office.Interop.Excel;
using SldWorks;
using SwConst;
using EdmLib;
//using EPDM.Interop.epdm;

namespace ExportExcelAddin
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
        public int type = 0;//1位基本件，2位标准件，3位外购件，4位未定义，5位未找到零件
    }
    class SldPrt : SldBsp
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
        public List<SldPrt> sldPrtList = new List<SldPrt>();//存储自制件
        public List<SldStd> sldStdList = new List<SldStd>();//存储标准件
        public List<SldBuy> sldBuyList = new List<SldBuy>();//存储外购件
        public List<SldBsp> sldBspList = new List<SldBsp>();//存储未识别的零件，独立的放到一个sheet中
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
        static public int prtPageSize = 19;//19

        static public int errors = 0;
        static public string errorStr = "";
        static public string templateExcelFilePath = "\\04-企业标准化\\03-软件项目\\SolidWorks及PDM\\PDM\\template-list.xlsx";

        
        public Workbook OpenExcel(string fileName)
        {
            Application app = new Application();
            Workbooks wbks = app.Workbooks;
            if (app != null)
            {

                System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
                app = null;
            }
            
            return wbks.Open(fileName);
        }
        public string GetLocalMachineRegistryValue(string iSubKey, string iValue)
        {
            RegistryKey registryKey = Registry.LocalMachine.OpenSubKey(iSubKey);
            return registryKey.GetValue(iValue).ToString();
        }
        ///<summary> method <c>GetDrawingDocBOMTable</c>
        ///从SolidWorks 的Drawing文档获取到其材料明细表，一般一个文档一个，多余不获取
        ///</summary>
        public void GetDrawingDocBOMTable(ModelDoc2 modelDoc, out BomTableAnnotation swBOMTable, out string configName, out string topFileName)
        {

            Feature feature = modelDoc.IFirstFeature();

            swBOMTable = null;
            configName = "";
            topFileName = "";
            BomFeature swBomFeat = null;
            while (feature != null)
            {
                Type type = feature.GetType();
                if ("BomFeat" == feature.GetTypeName())
                {
                    //Console.WriteLine("******************************");
                    //Console.WriteLine("Feature Name : " + feature.Name);

                    swBomFeat = (BomFeature)feature.GetSpecificFeature2();
                    break;
                }

                string name = feature.Name;
                feature = feature.IGetNextFeature();
            }
            if (swBomFeat == null || swBomFeat.GetTableAnnotationCount() <= 0)
            {
                return;
            }
            Feature swFeat = (Feature)swBomFeat.GetFeature();
            string featureName = swFeat.Name;
            swBOMTable = swBomFeat.IGetTableAnnotations(1);// default(TableAnnotation);
            topFileName = swBomFeat.GetReferencedModelName();
            int configCount = swBomFeat.GetConfigurationCount(true);
            //
            //object[] vTableArr = (object[])swBomFeat.GetTableAnnotations();
            //foreach (object vTable_loopVariable in vTableArr)
            //{
            //    object vTable = vTable_loopVariable;
            //    TableAnnotation swTable = (TableAnnotation)vTable;
            //    ProcessTableAnn(modelDoc, swTable);
            //}

            configName = "";
            if (configCount > 0)
            {
                bool visibility = true;
                configName = swBomFeat.IGetConfigurations(true, 0, ref visibility);
            }
            //swBOMTable.Iget
            return;
        }

        public void ProcessTableAnn(IEdmVault5 poVault, BomTableAnnotation swBOMTableAnn, string ConfigName, SldAsm asmPrd)
        {
            int nNumRow = 0;
            string ItemNumber = null;
            string PartNumber = null;
            TableAnnotation swTableAnn = (TableAnnotation)swBOMTableAnn;
            
            nNumRow = swTableAnn.RowCount;
            swBOMTableAnn = (BomTableAnnotation)swTableAnn;
            for (int j = 0; j < swTableAnn.RowCount; j++)
            {
                //
                //for (int i = 0; i < swTableAnn.ColumnCount;i++ )
                //    Console.WriteLine(swTableAnn.get_Text(j, i));
                //获取类别和特有信息
                if (j == swTableAnn.RowCount - 1)//最后一行为标题栏，跳过
                     continue;
                string filePath = "";
                string compName = "";
                string compConfig = "";
                int compCount = swBOMTableAnn.GetComponentsCount2(j, ConfigName, out ItemNumber, out PartNumber);
                object[] vPtArr  = swBOMTableAnn.GetComponents2(j, ConfigName);
                if (vPtArr!=null)
                {
                    for (int prIndex = 0; prIndex <= vPtArr.GetUpperBound(0); prIndex++)
                    {
                        Component2 swComp = (Component2)vPtArr[prIndex];
                        if ((swComp != null))
                        {
                            //Debug.Print("           Component Name :" + swComp.Name2 + "      Configuration Name : " + swComp.ReferencedConfiguration);
                            //Debug.Print("           Component Path :" + swComp.GetPathName());
                            filePath = swComp.GetPathName();
                            compName = swComp.Name2;
                            compConfig = swComp.ReferencedConfiguration;
                            break;
                        }
                        else
                        {
                            continue;
                        }
                    }

                    //Component2 comp2 = null;
                    //comp2 = swBOMTableAnn.IGetComponents2(j, ConfigName, 0);
                    //if (comp2 != null)
                    //{
                    //    filePath = comp2.GetPathName();
                    //    compName = comp2.Name2;
                    //    compConfig = comp2.ReferencedConfiguration;
                    //}
                }

                SldBsp bsp = null;
                if (filePath == "")
                {
                    errors++;
                    errorStr += "明细表第" + (j + 1) + "行" + swTableAnn.get_Text(j, 1) + " 未找到关联文件\n";
                    bsp = new SldBsp();
                    bsp.type = 5;
                }
                else
                {
                    bsp = GetSldPrdInfoFromFile(poVault, filePath);
                    if (bsp == null)
                    {
                        errors++;
                        errorStr += "明细表第" + (j + 1) + "行" + swTableAnn.get_Text(j, 1) + " 文件：" + filePath + " 在PDM中未找到\n";
                        bsp = new SldBsp();
                        bsp.type = 5;
                    }
                }
                
                bsp.path = filePath;
                //SldBsp bsp = new SldBsp();
                swTableAnn.get_Text(j, 0);//序号
                bsp.number = swTableAnn.get_Text(j, 1);//代号
                bsp.name = swTableAnn.get_Text(j, 2);//名称
                string amout = swTableAnn.get_Text(j, 3);
                int parseInt = 0;
                try
                {
                    parseInt = int.Parse(amout);
                }catch(Exception){
                    parseInt = 0;
                }
                bsp.amout = parseInt;//数量
                bsp.material = swTableAnn.get_Text(j, 4);//材料
                string weight = swTableAnn.get_Text(j, 5);
                double parseDouble = 0.0;
                try
                {
                    parseDouble = double.Parse(weight);
                }
                catch (Exception)
                {
                    parseDouble = 0.0;
                }
                bsp.weight = parseDouble;//单重
                bsp.totalWeight = bsp.weight * bsp.amout;// swTableAnn.get_Text(j, 6);//总重
                bsp.remark = swTableAnn.get_Text(j, 7);//备注
                //string number = swTableAnn.get_Text(j, 8);//测试
                //if (bsp.number == "") bsp.number = number;
                if (bsp is SldPrt)
                {
                    bsp.materialNumber = bsp.number==""?"":("03."+bsp.number);
                    asmPrd.sldPrtList.Add((SldPrt)bsp);
                }
                else if (bsp is SldStd)
                {
                    asmPrd.sldStdList.Add((SldStd)bsp);
                }
                else if (bsp is SldBuy)
                {
                    asmPrd.sldBuyList.Add((SldBuy)bsp);
                }
                else if (bsp is SldBsp)
                {
                    asmPrd.sldBspList.Add(bsp);
                }

            }
            return;
        }
        public SldAsm GetAsmIndoFromFile(IEdmVault5 poVault, string filePath)
        {
            IEdmFolder5 opParentFolder;
            IEdmFile5 poFile = poVault.GetFileFromPath(filePath, out opParentFolder);
            if (poFile == null)
                return null;
            IEdmVariableMgr5 varMgr = (IEdmVariableMgr5)poVault;
            if (varMgr == null)
                return null;

            IEdmEnumeratorVariable5 enumVar = poFile.GetEnumeratorVariable();
            if (enumVar==null)
                return null;
            if (filePath == "")
                return null;
            EdmStrLst5 cfgList = poFile.GetConfigurations();

            SldAsm asmPrd = new SldAsm();
            asmPrd.path = filePath;
            asmPrd.amout = 1;
            

            if (varMgr.GetVariable("代号") != null)
            {
                object tmpVar = null ;
                IEdmPos5 pos = cfgList.GetHeadPosition();
                string cfgName = null;
                while (!pos.IsNull)
                {
                   cfgName = cfgList.GetNext(pos);
                   enumVar.GetVar("代号", cfgName, out tmpVar);
                   if (tmpVar == null) continue;
                   if (!tmpVar.ToString().Equals(""))break;
                }
                if (tmpVar != null) asmPrd.number = tmpVar.ToString();
                if (asmPrd.number.IndexOf(".") > 0 && asmPrd.number.IndexOf(".") < asmPrd.number.Length)
                {
                    asmPrd.ztdm = asmPrd.number.Substring(0, asmPrd.number.IndexOf("."));
                    asmPrd.zjdm = asmPrd.number.Substring(asmPrd.number.IndexOf(".") + 1);
                }
                else {
                    asmPrd.ztdm = asmPrd.zjdm = asmPrd.number;
                }
            }
            if (varMgr.GetVariable("名称") != null)
            {
                object tmpVar = null;
                IEdmPos5 pos = cfgList.GetHeadPosition();
                string cfgName = null;
                while (!pos.IsNull)
                {
                    cfgName = cfgList.GetNext(pos);
                    enumVar.GetVar("名称", cfgName, out tmpVar);
                    if (tmpVar == null) continue;
                    if (!tmpVar.ToString().Equals("")) break;
                }
                if (tmpVar != null) asmPrd.name = tmpVar.ToString();
            }
            if (varMgr.GetVariable("设计") != null)
            {
                object tmpVar = null;
                IEdmPos5 pos = cfgList.GetHeadPosition();
                string cfgName = null;
                while (!pos.IsNull)
                {
                    cfgName = cfgList.GetNext(pos);
                    enumVar.GetVar("设计", cfgName, out tmpVar);
                    if (tmpVar == null) continue;
                    if (!tmpVar.ToString().Equals("")) break;
                }
                if (tmpVar != null) asmPrd.bzr = tmpVar.ToString();
            }
            if (varMgr.GetVariable("设计日期") != null)
            {
                object tmpVar = null;
                IEdmPos5 pos = cfgList.GetHeadPosition();
                string cfgName = null;
                while (!pos.IsNull)
                {
                    cfgName = cfgList.GetNext(pos);
                    enumVar.GetVar("设计日期", cfgName, out tmpVar);
                    if (tmpVar == null) continue;
                    if (!tmpVar.ToString().Equals("")) break;
                }
                if (tmpVar!=null)
                    asmPrd.bzsj = tmpVar.ToString();
                if (asmPrd.bzsj.IndexOf(" ") >= 0) asmPrd.bzsj = asmPrd.bzsj.Substring(0, asmPrd.bzsj.IndexOf(" "));
            }
            if (varMgr.GetVariable("批准") != null)
            {
                object tmpVar = null;
                IEdmPos5 pos = cfgList.GetHeadPosition();
                string cfgName = null;
                while (!pos.IsNull)
                {
                    cfgName = cfgList.GetNext(pos);
                    enumVar.GetVar("批准", cfgName, out tmpVar);
                    if (tmpVar == null) continue;
                    if (!tmpVar.ToString().Equals("")) break;
                }
                if (tmpVar != null) asmPrd.pzr = tmpVar.ToString();
            }
            if (varMgr.GetVariable("批准日期") != null)
            {
                object tmpVar = null;
                IEdmPos5 pos = cfgList.GetHeadPosition();
                string cfgName = null;
                while (!pos.IsNull)
                {
                    cfgName = cfgList.GetNext(pos);
                    enumVar.GetVar("批准日期", cfgName, out tmpVar);
                    if (tmpVar == null) continue;
                    if (!tmpVar.ToString().Equals("")) break;
                }
                if (tmpVar != null) asmPrd.pzsj = tmpVar.ToString();
                if (asmPrd.pzsj.IndexOf(" ") >= 0) asmPrd.pzsj = asmPrd.pzsj.Substring(0, asmPrd.pzsj.IndexOf(" "));
            }
            if (varMgr.GetVariable("阶段标记") != null)
            {
                object tmpVar = null;
                IEdmPos5 pos = cfgList.GetHeadPosition();
                string cfgName = null;
                while (!pos.IsNull)
                {
                    cfgName = cfgList.GetNext(pos);
                    enumVar.GetVar("阶段标记", cfgName, out tmpVar);
                    if (tmpVar == null) continue;
                    if (!tmpVar.ToString().Equals("")) break;
                }
                if (tmpVar != null) asmPrd.jdbj = tmpVar.ToString();
            }

            if (varMgr.GetVariable("设备型号") != null)
            {
                object tmpVar = null;
                IEdmPos5 pos = cfgList.GetHeadPosition();
                string cfgName = null;
                while (!pos.IsNull)
                {
                    cfgName = cfgList.GetNext(pos);
                    enumVar.GetVar("设备型号", cfgName, out tmpVar);
                    if (tmpVar == null) continue;
                    if (!tmpVar.ToString().Equals("")) break;
                }
                if (tmpVar != null) asmPrd.sbxh = tmpVar.ToString();
            }
            if (varMgr.GetVariable("Project Name") != null)
            {
                object tmpVar = null;
                IEdmPos5 pos = cfgList.GetHeadPosition();
                string cfgName = null;
                while (!pos.IsNull)
                {
                    cfgName = cfgList.GetNext(pos);
                    enumVar.GetVar("Project Name", cfgName, out tmpVar);
                    if (tmpVar == null) continue;
                    if (!tmpVar.ToString().Equals("")) break;
                }
                if (tmpVar != null) asmPrd.projName = tmpVar.ToString();
            }
            return asmPrd;
        }
        public SldBsp GetSldPrdInfoFromFile(IEdmVault5 poVault, string filePath)
        {
            IEdmFolder5 opParentFolder;
            //bspPrt = new SldBuy();
            IEdmFile5 poFile = poVault.GetFileFromPath(filePath, out opParentFolder);
            if (poFile == null)
            {
                return null;
            }
            IEdmEnumeratorVariable5 enumVar = poFile.GetEnumeratorVariable();
            if (enumVar == null) return null;
            IEdmVariableMgr5 varMgr = (IEdmVariableMgr5)poVault;
            if (varMgr == null)
                return null;
            
            object tmpVar;
            string partType = "";
            if (varMgr.GetVariable("Part Type") != null && enumVar.GetVar("Part Type", "@", out tmpVar) && tmpVar != null)
            {
                partType = tmpVar.ToString();
            }
            if (partType == "自制件")
            {

                SldPrt prt = new SldPrt();
                prt.type = 1;
                object route = null;
                if (varMgr.GetVariable("工艺路线") != null)
                {
                    if (enumVar.GetVar("工艺路线", "@", out route) && route != null)
                    {
                        prt.route = route.ToString();
                    }
                }
                
                return prt;

            }
            else if (partType == "标准件")
            {
                SldStd std = new SldStd();
                std.type = 2;
                return std;
            }
            else if(partType=="外购件")
            {
                SldBuy buy = new SldBuy();
                buy.type = 3;
                return buy;
            }
            else
            {
                SldBsp buy = new SldBsp();
                buy.type = 4;
                return buy;
            }

        }

        public bool SavePrtToWorkbook(Workbook workbook, SldAsm asmPrd)
        {
            if (asmPrd.sldPrtList.Count <= 0 && asmPrd.sldStdList.Count <= 0 && asmPrd.sldBuyList.Count <= 0)
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
            if(insHead){//插组件
                SldPrt prt = new SldPrt();
                prt.name = asmPrd.name;
                prt.number = asmPrd.number;
                prt.amout = 1;
                asmPrd.sldPrtList.Insert(0, prt);
                count++;
            }
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
            if (insHead && asmPrd.sldBuyList.Count > 0)
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
                worksheet.Range["K23"].Value = asmPrd.name;
                if (asmPrd.jdbj.IndexOf("S")>=0)
                    worksheet.Range["K29"].Value = "S";
                else if (asmPrd.jdbj.IndexOf("A") >= 0)
                    worksheet.Range["M29"].Value = "A";
                else if (asmPrd.jdbj.IndexOf("B") >= 0)
                    worksheet.Range["O29"].Value = "B";
                else
                    worksheet.Range["R29"].Value = asmPrd.jdbj;
                worksheet.Range["K31"].Value = "共 " + pageSize + " 页";
                worksheet.Range["S31"].Value = "第 " + (pageIndex + 1) + " 页";
                worksheet.Range["W31"].Value = asmPrd.sbxh;
                int row = 0;
                while (iEnum.MoveNext())
                {

                    SldPrt prt = iEnum.Current;
                    int totalIndex = row + rowStartIndex;
                    worksheet.Range["B" + totalIndex].Value = prt.materialNumber;
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

        public bool SaveBspToWorkbook(Workbook workbook, SldAsm asmPrd)
        {
            if (asmPrd.sldBspList.Count <= 0)
                return true;
            int count = asmPrd.sldBspList.Count;
            //直接写，不用格式
            

            Sheets sheets = workbook.Worksheets;
            Worksheet worksheet = null;
           // if (sheets.get_Item("未确定") == null)
           // {
                worksheet = (Worksheet)sheets.Add();
                worksheet.Name = "错误";
            //}
            //else
            //{
           //     worksheet =  (Worksheet)sheets.get_Item("未确定");
           // }
            if (count <= 0)
            {
                worksheet.Delete();
                return false;
            }
            List<SldBsp>.Enumerator iEnum = asmPrd.sldBspList.GetEnumerator();
            int colNum = 9;
            object[,] dataArray = new object[count + 1, colNum];
            
            {
                dataArray[0, 0] = "物料代码";
                dataArray[0, 1] = "图号";
                dataArray[0, 2] = "名称";
                dataArray[0, 3] = "数量";
                dataArray[0, 4] = "材料";
                dataArray[0, 5] = "重量";
                dataArray[0, 6] = "总重";
                dataArray[0, 7] = "备注";
                dataArray[0, 8] = "类型";
            }
            int i = 0;
            while (iEnum.MoveNext())
            {
                SldBsp bsp = iEnum.Current;
                dataArray[i + 1, 0] = bsp.materialNumber;
                dataArray[i + 1, 1] = bsp.number;
                dataArray[i + 1, 2] = bsp.name;
                dataArray[i + 1, 3] = bsp.amout == 0 ? "" : bsp.amout.ToString();
                dataArray[i + 1, 4] = bsp.material;
                dataArray[i + 1, 5] = bsp.weight == 0 ? "" : bsp.weight.ToString();
                dataArray[i + 1, 6] = bsp.totalWeight == 0 ? "" : bsp.totalWeight.ToString();
                dataArray[i + 1, 7] = bsp.remark;
                dataArray[i + 1, 8] = bsp.type == 4 ? "未定义类型" : "未找到关联文件";
                i++;
            }
            worksheet.Range["A1", worksheet.Cells[count+1, colNum]] = dataArray;
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
                worksheet.Range["K36"].Value = asmPrd.name;
                if (asmPrd.jdbj.IndexOf("S") >= 0)
                    worksheet.Range["K42"].Value = "S";
                else if (asmPrd.jdbj.IndexOf("A") >= 0)
                    worksheet.Range["L42"].Value = "A";
                else if (asmPrd.jdbj.IndexOf("B") >= 0)
                    worksheet.Range["M42"].Value ="B";
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
                worksheet.Range["K36"].Value = asmPrd.name;
                if (asmPrd.jdbj.IndexOf("S") >= 0)
                    worksheet.Range["K42"].Value = "S";
                else if (asmPrd.jdbj.IndexOf("A") >= 0)
                    worksheet.Range["L42"].Value = "A";
                else if (asmPrd.jdbj.IndexOf("B") >= 0)
                    worksheet.Range["M42"].Value = "B";
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
