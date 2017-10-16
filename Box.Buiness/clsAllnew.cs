using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using BOX.DB;
using BOX_InvoiceERP;
using DCTS.CustomComponents;

namespace Box.Buiness
{
    public class clsAllnew
    {
        private BackgroundWorker bgWorker1;
        private SortableBindingList<clsDatabaseinfo> sortableCaseListResult;
        public List<clsDatabaseinfo> Result;

        public log4net.ILog ProcessLogger { get; set; }
        public log4net.ILog ExceptionLogger { get; set; }
        public clsAllnew()
        {

            InitialSystemInfo();

        }
        private void InitialSystemInfo()
        {
            #region 初始化配置
            ProcessLogger = log4net.LogManager.GetLogger("ProcessLogger");
            ExceptionLogger = log4net.LogManager.GetLogger("SystemExceptionLogger");
            ProcessLogger.Fatal("System Start " + DateTime.Now.ToString());
            #endregion
        }



        public List<clsDatabaseinfo> InputclaimReport(ref BackgroundWorker bgWorker, string path)
        {
            try
            {
                bgWorker1 = bgWorker;
                Result = new List<clsDatabaseinfo>();
                ProcessLogger.Fatal("1006-- Input C Data start" + DateTime.Now.ToString());

                Result = ReadMAPFile(path);
                ProcessLogger.Fatal("20016-- Input C Data end" + DateTime.Now.ToString());
                #region 逻辑层

                List<clsDatabaseinfo> nonDuplicateList2 = Result.Where((x, i) => Result.FindIndex(z => z.zhongwenmingcheng == x.zhongwenmingcheng && z.jianyiguilei == x.jianyiguilei && z.yuanchanguo == x.yuanchanguo) == i).ToList();
                List<clsDatabaseinfo> nonDuplicateList = Result.Except(nonDuplicateList2).ToList();
                sortableCaseListResult = new SortableBindingList<clsDatabaseinfo>(Result);
                var gos = sortableCaseListResult.GroupBy(x => x.zhongwenmingcheng).ToList();

                //   dailyReportResults.Remove(dailyReportResults.Where(o => o.Id == filtered[i].Id).Single());
                List<clsDatabaseinfo> NewResult = new List<clsDatabaseinfo>();
                foreach (clsDatabaseinfo item in nonDuplicateList2)
                {

                    if (item.zhongwenmingcheng == "不锈钢制弹簧螺母")
                    {

                    }
                    List<clsDatabaseinfo> stockstate = Result.FindAll(o => o.zhongwenmingcheng == item.zhongwenmingcheng && o.jianyiguilei == item.jianyiguilei && o.yuanchanguo == item.yuanchanguo);

                    clsDatabaseinfo newitem = new clsDatabaseinfo();
                    newitem._id = item._id;
                    newitem.Message = item.Message;
                    newitem.Xuhao = item.Xuhao;
                    newitem.lingjianhao = item.lingjianhao;
                    newitem.yingwenmingcheng = item.yingwenmingcheng;
                    newitem.zhongwenmingcheng = item.zhongwenmingcheng;
                    newitem.jianyiguilei = item.jianyiguilei;
                    newitem.jianguantiaojian = item.jianguantiaojian;
                    newitem.yuanchanguo = item.yuanchanguo;



                    newitem.danwei = item.danwei;
                    newitem.shenbaoyaoshu = item.shenbaoyaoshu;
                    newitem.Input_Date = item.Input_Date;

                    var nullableQty = (from s in stockstate
                                       where s.lingjianshuliang != null && s.lingjianshuliang != ""
                                       select Convert.ToInt32(s.lingjianshuliang)).Sum();

                    var ZJZ = (from s in stockstate
                               where s.zongjinzhong != null && s.zongjinzhong != ""
                               select Convert.ToDouble(s.zongjinzhong)).Sum();


                    var ljzj = (from s in stockstate
                                where s.lingjianzongjia != null && s.lingjianzongjia != ""
                                select Convert.ToDouble(s.lingjianzongjia)).Sum();

                    newitem.lingjianshuliang = Math.Abs(nullableQty).ToString();
                    newitem.zongjinzhong = Math.Abs(ZJZ).ToString();
                    newitem.lingjianzongjia = Math.Abs(ljzj).ToString();

                    if (stockstate.Count > 1)
                        newitem.Message = "多条";
                    else if (stockstate.Count == 1)
                        newitem.Message = "单条";

                    NewResult.Add(newitem);

                }

                #region MyRegion
                //var groupedStockList = sortableCaseListResult.GroupBy(s => new { s.zhongwenmingcheng, s.jianyiguilei, s.yuanchanguo });

                //var sums = Result.GroupBy(s => new { s.zhongwenmingcheng, s.jianyiguilei, s.yuanchanguo });
                //foreach (var igrouping in sums)
                //{
                //    int dou = igrouping.Count();

                //    var item = igrouping.First();
                //    string sss = item.zhongwenmingcheng;
                //    if (sss == "不锈钢制弹簧螺母")
                //    {

                //    }

                //}



                //nonDuplicateList = new List<clsDatabaseinfo>();

                //nonDuplicateList = (from data in sums
                //                    where data.First().Xuhao != ""
                //                    select new clsDatabaseinfo
                //                        {
                //                            _id = data.First()._id,
                //                            Message = data.First().Message,
                //                            Xuhao = data.First().Xuhao,
                //                            lingjianhao = data.First().lingjianhao,
                //                            yingwenmingcheng = data.First().yingwenmingcheng,
                //                            zhongwenmingcheng = data.First().zhongwenmingcheng,
                //                            jianyiguilei = data.First().jianyiguilei,
                //                            jianguantiaojian = data.First().jianguantiaojian,
                //                            yuanchanguo = data.First().yuanchanguo,
                //                            lingjianshuliang = data.First().lingjianshuliang,
                //                            zongjinzhong = data.First().zongjinzhong,
                //                            lingjianzongjia = data.First().lingjianzongjia,
                //                            danwei = data.First().danwei,
                //                            shenbaoyaoshu = data.First().shenbaoyaoshu,
                //                            Input_Date = data.First().Input_Date
                //                        }).ToList();


                //int dd = sums.Count();
                ////var categories = from p in Result
                ////                 group p by new
                ////                 {
                ////                     p.zhongwenmingcheng,
                ////                     p.jianyiguilei,
                ////                     p.yuanchanguo
                ////                 };



                //var categories = from p in nonDuplicateList2
                //                 group p by new
                //                 {
                //                     p.zhongwenmingcheng,
                //                     p.jianyiguilei,
                //                     p.yuanchanguo
                //                 }
                //                     into g
                //                     select new
                //                     {
                //                         g.Key,
                //                         g
                //                     };

                ////List<clsDatabaseinfo> hebinghoude = (from clsDatabaseinfo data in categories
                ////                                     where data.zhongwenmingcheng != null

                ////                                     select new clsDatabaseinfo
                ////                                     {
                ////                                         _id = data._id,
                ////                                         Message = data.Message,
                ////                                         Xuhao = data.Xuhao,
                ////                                         lingjianhao = data.lingjianhao,
                ////                                         yingwenmingcheng = data.yingwenmingcheng,
                ////                                         zhongwenmingcheng = data.zhongwenmingcheng,
                ////                                         jianyiguilei = data.jianyiguilei,
                ////                                         jianguantiaojian = data.jianguantiaojian,
                ////                                         yuanchanguo = data.yuanchanguo,
                ////                                         lingjianshuliang = data.lingjianshuliang,
                ////                                         zongjinzhong = data.zongjinzhong,
                ////                                         lingjianzongjia = data.lingjianzongjia,
                ////                                         danwei = data.danwei,
                ////                                         shenbaoyaoshu = data.shenbaoyaoshu,
                ////                                         Input_Date = data.Input_Date
                ////                                     }
                ////                                ).ToList();



                #endregion
                ;

                #endregion
                //List<clsDatabaseinfo> list3 = hebinghoude.Concat(nonDuplicateList).ToList();
                bgWorker1.ReportProgress(0, "导出数据中,请等待.... " );
              
                DownExcel(NewResult);
                return NewResult;

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error:" + ex.Message);
                return null;

            }
            return null;
        }
        public List<clsDatabaseinfo> ReadMAPFile(string path)
        {

            List<clsDatabaseinfo> Result = new List<clsDatabaseinfo>();

            // string path = AppDomain.CurrentDomain.BaseDirectory + "Resources\\ALL MU.xls";
            System.Globalization.CultureInfo CurrentCI = System.Threading.Thread.CurrentThread.CurrentCulture;
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");

            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook analyWK = excelApp.Workbooks.Open(path, Type.Missing, true, Type.Missing,
                "htc", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            ProcessLogger.Fatal("2001-- OPEN 报关明细 " + DateTime.Now.ToString());

            Microsoft.Office.Interop.Excel.Worksheet WS = (Microsoft.Office.Interop.Excel.Worksheet)analyWK.Worksheets["报关明细"];
            Microsoft.Office.Interop.Excel.Range rng;
            rng = WS.get_Range(WS.Cells[2, 1], WS.Cells[WS.UsedRange.Rows.Count, 30]);
            int rowCount = WS.UsedRange.Rows.Count - 1;
            object[,] o = new object[1, 1];
            o = (object[,])rng.Value2;
            clsCommHelp.CloseExcel(excelApp, analyWK);

            for (int i = 1; i <= rowCount; i++)
            {
                bgWorker1.ReportProgress(0, "导入数据中  :  " + i.ToString() + "/" + rowCount.ToString());
                clsDatabaseinfo temp = new clsDatabaseinfo();

                #region 基本几何

                temp.Xuhao = "";
                if (o[i, 1] != null)
                    temp.Xuhao = o[i, 1].ToString().Trim();
                if (temp.Xuhao == "" || temp.Xuhao == null)
                    continue;

                temp.lingjianhao = "";
                if (o[i, 2] != null)
                    temp.lingjianhao = o[i, 2].ToString().Trim();


                temp.yingwenmingcheng = "";
                if (o[i, 3] != null)
                    temp.yingwenmingcheng = o[i, 3].ToString().Trim();


                temp.zhongwenmingcheng = "";
                if (o[i, 4] != null)
                    temp.zhongwenmingcheng = o[i, 4].ToString().Trim();


                temp.jianyiguilei = "";
                if (o[i, 5] != null)
                    temp.jianyiguilei = o[i, 5].ToString().Trim();

                temp.jianguantiaojian = "";
                if (o[i, 6] != null)
                    temp.jianguantiaojian = o[i, 6].ToString().Trim();

                temp.yuanchanguo = "";
                if (o[i, 7] != null)
                    temp.yuanchanguo = o[i, 7].ToString().Trim();
                temp.lingjianshuliang = "";
                if (o[i, 8] != null)
                    temp.lingjianshuliang = o[i, 8].ToString().Trim();

                temp.zongjinzhong = "";
                if (o[i, 11] != null)
                    temp.zongjinzhong = o[i, 11].ToString().Trim();

                temp.lingjianzongjia = "";
                if (o[i, 12] != null)
                    temp.lingjianzongjia = o[i, 12].ToString().Trim();

                temp.danwei = "";
                if (o[i, 13] != null)
                    temp.danwei = o[i, 13].ToString().Trim();

                temp.shenbaoyaoshu = "";
                if (o[i, 15] != null)
                    temp.shenbaoyaoshu = o[i, 15].ToString().Trim();
                temp.Input_Date = DateTime.Now.ToString("yyyyMMdd-HHmm");
                #endregion
                Result.Add(temp);
            }
            return Result;

        }
        private void DownExcel(List<clsDatabaseinfo> Results)
        {
            #region 获取模板路径
            System.Globalization.CultureInfo CurrentCI = System.Threading.Thread.CurrentThread.CurrentCulture;
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
            string fullPath = Path.Combine(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "System\\"), "Tel.xlsx");
            SaveFileDialog sfdDownFile = new SaveFileDialog();
            sfdDownFile.OverwritePrompt = false;
            string DesktopPath = Convert.ToString(System.Environment.GetFolderPath(Environment.SpecialFolder.Desktop));
            sfdDownFile.Filter = "Excel files (*.xls,*.xlsx)|*.xls;*.xlsx";
            string file = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Results\\");
            file = Convert.ToString(System.Environment.GetFolderPath(Environment.SpecialFolder.Desktop)) + "\\";
            sfdDownFile.FileName = Path.Combine(file, "System Output" + DateTime.Now.ToString("yyyyMMdd-ss"));
            string strExcelFileName = string.Empty;

            #endregion

            #region 导出前校验模板信息
            if (string.IsNullOrEmpty(sfdDownFile.FileName))
            {
                MessageBox.Show("File name can't be empty, please confirm, thanks!");
                return;
            }
            if (!File.Exists(fullPath))
            {
                MessageBox.Show("Template file does not exist, please confirm, thanks!");
                return;
            }
            else
            {
                strExcelFileName = sfdDownFile.FileName + ".xlsx";
            }
            #endregion
            #region Excel 初始化

            Microsoft.Office.Interop.Excel.Application ExcelApp;
            ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            System.Reflection.Missing missingValue = System.Reflection.Missing.Value;
            //
            ProcessLogger.Fatal("2301-- OPEN Tel " + DateTime.Now.ToString());

            Microsoft.Office.Interop.Excel._Workbook ExcelBook =
            ExcelApp.Workbooks.Open(fullPath, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue);
            #endregion
            #region Sheet 初始化
            try
            {
                {
                    Microsoft.Office.Interop.Excel._Worksheet ExcelSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelBook.Worksheets[1];

                    //打开时是否显示Excel
                    //ExcelApp.Visible = true;
                    //ExcelApp.ScreenUpdating = true;
                    int dou = ExcelSheet.UsedRange.Rows.Count + 1;
                    string las = "BT" + dou.ToString();
                    bool issave = false;
                    Microsoft.Office.Interop.Excel.Range rng = ExcelSheet.get_Range("A2", las);
                    rng.Delete();


            #endregion

                    #region 填充数据
                    int RowIndex = 1;
                    int xuhao = 1;


                    foreach (clsDatabaseinfo item in Results)
                    {


                        RowIndex++;
                        #region MyRegion

                        ExcelSheet.Cells[RowIndex, 1] = item.Xuhao;
                        ExcelSheet.Cells[RowIndex, 2] = item.lingjianhao;
                        ExcelSheet.Cells[RowIndex, 3] = item.zhongwenmingcheng;
                        ExcelSheet.Cells[RowIndex, 4] = item.jianyiguilei;
                        ExcelSheet.Cells[RowIndex, 5] = item.jianguantiaojian;
                        ExcelSheet.Cells[RowIndex, 6] = item.yuanchanguo;
                        ExcelSheet.Cells[RowIndex, 7] = item.lingjianshuliang;
                        ExcelSheet.Cells[RowIndex, 8] = item.zongjinzhong;
                        ExcelSheet.Cells[RowIndex, 9] = item.lingjianzongjia;
                        ExcelSheet.Cells[RowIndex, 10] = item.shenbaoyaoshu;

                        #endregion


                    }


                    ExcelBook.RefreshAll();
                    #region 写入文件
                    ProcessLogger.Fatal("2801-- Save  " + DateTime.Now.ToString());
                    ExcelApp.ScreenUpdating = true;
                    ExcelBook.SaveAs(strExcelFileName, missingValue, missingValue, missingValue, missingValue, missingValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, missingValue, missingValue, missingValue, missingValue, missingValue);
                    ExcelApp.DisplayAlerts = false;
                    issave = false;
                  

                    #endregion
                }
            }

            #region 异常处理
            catch (Exception ex)
            {
                ExcelApp.DisplayAlerts = false;
                ExcelApp.Quit();
                ExcelBook = null;
                ExcelApp = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
                throw ex;
            }
            #endregion

            #region Finally垃圾回收
            finally
            {
                ExcelBook.Close(false, missingValue, missingValue);
                ExcelBook = null;
                ExcelApp.DisplayAlerts = true;
                ExcelApp.Quit();
                clsKeyMyExcelProcess.Kill(ExcelApp);

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            #endregion

                    #endregion
        }

    }
}
