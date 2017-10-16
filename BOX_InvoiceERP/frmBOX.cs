using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Box.Buiness;
using BOX.DB;
using DCTS.CustomComponents;
using WeifenLuo.WinFormsUI.Docking;

namespace BOX_InvoiceERP
{
    public partial class frmBOX : DockContent
    {
        List<clsDatabaseinfo> Result = new List<clsDatabaseinfo>();
        // 后台执行控件
        private BackgroundWorker bgWorker;
        // 消息显示窗体
        private frmMessageShow frmMessageShow;
        // 后台操作是否正常完成
        private bool blnBackGroundWorkIsOK = false;
        //后加的后台属性显
        private bool backGroundRunResult;
        public log4net.ILog ProcessLogger { get; set; }
        public log4net.ILog ExceptionLogger { get; set; }
        string MCpath;
        private SortableBindingList<clsDatabaseinfo> sortableCaseListResult;

        List<clsDatabaseinfo> ALLResult;
        
        public frmBOX()
        {
            InitializeComponent();
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



        private void InitialBackGroundWorker()
        {
            bgWorker = new BackgroundWorker();
            bgWorker.WorkerReportsProgress = true;
            bgWorker.WorkerSupportsCancellation = true;
            bgWorker.RunWorkerCompleted +=
                new RunWorkerCompletedEventHandler(bgWorker_RunWorkerCompleted);
            bgWorker.ProgressChanged +=
                new ProgressChangedEventHandler(bgWorker_ProgressChanged);
        }

        private void bgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                blnBackGroundWorkIsOK = false;
            }
            else if (e.Cancelled)
            {
                blnBackGroundWorkIsOK = true;
            }
            else
            {
                blnBackGroundWorkIsOK = true;
            }
        }

        private void bgWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (frmMessageShow != null && frmMessageShow.Visible == true)
            {
                //设置显示的消息
                frmMessageShow.setMessage(e.UserState.ToString());
                //设置显示的按钮文字
                if (e.ProgressPercentage == clsConstant.Thread_Progress_OK)
                {
                    frmMessageShow.setStatus(clsConstant.Dialog_Status_Enable);
                }
            }
        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            MCpath = "";
            OpenFileDialog tbox = new OpenFileDialog();
            tbox.Multiselect = false;
            //  tbox.Filter = "Excel Files(*.xls,*.xlsx,*.xlsm,*.xlsb)|*.xls;*.xlsx;*.xlsm;*.xlsb";
            tbox.Filter = "Excel Files(*.xls,*.xlsx,*.xlsm,*.xlsb,*.txt)|*.xls;*.xlsx;*.xlsm;*.xlsb;*.txt";
            if (tbox.ShowDialog() == DialogResult.OK)
            {
                MCpath = tbox.FileName;
            }
            if (MCpath == null || MCpath == "")
                return;
            this.toolStripLabel2.Text = MCpath.Trim();

            if (MessageBox.Show(" 将导入新数据导入系统，是否继续 ?", "Info", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
            }
            else
                return;
            try
            {
                InitialBackGroundWorker();
                bgWorker.DoWork += new DoWorkEventHandler(inputdatacaipiao);

                bgWorker.RunWorkerAsync();
                // 启动消息显示画面
                frmMessageShow = new frmMessageShow(clsShowMessage.MSG_001,
                                                    clsShowMessage.MSG_007,
                                                    clsConstant.Dialog_Status_Disable);
                frmMessageShow.ShowDialog();
                // 数据读取成功后在画面显示
                if (blnBackGroundWorkIsOK)
                {
                    // this.dataGridView.DataSource = null;
                    this.dataGridView1.AutoGenerateColumns = false;
                    sortableCaseListResult = new SortableBindingList<clsDatabaseinfo>(Result);
                    this.dataGridView1.DataSource = sortableCaseListResult;
                    //显示信息
                    if (ALLResult != null)
                        label2.Text = ALLResult.Count.ToString();
                    if (Result != null)
                        label4.Text = Result.Count.ToString();
                    if (ALLResult != null && Result != null)
                        label6.Text = (ALLResult.Count - Result.Count).ToString();
                    this.toolStripLabel2.Text = "系统运行结束，请查看桌面已生成的文件";

                }
            }
            catch (Exception ex)
            {
                throw ex;
            }



        }
        private void inputdatacaipiao(object sender, DoWorkEventArgs e)
        {
            ALLResult = new List<clsDatabaseinfo>();

            //导入程序集
            DateTime oldDate = DateTime.Now;
            Result = new List<clsDatabaseinfo>();
            clsAllnew BusinessHelp = new clsAllnew();
            ProcessLogger.Fatal("1005--input kiajiang data" + DateTime.Now.ToString());
            Result = BusinessHelp.InputclaimReport(ref this.bgWorker, MCpath);
            ALLResult = BusinessHelp.Result;
            ProcessLogger.Fatal("1006-- Input finish" + DateTime.Now.ToString());
            DateTime FinishTime = DateTime.Now;  //   
            TimeSpan s = DateTime.Now - oldDate;
            string timei = s.Minutes.ToString() + ":" + s.Seconds.ToString();
            string Showtime = clsShowMessage.MSG_029 + timei.ToString();
            bgWorker.ReportProgress(clsConstant.Thread_Progress_OK, clsShowMessage.MSG_009 + "\r\n" + Showtime);
        }

        private void toolStripButton6_Click(object sender, EventArgs e)
        {
            downEXCEL(dataGridView1);


        }
        private void downEXCEL(DataGridView dgv)
        {
            if (dgv.Rows.Count == 0)
            {
                MessageBox.Show("Sorry , No Data Output !", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            var saveFileDialog = new SaveFileDialog();
            saveFileDialog.DefaultExt = ".csv";
            saveFileDialog.Filter = "csv|*.csv";
            string strFileName = "System  Info" + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
            saveFileDialog.FileName = strFileName;
            if (saveFileDialog.ShowDialog(this) == DialogResult.OK)
            {
                strFileName = saveFileDialog.FileName.ToString();
            }
            else
            {
                return;
            }
            FileStream fa = new FileStream(strFileName, FileMode.Create);
            StreamWriter sw = new StreamWriter(fa, Encoding.Unicode);
            string delimiter = "\t";
            string strHeader = "";
            for (int i = 0; i < dgv.Columns.Count; i++)
            {
                strHeader += dgv.Columns[i].HeaderText + delimiter;
            }
            sw.WriteLine(strHeader);

            //output rows data
            for (int j = 0; j < dgv.Rows.Count; j++)
            {
                string strRowValue = "";

                for (int k = 0; k < dgv.Columns.Count; k++)
                {
                    if (dgv.Rows[j].Cells[k].Value != null)
                    {
                        strRowValue += dgv.Rows[j].Cells[k].Value.ToString().Replace("\r\n", " ").Replace("\n", "") + delimiter;
                        if (dgv.Rows[j].Cells[k].Value.ToString() == "LIP201507-35")
                        {

                        }

                    }
                    else
                    {
                        strRowValue += dgv.Rows[j].Cells[k].Value + delimiter;
                    }
                }
                sw.WriteLine(strRowValue);
            }
            sw.Close();
            fa.Close();
            MessageBox.Show("Dear User, Down File  Successful ！", "System", MessageBoxButtons.OK, MessageBoxIcon.Information);


        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            Result = new List<clsDatabaseinfo>();
            ALLResult = new List<clsDatabaseinfo>();

            this.dataGridView1.AutoGenerateColumns = false;
            this.dataGridView1.DataSource = null;
            this.dataGridView1.DataSource = Result;
            //显示信息

            label2.Text = "0";

            label4.Text = "0";

            label6.Text = "0";
        }

        private void toolStripButton5_Click(object sender, EventArgs e)
        {

            this.dataGridView1.DataSource = null;
            if (Result != null)
            {

                this.dataGridView1.AutoGenerateColumns = false;
                sortableCaseListResult = new SortableBindingList<clsDatabaseinfo>(Result);
                this.dataGridView1.DataSource = sortableCaseListResult;
                //显示信息
                if (ALLResult != null)
                    label2.Text = ALLResult.Count.ToString();
                if (Result != null)
                    label4.Text = Result.Count.ToString();
                if (ALLResult != null && Result != null)
                    label6.Text = (ALLResult.Count - Result.Count).ToString();
                this.toolStripLabel2.Text = "系统运行结束，请查看桌面已生成的文件";
            }
        }


    }
}
