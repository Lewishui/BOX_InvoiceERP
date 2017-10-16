using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace BOX_InvoiceERP
{
    public partial class Main : Form
    {
        frmAboutBox aboutbox;
        Sunisoft.IrisSkin.SkinEngine se = null;
        public log4net.ILog ProcessLogger;
        public log4net.ILog ExceptionLogger;
        private frmBOX frmBOX;

        public Main()
        {
            InitializeComponent();
            aboutbox = new frmAboutBox();
            se = new Sunisoft.IrisSkin.SkinEngine();
            se.SkinAllForm = true;
            se.SkinFile = Path.Combine(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ""), "GlassBrown.ssk");
            #region Noway
            //DateTime oldDate = DateTime.Now;
            //DateTime dt3;
            //string endday = DateTime.Now.ToString("yyyy/MM/dd");
            //dt3 = Convert.ToDateTime(endday);
            //DateTime dt2;
            //dt2 = Convert.ToDateTime("2017/10/18");

            //TimeSpan ts = dt2 - dt3;
            //int timeTotal = ts.Days;
            //if (timeTotal < 0)
            //{
            //    MessageBox.Show("Please Contact your administrator !");
            //    this.Close();
            //}
            #endregion

        }

        private void 关于系统ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            aboutbox.ShowDialog();
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {


            if (frmBOX == null)
            {
                frmBOX = new frmBOX();
                frmBOX.FormClosed += new FormClosedEventHandler(FrmOMS_FormClosed);
            }
            if (frmBOX == null)
            {
                frmBOX = new frmBOX();
            }
            frmBOX.Show(this.dockPanel2);


        }
        void FrmOMS_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (sender is frmBOX)
            {
                frmBOX = null;
            }
        }
    }
}
