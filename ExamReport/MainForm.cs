using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using Telerik.WinControls.Data;
using Telerik.WinControls.UI;
namespace ExamReport
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
            Binder_datasource();
        }

        void Binder_datasource()
        {
            //string addr = System.IO.Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath);
            //string conn = @"Provider=vfpoledb;Data Source=" + addr + ";Collating Sequence=machine;";

            //OleDbConnection dbfConnection = new OleDbConnection(conn);

            //OleDbDataAdapter adpt = new OleDbDataAdapter("select * from " + "schoolcode", dbfConnection);
            ////OleDbDataAdapter adpt = new OleDbDataAdapter("select * from " + file + " where Qk<>1", dbfConnection);
            //DataSet mySet = new DataSet();

            //try
            //{
            //    adpt.Fill(mySet);
            //}
            //catch (OleDbException e)
            //{
            //    throw new Exception("数据库文件被占用，请关闭！");
            //}
            //dbfConnection.Close();

            //DatabaseGridView.BeginUpdate();
            //DataTable dt = MySqlHelper.GetDataSet(MySqlHelper.Conn, CommandType.Text, "select * from exam_meta_data", null).Tables[0];
            //DatabaseGridView.DataSource = dt;
            


            ////DatabaseGridView.MasterTemplate.AutoExpandGroups = true;
            ////DatabaseGridView.MasterTemplate.EnableFiltering = true;
            ////DatabaseGridView.ShowGroupPanel = true;
            ////DatabaseGridView.EnableHotTracking = true;

            //this.DatabaseGridView.TableElement.EndUpdate(false);


            ////DatabaseGridView.TableElement.CellSpacing = -1;
            //DatabaseGridView.TableElement.TableHeaderHeight = 35;
            //DatabaseGridView.TableElement.GroupHeaderHeight = 30;
            //DatabaseGridView.TableElement.RowHeight = 25;

            //DatabaseGridView.GroupDescriptors.Clear();
            //DatabaseGridView.GroupDescriptors.Add(new GridGroupByExpression("exam as exam format \"{0}: {1}\" Group By exam"));
        }

        private void button1_Click(object sender, EventArgs e)
        {
            MyWizard wizard = new MyWizard();
            wizard.Show();
        }

        private void zk_zt_button_Click(object sender, EventArgs e)
        {
            zk_zt_panel.Visible = true;
            zk_qx_panel.Visible = false;


        }

        private void zk_qx_button_Click(object sender, EventArgs e)
        {
            zk_qx_panel.Visible = true;
            zk_zt_panel.Visible = false;
        }

        private void radPageViewPage1_Paint(object sender, PaintEventArgs e)
        {

        }


        

      

        

    }
}
