using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Twin = Telerik.WinControls;
using MyParams = MySql.Data.MySqlClient.MySqlParameter;

namespace ExamReport
{
    public partial class mainform : Telerik.WinControls.UI.RadForm
    {
        public mainform()
        {
            InitializeComponent();
            TotalGrid_Load();
        }

        
        private void zk_zt_button_Click_1(object sender, EventArgs e)
        {
            zk_zt_panel.Show();
            zk_qx_panel.Hide();
        }

        private void zk_qx_button_Click(object sender, EventArgs e)
        {
            zk_qx_panel.Show();
            zk_zt_panel.Hide();
        }

        private void button11_Click(object sender, EventArgs e)
        {

        }

        public void TotalGrid_Load()
        {
            TotalGridView.MasterTemplate.AllowAddNewRow = false;
            TotalGridView.TableElement.BeginUpdate();
            DataTable data = MySqlHelper.GetDataSet(MySqlHelper.Conn, CommandType.Text, "select * from exam_meta_data", null).Tables[0];
            TotalGridView.DataSource = data.LanguageTrans();

            TotalGridView.TableElement.EndUpdate();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            MyWizard wizard = new MyWizard();
            wizard.Show();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string id = TotalGridView.CurrentRow.Cells[0].Value.ToString().Trim();
            //MyParams param = new MyParams("@id",  MySql.Data.MySqlClient.MySqlDbType.VarChar, 5);
            //param.Value = Convert.ToInt32(TotalGridView.CurrentRow.Cells[0].Value.ToString().Trim());
            //param.Value = "hk";

            int val = MySqlHelper.ExecuteNonQuery(MySqlHelper.Conn, CommandType.Text, "delete from exam_meta_data where id = " + id, null);
            //int val2 = MySqlHelper.ExecuteNonQuery(MySqlHelper.Conn, CommandType.Text, "insert into exam_meta_data (year,exam,sub,ans,grp,fullmark,zh) values ('2014', 'hk','yy','1','1',150,'0')", null);
            TotalGrid_Load();
        }

       

    }
}
