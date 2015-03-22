using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Twin = Telerik.WinControls;
using MyParams = MySql.Data.MySqlClient.MySqlParameter;
using Telerik.WinControls.UI;
using System.Resources;
using System.Data.OleDb;
using System.Linq;
using System.Threading;

namespace ExamReport
{
    public delegate void ProgressDelegate(string key, int status, string text);

    public delegate void MyErrorMessage(string key, string Message);
    public partial class mainform : Telerik.WinControls.UI.RadForm
    {
        Dictionary<string, Thread> thread_store = new Dictionary<string, Thread>();

        Dictionary<string, RadLabel> progress_label = new Dictionary<string, RadLabel>();
        Dictionary<string, RadButton> run_button = new Dictionary<string, RadButton>();
        Dictionary<string, RadButton> cancel_button = new Dictionary<string, RadButton>();
        Dictionary<string, RadWaitingBar> waiting_bar = new Dictionary<string, RadWaitingBar>();

        string currentdic = System.IO.Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath);
        //Thread thread;
        public mainform()
        {
            InitializeComponent();
            Grid_load();
            TreeLoadData();
            ZKTreeView.SelectedNodeChanged += ZKTreeNode_Selected;
            GKTreeView.SelectedNodeChanged += GKTreeNode_Selected;
            init_dictionary();
            save_address.Text = currentdic;
            gk_save_address.Text = currentdic;
        }
        void init_dictionary()
        {
            progress_label.Add("zk_zt", zk_zt_progress);
            progress_label.Add("zk_qx", zk_qx_ProgressLabel);
            progress_label.Add("gk_zt", gk_zt_progresslabel);

            run_button.Add("zk_zt", zk_zt_start);
            run_button.Add("zk_qx", zk_qx_start);
            run_button.Add("gk_zt", gk_zt_start);

            cancel_button.Add("zk_zt", zk_zt_cancel);
            cancel_button.Add("zk_qx", zk_qx_cancel);
            cancel_button.Add("gk_zt", gk_zt_cancel);

            waiting_bar.Add("zk_zt", zk_zt_waitingbar);
            waiting_bar.Add("zk_qx", zk_qx_WaitingBar);
            waiting_bar.Add("gk_zt", gk_zt_waitingbar);

            
        }
        public DataTable schoolcode_table;
        
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
        public void Grid_load()
        {
            DataTable data = MySqlHelper.GetDataSet(MySqlHelper.Conn, CommandType.Text, "select * from exam_meta_data", null).Tables[0];
            zk_gridview.MasterTemplate.AllowAddNewRow = false;
            zk_gridview.TableElement.BeginUpdate();
            

            DataTable zk_data = data.AsEnumerable().AsEnumerable().Where(c => c.Field<string>("exam").Equals("zk")).Select(c => new
            {
                year = c.Field<string>("year"),
                sub = c.Field<string>("sub"),
                ans = c.Field<string>("ans"),
                grp = c.Field<string>("grp"),
                fullmark = c.Field<int>("fullmark"),
                gtype = c.Field<string>("gtype"),
                gnum = c.Field<int>("gnum")
            }).ToDataTable();

            zk_gridview.DataSource = zk_data.LanguageTrans();
            zk_gridview.TableElement.EndUpdate();

            gk_gridview.MasterTemplate.AllowAddNewRow = false;
            gk_gridview.TableElement.BeginUpdate();
            DataTable gk_data = data.AsEnumerable().AsEnumerable().Where(c => c.Field<string>("exam").Equals("gk")).Select(c => new
            {
                year = c.Field<string>("year"),
                sub = c.Field<string>("sub"),
                ans = c.Field<string>("ans"),
                grp = c.Field<string>("grp"),
                fullmark = c.Field<int>("fullmark"),
                gtype = c.Field<string>("gtype"),
                gnum = c.Field<int>("gnum")
            }).ToDataTable();
            gk_gridview.DataSource = gk_data.LanguageTrans();
            gk_gridview.TableElement.EndUpdate();
        }
        

        private void button2_Click(object sender, EventArgs e)
        {
            //MyWizard wizard = new MyWizard();
            //wizard.Show();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string id = zk_gridview.CurrentRow.Cells[0].Value.ToString().Trim();
            ////MyParams param = new MyParams("@id",  MySql.Data.MySqlClient.MySqlDbType.VarChar, 5);
            ////param.Value = Convert.ToInt32(TotalGridView.CurrentRow.Cells[0].Value.ToString().Trim());
            ////param.Value = "hk";

            //int val = MySqlHelper.ExecuteNonQuery(MySqlHelper.Conn, CommandType.Text, "delete from exam_meta_data where id = " + id, null);
            ////int val2 = MySqlHelper.ExecuteNonQuery(MySqlHelper.Conn, CommandType.Text, "insert into exam_meta_data (year,exam,sub,ans,grp,fullmark,zh) values ('2014', 'hk','yy','1','1',150,'0')", null);
            //TotalGrid_Load();
        }

        private void TreeLoadData()
        {
            //ZKTreeView.Nodes.Add(new RadTreeNode("中考"));
            ZKTreeView.Nodes.Clear();
            ZKTreeView.Nodes.Add(new RadTreeNode("数据录入"));

            ZKTreeView.Nodes.Add(new RadTreeNode("总体"));

            ZKTreeView.Nodes.Add(new RadTreeNode("区县"));
            GKTreeView.Nodes.Clear();

            GKTreeView.Nodes.Add(new RadTreeNode("数据录入"));
            GKTreeView.Nodes.Add(new RadTreeNode("总体"));
            GKTreeView.Nodes.Add(new RadTreeNode("示范校"));
            GKTreeView.Nodes.Add(new RadTreeNode("城郊"));
            GKTreeView.Nodes.Add(new RadTreeNode("区县"));


            string conn = @"Provider=vfpoledb;Data Source=" + currentdic + ";Collating Sequence=machine;";

            OleDbConnection dbfConnection = new OleDbConnection(conn);

            OleDbDataAdapter adpt = new OleDbDataAdapter("select * from " + "schoolcode", dbfConnection);
            //OleDbDataAdapter adpt = new OleDbDataAdapter("select * from " + file + " where Qk<>1", dbfConnection);
            DataSet mySet = new DataSet();

            try
            {
                adpt.Fill(mySet);
            }
            catch (OleDbException e)
            {
                throw new Exception("数据库文件被占用，请关闭！");
            }
            dbfConnection.Close();

            schoolcode_table = mySet.Tables[0];

            DataTable qxdm = schoolcode_table.AsEnumerable().GroupBy(c => c.Field<string>("qxmc")).Select(c => new
            {
                qxmc = c.Key.ToString().Trim(),
                code = string.Join(",", c.GroupBy(p => p.Field<string>("qxdm")).Select(p => p.Key.ToString().Trim()).ToArray())
            }).ToDataTable();
            int count = 0;
            foreach (DataRow dr in qxdm.Rows)
            {
                //RadTreeNode node = new RadTreeNode(dr["qxmc"].ToString().Trim());
                ZKTreeView.Nodes[2].Nodes.Add(new RadTreeNode(dr["qxmc"].ToString().Trim()));
                GKTreeView.Nodes[4].Nodes.Add(new RadTreeNode(dr["qxmc"].ToString().Trim()));
                List<string> names = get_school_name(schoolcode_table, dr["code"].ToString().Trim());
                foreach (string name in names)
                    GKTreeView.Nodes[4].Nodes[count].Nodes.Add(new RadTreeNode(name));
                count++;
                
            }
            ZKTreeView.ExpandAll();
        }

        public List<string> get_school_name(DataTable dt, string code)
        {
            string[] singles = code.Split(',');
            List<string> result = new List<string>();
            foreach (string name in singles)
            {
                List<string> some = dt.AsEnumerable().Where(c => c.Field<string>("qxdm").Equals(name)).Select(c => c.Field<string>("zxmc")).ToList<string>();
                result.AddRange(some);
            }
            return result;
        }
        private void GKTreeNode_Selected(object sender, RadTreeViewEventArgs e)
        {
            RadTreeViewElement element = sender as RadTreeViewElement;
            if (element.SelectedNode.Text.Trim().Equals("总体"))
            {
                gk_zt_panel.Show();
                gk_sf_panel.Hide();
                gk_cj_panel.Hide();
                gk_qx_panel.Hide();
                gk_data_pre_panel.Hide();
                gk_docGroupBox.Show();
            }
            else if (element.SelectedNode.Text.Trim().Equals("区县") || (element.SelectedNode.Parent != null && element.SelectedNode.Parent.Text.Trim().Equals("区县")))
            {
                gk_zt_panel.Hide();
                gk_sf_panel.Hide();
                gk_cj_panel.Hide();
                gk_qx_panel.Show();
                gk_data_pre_panel.Hide();
                gk_docGroupBox.Show();
            }
            else if (element.SelectedNode.Text.Trim().Equals("数据录入"))
            {
                gk_zt_panel.Hide();
                gk_sf_panel.Hide();
                gk_cj_panel.Hide();
                gk_qx_panel.Hide();
                gk_data_pre_panel.Show();
                gk_docGroupBox.Hide();
            }
            else if (element.SelectedNode.Text.Trim().Equals("示范校"))
            {
                gk_zt_panel.Hide();
                gk_sf_panel.Show();
                gk_cj_panel.Hide();
                gk_qx_panel.Hide();
                gk_data_pre_panel.Hide();
                gk_docGroupBox.Show();
            }
            else if (element.SelectedNode.Text.Trim().Equals("城郊"))
            {
                gk_zt_panel.Hide();
                gk_sf_panel.Hide();
                gk_cj_panel.Show();
                gk_qx_panel.Hide();
                gk_data_pre_panel.Hide();
                gk_docGroupBox.Show();
            }
        }
        private void ZKTreeNode_Selected(object sender, RadTreeViewEventArgs e)
        {
            
            RadTreeViewElement element = sender as RadTreeViewElement;
            if (element.SelectedNode.Text.Trim().Equals("总体"))
            {
                DocGroupBox.Show();
                zk_zt_panel.Show();
                zk_qx_panel.Hide();
                DataPrePanel.Hide();
            }
            else if (element.SelectedNode.Text.Trim().Equals("区县") || (element.SelectedNode.Parent != null && element.SelectedNode.Parent.Text.Trim().Equals("区县")))
            {
                DocGroupBox.Show();
                zk_qx_panel.Show();
                zk_zt_panel.Hide();
                DataPrePanel.Hide();
            }
            else if (element.SelectedNode.Text.Trim().Equals("数据录入"))
            {
                DocGroupBox.Hide();
                zk_qx_panel.Hide();
                zk_zt_panel.Hide();
                DataPrePanel.Show();
            }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            MyWizard wizard = new MyWizard("中考", this);
            wizard.Show();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            foreach (GridViewRowInfo row in zk_gridview.Rows)
            {
                if (row.Cells["checkbox"].Value != null)
                {
                    try
                    {
                        DBHelper.delete_row(
                            row.Cells["year"].Value.ToString().Trim(),
                            "中考",
                            row.Cells["sub"].Value.ToString().Trim());
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            Grid_load();
        }

        private void radButton4_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.InitialDirectory = "C://";
            openFileDialog1.Filter = "Excel files (*.xls,*.xlsx)|*.xls;*.xlsx|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.RestoreDirectory = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
                qx_addr.Text = openFileDialog1.FileName;
        }

        private void radButton5_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.InitialDirectory = "C://";
            openFileDialog1.Filter = "Excel files (*.xls,*.xlsx)|*.xls;*.xlsx|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.RestoreDirectory = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
                cj_addr.Text = openFileDialog1.FileName;
        }

        private void radButton6_Click(object sender, EventArgs e)
        {
            
            if (string.IsNullOrEmpty(qx_addr.Text.Trim()))
            {
                Error("请输入区县学校分类文件地址！");
                return;
            }
            if (string.IsNullOrEmpty(cj_addr.Text.Trim()))
            {
                Error("请输入城郊分类文件地址！");
                return;
            }

            if (CheckGridView(zk_gridview))
                return;

            string QX_code = schoolcode_table.AsEnumerable().GroupBy(c => c.Field<string>("qxmc")).Select(c => new {
                school = c.Key.ToString().Trim(), code = string.Join(",", c.GroupBy(p => p.Field<string>("qxdm")).Select(p => p.Key.ToString().Trim()).ToArray())})
                .Where(c => c.school.Equals(ZKTreeView.SelectedNode.Text.Trim())).Select(c => c.code).First();
            
            Analysis analysis = new Analysis(this);
            analysis._gridview = zk_gridview;
            analysis.qx_addr = qx_addr.Text.Trim();
            analysis.cj_addr = cj_addr.Text.Trim();
            analysis.qx_code = QX_code;

            Thread thread = new Thread(new ThreadStart(analysis.zk_qx_start));
            thread.IsBackground = true;
            thread.SetApartmentState(ApartmentState.STA);
            thread_store.Add("zk_qx", thread);
            thread.Start();
        }
        public bool CheckGridView(RadGridView gridview)
        {
            int count = 0;

            foreach (GridViewRowInfo row in gridview.Rows)
            {
                if (row.Cells["checkbox"].Value != null)
                    count++;

            }
            if (count == 0)
            {
                MessageBox.Show("没有选择任何数据", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return true;
            }
            return false;
        }
        public void ErrorM(string key, string message)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new MyErrorMessage(ErrorM), key, message);
            }
            else
            {
                Error(message);
                if(thread_store.ContainsKey(key))
                {
                    Thread thread = thread_store[key];
                    if(thread.IsAlive)
                    {
                        thread.Abort();
                        thread_store.Remove(key);
                        ShowPro(key, 2, "");
                    }

                }
                else
                {
                }

            }
        }
        private bool Error(string errormessage)
        {
            MessageBox.Show(errormessage, "提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return false;
        }

        public void ShowPro(string key, int status, string text)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new ProgressDelegate(ShowPro), key, status, text);
            }
            else
            {
                this.progress_label[key].Text = text;
                switch (status)
                {
                    case 0:
                        this.run_button[key].Enabled = false;
                        this.cancel_button[key].Enabled = true;
                        this.waiting_bar[key].StartWaiting();
                        break;
                    case 1:
                        break;
                    case 2:
                        this.run_button[key].Enabled = true;
                        this.cancel_button[key].Enabled = false;
                        this.waiting_bar[key].StopWaiting();
                        break;
                    default:
                        break;
                }


            }
        }

        private void zk_qx_cancel_Click(object sender, EventArgs e)
        {
            Thread thread = thread_store["zk_qx"];
            if (thread.IsAlive)
            {
                thread.Abort();
                thread_store.Remove("zk_qx");
                ShowPro("zk_qx", 2, "");
            }
        }

        private void zk_zt_start_Click(object sender, EventArgs e)
        {
            if (CheckGridView(zk_gridview))
                return;
            Analysis analysis = new Analysis(this);
            analysis._gridview = zk_gridview;
            analysis.CurrentDirectory = currentdic;
            analysis.save_address = save_address.Text;
            analysis.isVisible = zk_isVisible.Checked;
            Thread thread = new Thread(new ThreadStart(analysis.zk_zt_start));
            thread.IsBackground = true;
            thread.SetApartmentState(ApartmentState.STA);
            thread_store.Add("zk_zt", thread);
            thread.Start();
        }

        private void zk_zt_cancel_Click(object sender, EventArgs e)
        {
            Thread thread = thread_store["zk_zt"];
            if (thread.IsAlive)
            {
                thread.Abort();
                thread_store.Remove("zk_zt");
                ShowPro("zk_zt", 2, "");
            }
        }

        private void gk_data_import_Click(object sender, EventArgs e)
        {
            MyWizard wizard = new MyWizard("高考", this);
            wizard.Show();
        }

        private void gk_data_delete_Click(object sender, EventArgs e)
        {
            foreach (GridViewRowInfo row in gk_gridview.Rows)
            {
                if (row.Cells["checkbox"].Value != null)
                {
                    try
                    {
                        DBHelper.delete_row(
                            row.Cells["year"].Value.ToString().Trim(),
                            "高考",
                            row.Cells["sub"].Value.ToString().Trim());
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            Grid_load();
        }

        private void radButton6_Click_1(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.InitialDirectory = "C://";
            openFileDialog1.Filter = "Excel files (*.xls,*.xlsx)|*.xls;*.xlsx|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.RestoreDirectory = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
                gk_sf_addr.Text = openFileDialog1.FileName;
        }

        private void radButton2_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.InitialDirectory = "C://";
            openFileDialog1.Filter = "Excel files (*.xls,*.xlsx)|*.xls;*.xlsx|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.RestoreDirectory = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
                gk_cj_addr.Text = openFileDialog1.FileName;
        }

        private void radButton3_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.InitialDirectory = "C://";
            openFileDialog1.Filter = "Excel files (*.xls,*.xlsx)|*.xls;*.xlsx|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.RestoreDirectory = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
                gk_qx_xx_addr.Text = openFileDialog1.FileName;
        }

        private void radButton9_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.InitialDirectory = "C://";
            openFileDialog1.Filter = "Excel files (*.xls,*.xlsx)|*.xls;*.xlsx|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.RestoreDirectory = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
                gk_qx_sf_addr.Text = openFileDialog1.FileName;
        }

        private void radButton10_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.InitialDirectory = "C://";
            openFileDialog1.Filter = "Excel files (*.xls,*.xlsx)|*.xls;*.xlsx|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.RestoreDirectory = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
                gk_qx_cj_addr.Text = openFileDialog1.FileName;
        }

        private void gk_zt_start_Click(object sender, EventArgs e)
        {
            if (CheckGridView(gk_gridview))
                return;
            
            Analysis analysis = new Analysis(this);
            analysis._gridview = gk_gridview;
            analysis.save_address = gk_save_address.Text;
            analysis.isVisible = gk_isVisible.Checked;
            analysis.CurrentDirectory = currentdic;
            Thread thread = new Thread(new ThreadStart(analysis.gk_zt_start));
            thread.IsBackground = true;
            thread.SetApartmentState(ApartmentState.STA);
            thread_store.Add("gk_zt", thread);
            thread.Start();
        }

        private void gk_zt_cancel_Click(object sender, EventArgs e)
        {
            Thread thread = thread_store["gk_zt"];
            if (thread.IsAlive)
            {
                thread.Abort();
                thread_store.Remove("gk_zt");
                ShowPro("gk_zt", 2, "");
            }
        }

        private void radButton7_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog openFolder = new FolderBrowserDialog();
            openFolder.ShowNewFolderButton = true;
            openFolder.Description = "保存至";
            if (openFolder.ShowDialog() == DialogResult.OK)
                gk_save_address.Text = openFolder.SelectedPath;
        }

        private void radButton1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog openFolder = new FolderBrowserDialog();
            openFolder.ShowNewFolderButton = true;
            openFolder.Description = "保存至";
            if (openFolder.ShowDialog() == DialogResult.OK)
                save_address.Text = openFolder.SelectedPath;
        }
    }
}
