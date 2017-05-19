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
    public delegate void CheckStuIDMethod(int status, string message);
    public delegate void MyErrorMessage(string key, string Message);
    public partial class mainform : Telerik.WinControls.UI.RadForm
    {
        Dictionary<string, Thread> thread_store = new Dictionary<string, Thread>();

        Dictionary<string, RadLabel> progress_label = new Dictionary<string, RadLabel>();
        Dictionary<string, RadButton> run_button = new Dictionary<string, RadButton>();
        Dictionary<string, RadButton> cancel_button = new Dictionary<string, RadButton>();
        Dictionary<string, RadWaitingBar> waiting_bar = new Dictionary<string, RadWaitingBar>();

        Dictionary<string, string> schoolcode_kv = new Dictionary<string, string>();
        Dictionary<string, string> school_qx = new Dictionary<string, string>();

        DataTable Cust_data = new DataTable();
        CustomRelation temp_cust = new CustomRelation();
        public DataTable schoolcode_table;
        public DataTable school;
        string currentdic = System.IO.Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath);
        //Thread thread;
        public mainform()
        {
            InitializeComponent();
            Grid_load();
            TreeLoadData();
            ZKTreeView.SelectedNodeChanged += ZKTreeNode_Selected;
            GKTreeView.SelectedNodeChanged += GKTreeNode_Selected;
            HKTreeView.SelectedNodeChanged += HKTreeNode_Selected;

            gk_gridview.EditorRequired += radGridView1_EditorRequired;

            gk_isVisible.Checked = true;
            zk_isVisible.Checked = true;
            hk_group_button.IsChecked = true;
            qx_combo.SelectedIndex = 0;
            xx_combo.SelectedIndex = 0;

            init_dictionary();
            init_config_addr();
            save_address.Text = currentdic;
            gk_save_address.Text = currentdic;
            hk_save_addr.Text = currentdic;

            wk_radio.IsChecked = true;
            lk_radio.IsChecked = false;

            excellent_high.Value = 100m;
            excellent_low.Value = 85m;
            well_high.Value = 85m;
            well_low.Value = 70m;
            pass_high.Value = 70m;
            pass_low.Value = 60m;
            fail_high.Value = 60m;
            fail_low.Value = 0m;

            int curryear = DateTime.Now.Year;
            for (int i = curryear - 10; i < curryear + 10; i++)
            {
                year_list.Items.Add(i);
                zk_yearlist.Items.Add(i);
                compare_year1.Items.Add(i);
                compare_year2.Items.Add(i);
                qx_comp_year1.Items.Add(i);
                qx_comp_year2.Items.Add(i);
                cj_comp_year.Items.Add(i);
            }

            //for (int j = 1; j <= 12; j++)
            //    zk_currmonth.Items.Add(j.ToString() + "月");
            year_list.SelectedItem = curryear;
            zk_yearlist.SelectedItem = curryear;
            currmonth.SelectedItem = DateTime.Now.Month.ToString() + "月";
            zk_currmonth.SelectedItem = DateTime.Now.Month.ToString() + "月";

            compare_year2.SelectedItem = curryear;
            compare_year1.SelectedItem = curryear - 1;

            qx_comp_year2.SelectedItem = curryear;
            qx_comp_year1.SelectedItem = curryear - 1;

            cj_comp_year.SelectedItem = curryear;

            

            Cust_data.Columns.Add("name", typeof(string));
            Cust_data.Columns.Add("condition", typeof(string));

            CustomGridView.TableElement.BeginUpdate();
            CustomGridView.MasterTemplate.AllowAddNewRow = false;
            CustomGridView.MasterTemplate.AllowDragToGroup = false;
            CustomGridView.MasterTemplate.AllowEditRow = false;
            CustomGridView.DataSource = Cust_data;
            CustomGridView.TableElement.EndUpdate();
        }
        void init_config_addr()
        {
            string cj_addr = @"\config\gk_cj.xlsx";
            string sf_addr = @"\config\gk_sf.xlsx";

            gk_cj_addr.Text = currentdic + cj_addr;
            gk_qx_cj_addr.Text = currentdic + cj_addr;
            gk_xx_cj_addr.Text = currentdic + cj_addr;

            gk_sf_addr.Text = currentdic + sf_addr;
            gk_qx_sf_addr.Text = currentdic + sf_addr;
            gk_xx_sf_addr.Text = currentdic + sf_addr;

            gk_xz_cj_addr.Text = currentdic + cj_addr;

            zk_xz_qx_addr.Text = currentdic + cj_addr;
        }
        void init_dictionary()
        {
            progress_label.Add("zk_zt", zk_zt_progress);
            progress_label.Add("zk_qx", zk_qx_ProgressLabel);
            progress_label.Add("gk_zt", gk_zt_progresslabel);
            progress_label.Add("gk_cj", gk_cj_progresslabel);
            progress_label.Add("gk_sf", gk_sf_progresslabel);
            progress_label.Add("gk_qx", gk_qx_progresslabel);
            progress_label.Add("gk_xx", gk_xx_progresslabel);
            progress_label.Add("hk_zt", hk_zt_progresslabel);
            progress_label.Add("gk_cus", gk_cus_progresslabel);
            progress_label.Add("hk_script", hk_script_progresslabel);
            progress_label.Add("gk_xz", gk_xz_progresslabel);
            progress_label.Add("zk_xz", zk_xz_progresslabel);
            progress_label.Add("gk_cp", compare_progresslabel);
            progress_label.Add("gk_cj_cp", cj_comp_progresslabel);
            progress_label.Add("gk_qx_cp", qx_comp_progresslabel);

            run_button.Add("zk_zt", zk_zt_start);
            run_button.Add("zk_qx", zk_qx_start);
            run_button.Add("gk_zt", gk_zt_start);
            run_button.Add("gk_cj", gk_cj_start);
            run_button.Add("gk_sf", gk_sf_start);
            run_button.Add("gk_qx", gk_qx_start);
            run_button.Add("gk_xx", gk_xx_start);
            run_button.Add("hk_zt", hk_start);
            run_button.Add("gk_cus", gk_cus_start);
            run_button.Add("hk_script", hk_script_start);
            run_button.Add("gk_xz", gk_xz_start);
            run_button.Add("zk_xz", zk_xz_start);
            run_button.Add("gk_cp", compare_total_start);
            run_button.Add("gk_cj_cp", cj_comp_start);
            run_button.Add("gk_qx_cp", qx_comp_start);

            cancel_button.Add("zk_zt", zk_zt_cancel);
            cancel_button.Add("zk_qx", zk_qx_cancel);
            cancel_button.Add("gk_zt", gk_zt_cancel);
            cancel_button.Add("gk_cj", gk_cj_cancel);
            cancel_button.Add("gk_sf", gk_sf_cancel);
            cancel_button.Add("gk_qx", gk_qx_cancel);
            cancel_button.Add("gk_xx", gk_xx_cancel);
            cancel_button.Add("hk_zt", hk_cancel);
            cancel_button.Add("gk_cus", gk_cus_cancel);
            cancel_button.Add("hk_script", hk_script_cancel);
            cancel_button.Add("gk_xz", gk_xz_cancel);
            cancel_button.Add("zk_xz", zk_xz_cancel);
            cancel_button.Add("gk_cp", compare_total_cancel);
            cancel_button.Add("gk_cj_cp", cj_comp_cancel);
            cancel_button.Add("gk_qx_cp", qx_comp_cancel);


            waiting_bar.Add("zk_zt", zk_zt_waitingbar);
            waiting_bar.Add("zk_qx", zk_qx_WaitingBar);
            waiting_bar.Add("gk_zt", gk_zt_waitingbar);
            waiting_bar.Add("gk_cj", gk_cj_waitingbar);
            waiting_bar.Add("gk_sf", gk_sf_waitingbar);
            waiting_bar.Add("gk_qx", gk_qx_waitingbar);
            waiting_bar.Add("gk_xx", gk_xx_waitingbar);
            waiting_bar.Add("hk_zt", hk_waitingbar);
            waiting_bar.Add("gk_cus", gk_cus_waitingbar);
            waiting_bar.Add("hk_script", hk_script_progressbar);
            waiting_bar.Add("gk_xz", gk_xz_waitingbar);
            waiting_bar.Add("zk_xz", zk_xz_waitingbar);
            waiting_bar.Add("gk_cp", compare_waitingbar);
            waiting_bar.Add("gk_cj_cp", cj_comp_waitingbar);
            waiting_bar.Add("gk_qx_cp", qx_comp_waitingbar);

            
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
        public void DBCheck()
        {
            MySqlHelper.ExecuteNonQuery(MySqlHelper.ConnNoDB, CommandType.Text, "create database if not exists `examdata`", null);
            MySqlHelper.ExecuteNonQuery(MySqlHelper.Conn, CommandType.Text, "create table if not exists `exam_meta_data`"
                + "(year varchar(255),exam varchar(255),sub varchar(255),ans varchar(255),grp varchar(255),fullmark int,zh int,gtype varchar(255),gnum int)", null);
        }
        public void Grid_load()
        {
            DBCheck();
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

            zk_gridview.DataSource = zk_data.LanguageTrans("zk");
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
            gk_gridview.DataSource = gk_data.LanguageTrans("gk");

            foreach (GridViewRowInfo row in gk_gridview.Rows)
            {
                if (row.Cells["sub"].Value.ToString().Trim().Equals("语文")
                    || row.Cells["sub"].Value.ToString().Trim().Equals("英语"))
                {
                    GridViewComboBoxColumn col = (GridViewComboBoxColumn)gk_gridview.Columns["SpecChoice"] ;
                    col.DataSource = Utils.ywyy_combo;
                    row.Cells["SpecChoice"].Value = Utils.ywyy_combo[1];
                }
                else if (row.Cells["sub"].Value.ToString().Contains("理综")
                    || row.Cells["sub"].Value.ToString().Contains("文综"))
                {
                    GridViewComboBoxColumn col = (GridViewComboBoxColumn)gk_gridview.Columns["SpecChoice"];
                    col.DataSource = Utils.zh_combo;
                    row.Cells["SpecChoice"].Value = Utils.zh_combo[0];
                }
                else
                {
                    GridViewComboBoxColumn col = (GridViewComboBoxColumn)gk_gridview.Columns["SpecChoice"];
                    col.DataSource = Utils.null_combo;
                }
            }
            gk_gridview.TableElement.EndUpdate();

            HKGridView.MasterTemplate.AllowAddNewRow = false;
            HKGridView.TableElement.BeginUpdate();
            DataTable hk_data = data.AsEnumerable().AsEnumerable().Where(c => c.Field<string>("exam").Equals("hk")).Select(c => new
            {
                year = c.Field<string>("year"),
                sub = c.Field<string>("sub"),
                ans = c.Field<string>("ans"),
                grp = c.Field<string>("grp"),
                fullmark = c.Field<int>("fullmark"),
                gtype = c.Field<string>("gtype"),
                gnum = c.Field<int>("gnum")
            }).ToDataTable();
            HKGridView.DataSource = hk_data.LanguageTrans("hk");

            HKGridView.TableElement.EndUpdate();

            

        }
        void radGridView1_EditorRequired(object sender, EditorRequiredEventArgs e)
        {
            if (gk_gridview.CurrentColumn is GridViewComboBoxColumn)
            //if (gk_gridview.CurrentColumn is GridViewCheckBoxColumn)
            {
                if (gk_gridview.CurrentRow.Cells["sub"].Value.ToString().Trim().Equals("语文")
                    || gk_gridview.CurrentRow.Cells["sub"].Value.ToString().Trim().Equals("英语"))
                {
                    GridViewComboBoxColumn col = (GridViewComboBoxColumn)gk_gridview.CurrentColumn;
                    col.DataSource = Utils.ywyy_combo;
                    gk_gridview.CurrentRow.Cells["SpecChoice"].Value = Utils.ywyy_combo[1];
                }
                else if (gk_gridview.CurrentRow.Cells["sub"].Value.ToString().Contains("理综")
                    || gk_gridview.CurrentRow.Cells["sub"].Value.ToString().Contains("文综"))
                {
                    GridViewComboBoxColumn col = (GridViewComboBoxColumn)gk_gridview.CurrentColumn;
                    col.DataSource = Utils.zh_combo;
                    gk_gridview.CurrentRow.Cells["SpecChoice"].Value = Utils.zh_combo[0];
                }
                else
                {
                    GridViewComboBoxColumn col = (GridViewComboBoxColumn)gk_gridview.CurrentColumn;
                    col.DataSource = Utils.null_combo;
                }
            }
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
            ZKTreeView.Nodes.Add(new RadTreeNode("行政版"));

            GKTreeView.Nodes.Clear();

            GKTreeView.Nodes.Add(new RadTreeNode("数据录入"));
            GKTreeView.Nodes.Add(new RadTreeNode("总体"));
            GKTreeView.Nodes.Add(new RadTreeNode("示范校"));
            GKTreeView.Nodes.Add(new RadTreeNode("城郊"));
            GKTreeView.Nodes.Add(new RadTreeNode("自定义"));
            GKTreeView.Nodes.Add(new RadTreeNode("区县"));
            GKTreeView.Nodes.Add(new RadTreeNode("学校"));
            GKTreeView.Nodes.Add(new RadTreeNode("行政版"));

            HKTreeView.Nodes.Clear();
            HKTreeView.Nodes.Add(new RadTreeNode("数据录入"));
            HKTreeView.Nodes.Add(new RadTreeNode("总体"));
            HKTreeView.Nodes.Add(new RadTreeNode("成绩报告单"));
            string conn = @"Provider=vfpoledb;Data Source=" + currentdic + ";Collating Sequence=machine;";

            OleDbConnection dbfConnection = new OleDbConnection(conn);

            OleDbDataAdapter adpt = new OleDbDataAdapter("select * from " + "schoolcode", dbfConnection);
            OleDbDataAdapter adpt2 = new OleDbDataAdapter("select * from " + "school", dbfConnection);
            //OleDbDataAdapter adpt = new OleDbDataAdapter("select * from " + file + " where Qk<>1", dbfConnection);
            DataSet mySet = new DataSet();
            DataSet mySet2 = new DataSet();
            try
            {
                adpt.Fill(mySet);
                adpt2.Fill(mySet2);
            }
            catch (OleDbException e)
            {
                throw new Exception("数据库文件被占用，请关闭！");
            }
            dbfConnection.Close();

            schoolcode_table = mySet.Tables[0];
            school = mySet2.Tables[0];
            DataTable qxdm = schoolcode_table.AsEnumerable().GroupBy(c => c.Field<string>("qxmc")).Select(c => new
            {
                qxmc = c.Key.ToString().Trim(),
                code = string.Join(",", c.GroupBy(p => p.Field<string>("qxdm")).Select(p => p.Key.ToString().Trim()).ToArray())
            }).ToDataTable();

            qx_comp_qx.DataSource = qxdm;
            qx_comp_qx.DisplayMember = "qxmc";
            qx_comp_qx.ValueMember = "code";
            qx_comp_comp.DataSource = qxdm;
            qx_comp_comp.DisplayMember = "qxmc";
            qx_comp_comp.ValueMember = "code";

            int count = 0;
            GKTreeView.Nodes[6].CheckType = CheckType.CheckBox;
            foreach (DataRow dr in qxdm.Rows)
            {
                //RadTreeNode node = new RadTreeNode(dr["qxmc"].ToString().Trim());
                ZKTreeView.Nodes[2].Nodes.Add(new RadTreeNode(dr["qxmc"].ToString().Trim()));
                GKTreeView.Nodes[5].Nodes.Add(new RadTreeNode(dr["qxmc"].ToString().Trim()));
                GKTreeView.Nodes[6].Nodes.Add(new RadTreeNode(dr["qxmc"].ToString().Trim()));
                GKTreeView.Nodes[6].Nodes[count].CheckType = CheckType.CheckBox;
                
                List<string> names = get_school_name(schoolcode_table, dr["code"].ToString().Trim());
                int children_count = 0;
                foreach (string name in names)
                {
                    GKTreeView.Nodes[6].Nodes[count].Nodes.Add(new RadTreeNode(name));
                    GKTreeView.Nodes[6].Nodes[count].Nodes[children_count].CheckType = CheckType.CheckBox;
                    children_count++;
                }
                count++;
                
            }
            schoolcode_kv = schoolcode_table.AsEnumerable().Select(c => new 
            {
                key = c.Field<string>("zxmc").ToString().Trim(),
                value = c.Field<string>("zxdm")
            }).ToDictionary(c => c.key, c => c.value);

            school_qx = schoolcode_table.AsEnumerable().Select(c => new
            {
                school = c.Field<string>("zxdm"),
                qx = c.Field<string>("qxmc").Trim()
            }).Join(qxdm.AsEnumerable(), s => s.qx, c => c.Field<string>("qxmc"), (s, c) => new
            {
                school = s.school,
                qx = c.Field<string>("code")

            }).ToDictionary(c => c.school, c => c.qx);
            GKTreeView.NodeCheckedChanged += new RadTreeView.TreeViewEventHandler(GKTreeView_NodeCheckedChanged);
            ZKTreeView.ExpandAll();

            DataTable qx_data =
                schoolcode_table.AsEnumerable().GroupBy(c => c.Field<string>("qxmc")).Select(c => new
                {
                    school = c.Key.ToString().Trim(),
                    code = string.Join(",", c.GroupBy(p => p.Field<string>("qxdm")).Select(p => p.Key.ToString().Trim()).ToArray())
                }).ToDataTable();

            DataRow total = qx_data.NewRow();
            total["school"] = "全部";
            total["code"] = "0";
            qx_data.Rows.InsertAt(total, 0);

            qx_combo.DataSource = qx_data;
            qx_combo.DisplayMember = "school";
            qx_combo.ValueMember = "code";

            qx_combo.ResetText();
        }
        private void GKTreeView_NodeCheckedChanged(object sender, RadTreeViewEventArgs e)
        {
            CheckAllChildNodes(e.Node, e.Node.Checked);
            //bool bol = true;
            //if (e.Node.Parent != null)
            //{
            //    for (int i = 0; i < e.Node.Parent.Nodes.Count; i++)
            //    {
            //        if (!e.Node.Parent.Nodes[i].Checked)
            //            bol = false;
            //    }
            //    e.Node.Parent.Checked = bol;
            //}

        }
        public void CheckAllChildNodes(RadTreeNode treenode, bool nodechecked)
        {
            foreach (RadTreeNode node in treenode.Nodes)
            {
                node.Checked = nodechecked;
                if(node.Nodes.Count > 0)
                    this.CheckAllChildNodes(node, nodechecked);
            }
        }
        
        public List<string> get_school_name(DataTable dt, string code)
        {
            string[] singles = code.Split(',');
            List<string> result = new List<string>();
            foreach (string name in singles)
            {
                List<string> some = dt.AsEnumerable().Where(c => c.Field<string>("qxdm").Equals(name)).Select(c => c.Field<string>("zxmc").ToString().Trim()).ToList<string>();
                result.AddRange(some);
            }
            return result;
        }
        private void HKTreeNode_Selected(object sender, RadTreeViewEventArgs e)
        {
            RadTreeViewElement element = sender as RadTreeViewElement;
            if (element.SelectedNode.Text.Trim().Equals("数据录入"))
            {
                hk_pre_panel.Show();
                hk_zt_panel.Hide();
                hk_config_panel.Hide();
                hk_markReporter_panel.Hide();
            }
            else if (element.SelectedNode.Text.Trim().Equals("总体"))
            {
                hk_pre_panel.Hide();
                hk_zt_panel.Show();
                hk_config_panel.Show();
                hk_markReporter_panel.Hide();
            }
            else if (element.SelectedNode.Text.Trim().Equals("成绩报告单"))
            {
                hk_pre_panel.Hide();
                hk_markReporter_panel.Show();
                hk_config_panel.Show();
                hk_zt_panel.Hide();

                
            }
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
                gk_xx_panel.Hide();
                custom_panel.Hide();
                gk_xz_panel.Hide();
            }
            else if (element.SelectedNode.Text.Trim().Equals("区县") || (element.SelectedNode.Parent != null && element.SelectedNode.Parent.Text.Trim().Equals("区县")))
            {
                gk_zt_panel.Hide();
                gk_sf_panel.Hide();
                gk_cj_panel.Hide();
                gk_qx_panel.Show();
                gk_data_pre_panel.Hide();
                gk_docGroupBox.Show();
                gk_xx_panel.Hide();
                custom_panel.Hide();
                gk_xz_panel.Hide();
            }
            else if (element.SelectedNode.Text.Trim().Equals("数据录入"))
            {
                gk_zt_panel.Hide();
                gk_sf_panel.Hide();
                gk_cj_panel.Hide();
                gk_qx_panel.Hide();
                gk_data_pre_panel.Show();
                gk_docGroupBox.Hide();
                gk_xx_panel.Hide();
                custom_panel.Hide();
                gk_xz_panel.Hide();
            }
            else if (element.SelectedNode.Text.Trim().Equals("示范校"))
            {
                gk_zt_panel.Hide();
                gk_sf_panel.Show();
                gk_cj_panel.Hide();
                gk_qx_panel.Hide();
                gk_data_pre_panel.Hide();
                gk_docGroupBox.Show();
                gk_xx_panel.Hide();
                custom_panel.Hide();
                gk_xz_panel.Hide();
            }
            else if (element.SelectedNode.Text.Trim().Equals("城郊"))
            {
                gk_zt_panel.Hide();
                gk_sf_panel.Hide();
                gk_cj_panel.Show();
                gk_qx_panel.Hide();
                gk_data_pre_panel.Hide();
                gk_docGroupBox.Show();
                gk_xx_panel.Hide();
                custom_panel.Hide();
                gk_xz_panel.Hide();
            }
            else if (element.SelectedNode.Text.Trim().Equals("自定义"))
            {
                gk_zt_panel.Hide();
                gk_sf_panel.Hide();
                gk_cj_panel.Hide();
                gk_qx_panel.Hide();
                gk_data_pre_panel.Hide();
                gk_docGroupBox.Show();
                gk_xx_panel.Hide();
                custom_panel.Show();
                gk_xz_panel.Hide();
            }
            else if (element.SelectedNode.Text.Trim().Equals("行政版"))
            {
                gk_xz_panel.Show();
                gk_zt_panel.Hide();
                gk_sf_panel.Hide();
                gk_cj_panel.Hide();
                gk_qx_panel.Hide();
                gk_data_pre_panel.Hide();
                gk_docGroupBox.Show();
                gk_xx_panel.Hide();
                custom_panel.Hide();
            }

            else if (element.SelectedNode.Text.Trim().Equals("学校")
                || (element.SelectedNode.Parent != null && element.SelectedNode.Parent.Text.Trim().Equals("学校"))
                || (element.SelectedNode.Parent.Parent != null && element.SelectedNode.Parent.Parent.Text.Trim().Equals("学校")))
            {
                gk_zt_panel.Hide();
                gk_sf_panel.Hide();
                gk_cj_panel.Hide();
                gk_qx_panel.Hide();
                gk_data_pre_panel.Hide();
                gk_docGroupBox.Show();
                gk_xx_panel.Show();
                custom_panel.Hide();
                gk_xz_panel.Hide();
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
                zk_xz_panel.Hide();
            }
            else if (element.SelectedNode.Text.Trim().Equals("区县") || (element.SelectedNode.Parent != null && element.SelectedNode.Parent.Text.Trim().Equals("区县")))
            {
                DocGroupBox.Show();
                zk_qx_panel.Show();
                zk_zt_panel.Hide();
                DataPrePanel.Hide();
                zk_xz_panel.Hide();
            }
            else if (element.SelectedNode.Text.Trim().Equals("数据录入"))
            {
                DocGroupBox.Hide();
                zk_qx_panel.Hide();
                zk_zt_panel.Hide();
                DataPrePanel.Show();
                zk_xz_panel.Hide();
            }
            else if (element.SelectedNode.Text.Trim().Equals("行政版"))
            {
                DocGroupBox.Show();
                zk_qx_panel.Hide();
                zk_zt_panel.Hide();
                DataPrePanel.Hide();
                zk_xz_panel.Show();
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
            analysis.curryear = zk_yearlist.SelectedItem.ToString().Trim();
            analysis.currmonth = zk_currmonth.SelectedItem.ToString().Trim();
            analysis.CurrentDirectory = currentdic;
            analysis.save_address = save_address.Text;
            analysis.isVisible = zk_isVisible.Checked;

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
                if (row.Cells["checkbox"].Value != null && (bool)row.Cells["checkbox"].Value)
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
                        if (thread_store.ContainsKey(key))
                            thread_store.Remove(key);
                        break;
                    default:
                        break;
                }


            }
        }

        private void zk_qx_cancel_Click(object sender, EventArgs e)
        {
            if (thread_store.ContainsKey("zk_qx"))
            {
                Thread thread = thread_store["zk_qx"];
                if (thread.IsAlive)
                {
                    thread.Abort();
                    thread_store.Remove("zk_qx");
                    ShowPro("zk_qx", 2, "");
                }
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
            analysis.curryear = zk_yearlist.SelectedItem.ToString().Trim();
            analysis.currmonth = zk_currmonth.SelectedItem.ToString().Trim();
            Thread thread = new Thread(new ThreadStart(analysis.zk_zt_start));
            thread.IsBackground = true;
            thread.SetApartmentState(ApartmentState.STA);
            thread_store.Add("zk_zt", thread);
            thread.Start();
        }

        private void zk_zt_cancel_Click(object sender, EventArgs e)
        {
            if (thread_store.ContainsKey("zk_zt"))
            {
                Thread thread = thread_store["zk_zt"];
                if (thread.IsAlive)
                {
                    thread.Abort();
                    thread_store.Remove("zk_zt");
                    ShowPro("zk_zt", 2, "");
                }
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
                if (row.Cells["checkbox"].Value != null && (bool)row.Cells["checkbox"].Value == true)
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
            analysis.curryear = year_list.SelectedItem.ToString().Trim();
            analysis.currmonth = currmonth.SelectedItem.ToString().Trim();
            Thread thread = new Thread(new ThreadStart(analysis.gk_zt_start));
            thread.IsBackground = true;
            thread.SetApartmentState(ApartmentState.STA);
            thread_store.Add("gk_zt", thread);
            thread.Start();
        }

        private void gk_zt_cancel_Click(object sender, EventArgs e)
        {
            if (thread_store.ContainsKey("gk_zt"))
            {
                Thread thread = thread_store["gk_zt"];
                if (thread.IsAlive)
                {
                    thread.Abort();
                    thread_store.Remove("gk_zt");
                    ShowPro("gk_zt", 2, "");
                }
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

        private void gk_cj_start_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(gk_cj_addr.Text.Trim()))
            {
                Error("请输入城郊分类文件地址！");
                return;
            }
            
            if (CheckGridView(gk_gridview))
                return;

            Analysis analysis = new Analysis(this);
            analysis._gridview = gk_gridview;
            analysis.save_address = gk_save_address.Text;
            analysis.isVisible = gk_isVisible.Checked;
            analysis.CurrentDirectory = currentdic;
            analysis.cj_addr = gk_cj_addr.Text.ToString().Trim();
            analysis.curryear = year_list.SelectedItem.ToString().Trim();
            analysis.currmonth = currmonth.SelectedItem.ToString().Trim();
            Thread thread = new Thread(new ThreadStart(analysis.gk_cj_start));
            thread.IsBackground = true;
            thread.SetApartmentState(ApartmentState.STA);
            thread_store.Add("gk_cj", thread);
            thread.Start();
        }

        private void gk_sf_start_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(gk_sf_addr.Text.Trim()))
            {
                Error("请输入示范学校分类文件地址！");
                return;
            }
            
            if (CheckGridView(gk_gridview))
                return;

            Analysis analysis = new Analysis(this);
            analysis._gridview = gk_gridview;
            analysis.save_address = gk_save_address.Text;
            analysis.isVisible = gk_isVisible.Checked;
            analysis.CurrentDirectory = currentdic;
            analysis.sf_addr = gk_sf_addr.Text.ToString().Trim();
            analysis.curryear = year_list.SelectedItem.ToString().Trim();
            analysis.currmonth = currmonth.SelectedItem.ToString().Trim();
            Thread thread = new Thread(new ThreadStart(analysis.gk_sf_start));
            thread.IsBackground = true;
            thread.SetApartmentState(ApartmentState.STA);
            thread_store.Add("gk_sf", thread);
            thread.Start();
        }

        private void gk_sf_cancel_Click(object sender, EventArgs e)
        {
            if (thread_store.ContainsKey("gk_sf"))
            {
                Thread thread = thread_store["gk_sf"];
                if (thread.IsAlive)
                {
                    thread.Abort();
                    thread_store.Remove("gk_sf");
                    ShowPro("gk_sf", 2, "");
                }
            }
        }

        private void gk_qx_start_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(gk_qx_xx_addr.Text.Trim()))
            {
                Error("请输入区县学校分类文件地址！");
                return;
            }
            if (string.IsNullOrEmpty(gk_qx_cj_addr.Text.Trim()))
            {
                Error("请输入城郊分类文件地址！");
                return;
            }
            if (string.IsNullOrEmpty(gk_qx_sf_addr.Text.Trim()))
            {
                Error("请输入示范学校分类文件地址！");
                return;
            }
            if (CheckGridView(gk_gridview))
                return;
            if (GKTreeView.SelectedNode.Text.Trim().Equals("区县"))
            {
                Error("请选择区县！");
                return;
            }
            Analysis analysis = new Analysis(this);
            analysis._gridview = gk_gridview;
            analysis.save_address = gk_save_address.Text;
            analysis.isVisible = gk_isVisible.Checked;
            analysis.CurrentDirectory = currentdic;
            analysis.qx_addr = gk_qx_xx_addr.Text.Trim();
            analysis.cj_addr = gk_qx_cj_addr.Text.Trim();
            analysis.sf_addr = gk_qx_sf_addr.Text.Trim();
            analysis.curryear = year_list.SelectedItem.ToString().Trim();
            analysis.currmonth = currmonth.SelectedItem.ToString().Trim();
            string QX_code = schoolcode_table.AsEnumerable().GroupBy(c => c.Field<string>("qxmc")).Select(c => new
            {
                school = c.Key.ToString().Trim(),
                code = string.Join(",", c.GroupBy(p => p.Field<string>("qxdm")).Select(p => p.Key.ToString().Trim()).ToArray())
            })
                .Where(c => c.school.Equals(GKTreeView.SelectedNode.Text.Trim())).Select(c => c.code).First();

            analysis.qx_code = QX_code;
            analysis.QX = GKTreeView.SelectedNode.Text.Trim();
            Thread thread = new Thread(new ThreadStart(analysis.gk_qx_start));
            thread.IsBackground = true;
            thread.SetApartmentState(ApartmentState.STA);
            thread_store.Add("gk_qx", thread);
            thread.Start();
        }

        private void gk_qx_cancel_Click(object sender, EventArgs e)
        {
            if (thread_store.ContainsKey("gk_qx"))
            {
                Thread thread = thread_store["gk_qx"];
                if (thread.IsAlive)
                {
                    thread.Abort();
                    thread_store.Remove("gk_qx");
                    ShowPro("gk_qx", 2, "");
                }
            }
        }

        private void gk_cj_cancel_Click(object sender, EventArgs e)
        {
            if (thread_store.ContainsKey("gk_cj"))
            {
                Thread thread = thread_store["gk_cj"];
                if (thread.IsAlive)
                {
                    thread.Abort();
                    thread_store.Remove("gk_cj");
                    ShowPro("gk_cj", 2, "");
                }
            }
        }

        private void radButton11_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.InitialDirectory = "C://";
            openFileDialog1.Filter = "Excel files (*.xls,*.xlsx)|*.xls;*.xlsx|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.RestoreDirectory = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
                gk_xx_sf_addr.Text = openFileDialog1.FileName;
        }

        private void radButton8_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.InitialDirectory = "C://";
            openFileDialog1.Filter = "Excel files (*.xls,*.xlsx)|*.xls;*.xlsx|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.RestoreDirectory = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
                gk_xx_cj_addr.Text = openFileDialog1.FileName;
        }

        private void gk_xx_start_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(gk_xx_sf_addr.Text.Trim()))
            {
                Error("请输入区县学校分类文件地址！");
                return;
            }
            if (string.IsNullOrEmpty(gk_xx_cj_addr.Text.Trim()))
            {
                Error("请输入城郊分类文件地址！");
                return;
            }
            if (CheckGridView(gk_gridview))
                return;
            Dictionary<string, string> school = TreeViewCheck(GKTreeView.Nodes[6]);

            if (school.Count == 0)
            {
                Error("请勾选报告学校！");
                return;
            }
            Analysis analysis = new Analysis(this);
            analysis._gridview = gk_gridview;
            analysis.save_address = gk_save_address.Text;
            analysis.isVisible = gk_isVisible.Checked;
            analysis.CurrentDirectory = currentdic;
            analysis.cj_addr = gk_xx_cj_addr.Text.Trim();
            analysis.sf_addr = gk_xx_sf_addr.Text.Trim();
            analysis.school = school;
            analysis.school_qx = school_qx;
            analysis.curryear = year_list.SelectedItem.ToString().Trim();
            analysis.currmonth = currmonth.SelectedItem.ToString().Trim();
            Thread thread = new Thread(new ThreadStart(analysis.gk_xx_start));
            thread.IsBackground = true;
            thread.SetApartmentState(ApartmentState.STA);
            thread_store.Add("gk_xx", thread);
            thread.Start();
        }

        private Dictionary<string, string> TreeViewCheck(RadTreeNode treenode)
        {
            Dictionary<string, string> result = new Dictionary<string,string>();
            if (treenode.Checked)
                return schoolcode_kv;

            foreach (RadTreeNode node in treenode.Nodes)
            {
                if (node.Nodes.Count != 0)
                    foreach (RadTreeNode child in node.Nodes)
                        if (child.Checked)
                            result.Add(child.Name, schoolcode_kv[child.Name]);
            }
            return result;

        }

        private void gk_xx_cancel_Click(object sender, EventArgs e)
        {
            if (thread_store.ContainsKey("gk_xx"))
            {
                Thread thread = thread_store["gk_xx"];
                if (thread.IsAlive)
                {
                    thread.Abort();
                    thread_store.Remove("gk_xx");
                    ShowPro("gk_xx", 2, "");
                }
            }
        }

        private void radButton12_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog openFolder = new FolderBrowserDialog();
            openFolder.ShowNewFolderButton = true;
            openFolder.Description = "保存至";
            if (openFolder.ShowDialog() == DialogResult.OK)
                hk_save_addr.Text = openFolder.SelectedPath;
        }

        private void hk_start_Click(object sender, EventArgs e)
        {
            if (CheckGridView(HKGridView))
                return;
            if (!hk_check())
                return;
            Analysis analysis = new Analysis(this);
            analysis._gridview = HKGridView;
            analysis.save_address = hk_save_addr.Text;
            analysis.isVisible = hk_isVisible.Checked;
            analysis.CurrentDirectory = currentdic;
            analysis.hk_hierarchy = new Analysis.HK_hierarchy();
            analysis.hk_hierarchy.excellent_low = excellent_low.Value;
            analysis.hk_hierarchy.excellent_high = excellent_high.Value;
            analysis.hk_hierarchy.well_low = well_low.Value;
            analysis.hk_hierarchy.well_high = well_high.Value;
            analysis.hk_hierarchy.pass_low = pass_low.Value;
            analysis.hk_hierarchy.pass_high = pass_high.Value;
            analysis.hk_hierarchy.fail_low = fail_low.Value;
            analysis.hk_hierarchy.fail_high = fail_high.Value;
            Thread thread = new Thread(new ThreadStart(analysis.hk_zt_start));
            thread.IsBackground = true;
            thread.SetApartmentState(ApartmentState.STA);
            thread_store.Add("hk_zt", thread);
            thread.Start();

        }

        private bool hk_check()
        {
            if (Math.Abs(excellent_low.Value) != excellent_low.Value ||
                Math.Abs(excellent_high.Value) != excellent_high.Value ||
                Math.Abs(well_low.Value) != well_low.Value ||
                Math.Abs(well_high.Value) != well_high.Value ||
                Math.Abs(pass_low.Value) != pass_low.Value ||
                Math.Abs(pass_high.Value) != pass_high.Value ||
                Math.Abs(fail_low.Value) != fail_low.Value ||
                Math.Abs(fail_high.Value) != fail_high.Value)
                return Error("会考成绩区域不能为负！");
            if (!(fail_low.Value < fail_high.Value &&
                pass_low.Value < pass_high.Value &&
                well_low.Value < well_high.Value &&
                excellent_low.Value < excellent_high.Value))
                return Error("会考成绩设置错误！");
            return true;
        }

        private void hk_cancel_Click(object sender, EventArgs e)
        {
            if (thread_store.ContainsKey("hk_zt"))
            {
                Thread thread = thread_store["hk_zt"];
                if (thread.IsAlive)
                {
                    thread.Abort();
                    thread_store.Remove("hk_zt");
                    ShowPro("hk_zt", 2, "");
                }
            }
        }

        private void hk_import_Click(object sender, EventArgs e)
        {
            MyWizard wizard = new MyWizard("会考", this);
            wizard.Show();
        }

        private void hk_delete_Click(object sender, EventArgs e)
        {
            foreach (GridViewRowInfo row in HKGridView.Rows)
            {
                if (row.Cells["checkbox"].Value != null && (bool)row.Cells["checkbox"].Value == true)
                {
                    try
                    {
                        DBHelper.delete_row(
                            row.Cells["year"].Value.ToString().Trim(),
                            "会考",
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

        private void custom_col_Click(object sender, EventArgs e)
        {
            List<string> names = new List<string>();
            int count = 0;
            foreach (GridViewRowInfo row in gk_gridview.Rows)
            {
                if (row.Cells["checkbox"].Value != null && (bool)row.Cells["checkbox"].Value == true)
                {
                    string year = row.Cells["year"].Value.ToString().Trim();
                    string exam = "gk";
                    string chi_sub = row.Cells["sub"].Value.ToString().Trim();
                    string sub = Utils.language_trans(chi_sub);

                    
                    MetaData mdata = new MetaData(year, exam, sub);

                    names.AddRange(mdata.get_column_name());
                    count++;
                }
            }
            
            List<string> name = names.GroupBy(c => c).Select(c => new
            {
                count = c.Count(),
                name = c.Key.Trim()
            }).Where(c => c.count == count).Select(c => c.name).ToList();

            custom_col.DataSource = name;
        }

        private void custom_insert_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(custom_name.Text.Trim()))
            {
                Error("标记名不能为空");
                return;
            }
            if (custom_col.SelectedItem == null)
            {
                Error("列名不能为空");
                return;
            }
            if (custom_comp.SelectedItem == null)
            {
                Error("条件不能为空");
                return;
            }
            if (string.IsNullOrEmpty(custom_value.Text.Trim()))
            {
                Error("值不能为空");
                return;
            }

            if (string.IsNullOrEmpty(custom_result.Text.Trim()))
            {
                temp_cust.set_tag(custom_name.Text.Trim());
                custom_result.Text = custom_name.Text.Trim() + ":   ";
            }
            else
            {
                if (custom_relation.SelectedItem == null)
                {
                    Error("关系名不能为空");
                    return;
                }
                custom_result.Text = custom_result.Text + " " + Utils.OperatorTrans(custom_relation.SelectedItem.Text.Trim());
                temp_cust.insert(Utils.OperatorTrans(custom_relation.SelectedItem.Text.Trim()));
            }

            StringBuilder sb = new StringBuilder();
            sb.Append(custom_col.SelectedItem.Text.Trim());
            sb.Append(" " + Utils.OperatorTrans(custom_comp.SelectedItem.Text.Trim()) + " ");
            string type = "varchar";
            foreach (GridViewRowInfo row in gk_gridview.Rows)
            {
                if (row.Cells["checkbox"].Value != null && (bool)row.Cells["checkbox"].Value == true)
                {
                    string year = row.Cells["year"].Value.ToString().Trim();
                    string exam = "gk";
                    string chi_sub = row.Cells["sub"].Value.ToString().Trim();
                    string sub = Utils.language_trans(chi_sub);


                    MetaData mdata = new MetaData(year, exam, sub);

                    type = mdata.get_column_type(custom_col.SelectedItem.Text.Trim());
                    break;
                }
            }
            if (type.Contains("char"))
            {
                sb.Append("'");
                sb.Append(custom_value.Text.Trim());
                sb.Append("'");
            }
            else if (custom_comp.SelectedItem.Text.Trim().Equals("近似于"))
            {
                Error("不能近似于一个数字！");
                if (!temp_cust.isEmpty())
                    temp_cust.revoke();
                custom_col.ResetText();
                custom_comp.ResetText();
                custom_value.Text = "";
                custom_relation.ResetText();
                custom_result.Text = "";
                return;
            }
            else
                sb.Append(custom_value.Text.Trim());

            custom_result.Text += " " + sb.ToString();
            temp_cust.insert(sb.ToString());
        }

        private void radButton13_Click(object sender, EventArgs e)
        {
            custom_col.ResetText();
            custom_comp.ResetText();
            custom_value.Text = "";
            custom_relation.ResetText();
            custom_result.Text = "";

            temp_cust.reset();
        }

        private void cus_result_insert_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(custom_result.Text.Trim()))
            {
                Error("输入为空");
                return;
            }

            DataRow newrow = Cust_data.NewRow();
            newrow["name"] = temp_cust._tag;
            newrow["condition"] = temp_cust.get_relation();

            CustomGridView.TableElement.BeginUpdate();
            Cust_data.Rows.Add(newrow);
            CustomGridView.TableElement.EndUpdate();

            radButton13_Click(null, null);

        }

        private void cus_result_revoke_Click(object sender, EventArgs e)
        {
            GridViewDataRowInfo dataRowInfo = CustomGridView.CurrentRow as GridViewDataRowInfo;
            if (dataRowInfo != null)
            {
                CustomGridView.Rows.Remove(dataRowInfo);
            }
        }

        private void gk_cus_start_Click(object sender, EventArgs e)
        {
            if (CheckGridView(gk_gridview))
                return;

            if (Cust_data.Rows.Count == 0)
            {
                Error("自定义分组不能为空，请至少定义一组");
                return;
            }

            Analysis analysis = new Analysis(this);
            analysis.custom_data = Cust_data.Copy();
            analysis._gridview = gk_gridview;
            analysis.save_address = gk_save_address.Text;
            analysis.isVisible = gk_isVisible.Checked;
            analysis.CurrentDirectory = currentdic;
            analysis.curryear = year_list.SelectedItem.ToString().Trim();
            analysis.currmonth = currmonth.SelectedItem.ToString().Trim();
            Thread thread = new Thread(new ThreadStart(analysis.gk_custom_start));
            thread.IsBackground = true;
            thread.SetApartmentState(ApartmentState.STA);
            thread_store.Add("gk_cus", thread);
            thread.Start();
        }

        private void gk_cus_cancel_Click(object sender, EventArgs e)
        {
            if (thread_store.ContainsKey("gk_cus"))
            {
                Thread thread = thread_store["gk_cus"];
                if (thread.IsAlive)
                {
                    thread.Abort();
                    thread_store.Remove("gk_cus");
                    ShowPro("gk_cus", 2, "");
                }
            }
        }

        public void CheckStuID(int status, string message)
        {
            if (this.InvokeRequired)
                this.Invoke(new CheckStuIDMethod(CheckStuID), status, message);
            else
            {
                switch (status)
                {
                    case 1:

                        MessageBoxButtons messButton = MessageBoxButtons.OKCancel;
                        DialogResult dr = MessageBox.Show(message + "\n仍然继续嘛？", "是否继续", messButton);
                        if (dr == DialogResult.Cancel)
                        {
                            ShowPro("hk_script", 100, "取消");
                        }
                        break;
                    case 2:

                        Error(message);
                        ShowPro("hk_script", 100, "错误退出");
                        break;
                    case 3:
                        MessageBoxButtons deleteButton = MessageBoxButtons.OKCancel;
                        DialogResult delete_dr = MessageBox.Show(message + "\n删除重复成绩继续？", "是否继续", deleteButton);
                        if (delete_dr == DialogResult.Cancel)
                        {
                            ShowPro("hk_script", 100, "取消");
                        }
                        break;
                    default:
                        break;
                }
            }
        }

        private void hk_script_start_Click(object sender, EventArgs e)
        {
            if (CheckGridView(HKGridView))
                return;
            if (!hk_check())
                return;
            Analysis analysis = new Analysis(this);
            analysis._gridview = HKGridView;
            analysis.save_address = hk_save_addr.Text;
            analysis.isVisible = hk_isVisible.Checked;
            analysis.CurrentDirectory = currentdic;
            analysis.hk_hierarchy = new Analysis.HK_hierarchy();
            analysis.hk_hierarchy.excellent_low = excellent_low.Value;
            analysis.hk_hierarchy.excellent_high = excellent_high.Value;
            analysis.hk_hierarchy.well_low = well_low.Value;
            analysis.hk_hierarchy.well_high = well_high.Value;
            analysis.hk_hierarchy.pass_low = pass_low.Value;
            analysis.hk_hierarchy.pass_high = pass_high.Value;
            analysis.hk_hierarchy.fail_low = fail_low.Value;
            analysis.hk_hierarchy.fail_high = fail_high.Value;
            analysis.date = dateTimePicker.Value.Year.ToString() + "年" + dateTimePicker.Value.Month.ToString() + "月";
            if (hk_group_button.IsChecked)
            {
                analysis.hk_script_type = Utils.UnionType.QX_XX;
                analysis.qx = qx_combo.SelectedValue.ToString().Trim();
                analysis.xx = xx_combo.SelectedValue.ToString().Trim();
                analysis.qx_name = qx_combo.Text.Trim();
                analysis.xx_name = xx_combo.Text.Trim();
            }

            if (hk_id_button.IsChecked)
            {
                analysis.hk_script_type = Utils.UnionType.ID;
                analysis.stu_id = stu_id.Text.Trim();
            }
            Thread thread = new Thread(new ThreadStart(analysis.hk_script_start));
            thread.IsBackground = true;
            thread.SetApartmentState(ApartmentState.STA);
            thread_store.Add("hk_script", thread);
            thread.Start();
        }

        private void qx_combo_SelectedIndexChanged(object sender, Twin.UI.Data.PositionChangedEventArgs e)
        {
            if (qx_combo.Text.Equals(""))
                return;
            if (qx_combo.DataSource != null && !qx_combo.Text.Equals("0"))
            {

                string code_str = qx_combo.SelectedValue.ToString();
                string[] code = code_str.Split(new char[1] { ',' });

                DataTable DT = school.AsEnumerable().Where(c => code.Contains(c.Field<string>("xxdm").ToString().Trim().Substring(0, 2))).Select(c => new
                {
                    code = c.Field<string>("xxdm").ToString().Trim(),
                    school = c.Field<string>("xxmc").ToString().Trim()
                }).ToDataTable();
                DataRow total = DT.NewRow();
                total["code"] = "0";
                total["school"] = "全部";
                DT.Rows.InsertAt(total, 0);

                xx_combo.DataSource = DT;

                xx_combo.DisplayMember = "school";
                xx_combo.ValueMember = "code";

                xx_combo.ResetText();
            }
            else
            {
                xx_combo.DataSource = null;
                xx_combo.ResetText();
            }
        }

        private void gk_xz_start_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(gk_xz_cj_addr.Text.Trim()))
            {
                Error("请输入城郊分类文件地址！");
                return;
            }

            if (wen_first_level.Value == 0 || wen_second_level.Value == 0 || wen_third_level.Value == 0)
            {
                Error("重点线不能为零！");
                return;
            }
            if (wen_first_level.Value < wen_second_level.Value)
            {
                Error("一本线不应该低于二本线吧？");
                return;
            }
            if (wen_second_level.Value < wen_third_level.Value)
            {
                Error("二本线不应该低于三本线吧？");
                return;
            }

            if (CheckGridView(gk_gridview))
                return;

            Analysis analysis = new Analysis(this);
            analysis._gridview = gk_gridview;
            analysis.save_address = gk_save_address.Text;
            analysis.isVisible = gk_isVisible.Checked;
            analysis.CurrentDirectory = currentdic;
            analysis.cj_addr = gk_xz_cj_addr.Text.ToString().Trim();
            analysis.curryear = year_list.SelectedItem.ToString().Trim();
            analysis.currmonth = currmonth.SelectedItem.ToString().Trim();

            analysis.wen_first_level = Convert.ToInt32(wen_first_level.Value);
            analysis.wen_second_level = Convert.ToInt32(wen_second_level.Value);
            analysis.wen_third_level = Convert.ToInt32(wen_third_level.Value);

            analysis.li_first_level = Convert.ToInt32(li_first_level.Value);
            analysis.li_second_level = Convert.ToInt32(li_second_level.Value);
            analysis.li_third_level = Convert.ToInt32(li_third_level.Value);

            Thread thread = new Thread(new ThreadStart(analysis.gk_xz_start));
            thread.IsBackground = true;
            thread.SetApartmentState(ApartmentState.STA);
            thread_store.Add("gk_xz", thread);
            thread.Start();
        }

        private void gk_xz_cancel_Click(object sender, EventArgs e)
        {
            if (thread_store.ContainsKey("gk_xz"))
            {
                Thread thread = thread_store["gk_xz"];
                if (thread.IsAlive)
                {
                    thread.Abort();
                    thread_store.Remove("gk_xz");
                    ShowPro("gk_xz", 2, "");
                }
            }
        }

        private void zk_xz_start_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(zk_xz_qx_addr.Text.Trim()))
            {
                Error("请输入城郊分类文件地址！");
                return;
            }

            if (zk_first_level.Value == 0 || zk_second_level.Value == 0)
            {
                Error("重点线不能为零！");
                return;
            }
            if (zk_first_level.Value < zk_second_level.Value)
            {
                Error("重点线不应该低于普通线吧？");
                return;
            }
            

            if (CheckGridView(zk_gridview))
                return;

            Analysis analysis = new Analysis(this);
            analysis._gridview = zk_gridview;
            analysis.save_address = save_address.Text;
            analysis.isVisible = zk_isVisible.Checked;
            analysis.CurrentDirectory = currentdic;
            analysis.cj_addr = zk_xz_qx_addr.Text.ToString().Trim();
            analysis.curryear = zk_yearlist.SelectedItem.ToString().Trim();
            analysis.currmonth = zk_currmonth.SelectedItem.ToString().Trim();

            analysis.zk_first_level = Convert.ToInt32(zk_first_level.Value);
            analysis.zk_second_level = Convert.ToInt32(zk_second_level.Value);

            Thread thread = new Thread(new ThreadStart(analysis.zk_xz_start));
            thread.IsBackground = true;
            thread.SetApartmentState(ApartmentState.STA);
            thread_store.Add("zk_xz", thread);
            thread.Start();
        }

        private void zk_xz_cancel_Click(object sender, EventArgs e)
        {
            if (thread_store.ContainsKey("zk_xz"))
            {
                Thread thread = thread_store["zk_xz"];
                if (thread.IsAlive)
                {
                    thread.Abort();
                    thread_store.Remove("zk_xz");
                    ShowPro("zk_xz", 2, "");
                }
            }
        }

        private void radButton15_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.InitialDirectory = "C://";
            openFileDialog1.Filter = "Excel files (*.xls,*.xlsx)|*.xls;*.xlsx|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.RestoreDirectory = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
                zk_xz_qx_addr.Text = openFileDialog1.FileName;
        }

        private void compare_total_start_Click(object sender, EventArgs e)
        {
            Analysis analysis = new Analysis(this);
            analysis.compare_year1 = compare_year1.SelectedItem.ToString().Trim();
            analysis.compare_year2 = compare_year2.SelectedItem.ToString().Trim();

            Thread thread = new Thread(new ThreadStart(analysis.gk_comp_start));
            thread.IsBackground = true;
            thread.SetApartmentState(ApartmentState.STA);
            thread_store.Add("gk_cp", thread);
            thread.Start();

        }

        private void compare_total_cancel_Click(object sender, EventArgs e)
        {
            if (thread_store.ContainsKey("gk_cp"))
            {
                Thread thread = thread_store["gk_cp"];
                if (thread.IsAlive)
                {
                    thread.Abort();
                    thread_store.Remove("gk_cp");
                    ShowPro("gk_cp", 2, "");
                }
            }
        }

        private void radButton17_Click(object sender, EventArgs e)
        {
            Analysis analysis = new Analysis(this);
            analysis.compare_cj_is_wk = wk_radio.IsChecked;
            analysis.cj_comp_year = cj_comp_year.SelectedItem.ToString().Trim();

            Thread thread = new Thread(new ThreadStart(analysis.gk_cj_comp_start));
            thread.IsBackground = true;
            thread.SetApartmentState(ApartmentState.STA);
            thread_store.Add("gk_cj_cp", thread);
            thread.Start();
        }

        private void cj_comp_cancel_Click(object sender, EventArgs e)
        {
            if (thread_store.ContainsKey("gk_cj_cp"))
            {
                Thread thread = thread_store["gk_cj_cp"];
                if (thread.IsAlive)
                {
                    thread.Abort();
                    thread_store.Remove("gk_cj_cp");
                    ShowPro("gk_cj_cp", 2, "");
                }
            }
        }

       
    }
}
