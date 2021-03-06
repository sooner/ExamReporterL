﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Telerik.WinControls.UI;
using System.Threading;

namespace ExamReport
{

    //public delegate void MyDelegate(int i, int status);
    public delegate void CheckDataMethod(int status, string message);
    public partial class MyWizard : Form
    {
        Thread thread;
        public LoadDatabase ld;
        public mainform _form;
        public MyWizard(string sub, mainform form)
        {
            InitializeComponent();
            _form = form;
            exam.SelectedItem = sub;
            int curryear = DateTime.Now.Year;
            for (int i = curryear - 10; i < curryear + 10; i++)
                exam_date.Items.Add(i);
            exam_date.SelectedItem = curryear;

            radWizard1.Next += new WizardCancelEventHandler(radWizard_Next);
            radWizard1.Cancel += new EventHandler(radWizard1_Cancel);
            radWizard1.Finish += new EventHandler(radWizard_Finish);

            zf_panel.Show();
            zh_panel.Hide();
            zh_panel2.Hide();
        }

        void radWizard1_Cancel(object sender, EventArgs e)
        {
            if (this.radWizard1.SelectedPage == this.radWizard1.Pages[1])
            {
                if (thread.IsAlive)
                {
                    thread.Abort();
                    ShowPro(100, 2);
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.InitialDirectory = "C://";
            openFileDialog1.Filter = "DBF files (*.dbf)|*.dbf|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.RestoreDirectory = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
                database_addr.Text = openFileDialog1.FileName;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.InitialDirectory = "C://";
            openFileDialog1.Filter = "Excel files (*.xls,*.xlsx)|*.xls;*.xlsx|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.RestoreDirectory = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
                ans_addr.Text = openFileDialog1.FileName;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.InitialDirectory = "C://";
            openFileDialog1.Filter = "Excel files (*.xls,*.xlsx)|*.xls;*.xlsx|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.RestoreDirectory = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
                group_addr.Text = openFileDialog1.FileName;
        }
        private void radWizard_Next(object sender, WizardCancelEventArgs e)
        {
            if (this.radWizard1.SelectedPage == this.radWizard1.Pages[1])
            {
                e.Cancel = true;
                radWizard1.NextButton.Enabled = false;
                radWizard1.BackButton.Enabled = false;
                //this.radWizard1.SelectedPage = this.radWizard1.Pages[0];
                start_process();
            }
            

        }

        public void radWizard_Finish(object sender, EventArgs e)
        {
            _form.Grid_load();
            this.Close();
        }

        

        private void start_process()
        {
            ld = new LoadDatabase();
            ld.wizard = this;
            ld.year = exam_date.SelectedItem.ToString();
            ld.exam = exam.SelectedItem.ToString();
            ld.sub = subject.SelectedItem.ToString();
            ld.database_str = database_addr.Text;
            if (!(ld.sub.Equals("总分") || ld.sub.Contains("行政版")))
            {
                ld.ans_str = ans_addr.Text;
                ld.group_str = group_addr.Text;


                if (ld.sub.Contains("理综") || ld.sub.Contains("文综"))
                {


                    if (string.IsNullOrEmpty(zh_addr.Text.Trim()))
                    {
                        Error("请输入综合分类文件地址！");
                        return;
                    }
                    if (Math.Abs(single_fullmark.Value) != single_fullmark.Value)
                    {
                        Error("单科总分不能为负！");
                        return;
                    }
                    if (single_fullmark.Value > fullmark.Value)
                    {
                        Error("单科成绩不能大于总成绩！");
                        return;
                    }
                    ld.wenli_str = zh_addr.Text;

                    ld.sub_fullmark = single_fullmark.Value;

                }
                if (Popu_choice.Checked)
                {
                    ld.grouptype = ZK_database.GroupType.population;
                    ld.divider = popu_num.Value;
                }
                else if (Mark_choice.Checked)
                {

                    ld.grouptype = ZK_database.GroupType.totalmark;
                    ld.divider = remark_num.Value;

                }
                else
                {
                    Error("必须选择一种分类方式");
                }
                if (ld.divider <= 0)
                    Error("分组数必须大于0");
            }
            ld.fullmark = fullmark.Value;
            ld.PartialRight = PartialRight.Value;
            if (sub_iszero.Checked)
                ld.sub_iszero = true;
            else
                ld.sub_iszero = false;
            if (fullmark_iszero.Checked)
                ld.fullmark_iszero = true;
            else
                ld.fullmark_iszero = false;

            thread = new Thread(new ThreadStart(ld.start_process));
            thread.IsBackground = true;
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
        }
        public void ShowPro(int value, int status)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new MyDelegate(ShowPro), value, status);
            }
            else
            {
                this.progressBar.Value = value;
                switch (status)
                {
                    case 0:
                        break;
                    case 1:
                        break;
                    case 2:
                        radWizard1.NextButton.Enabled = true;
                        radWizard1.BackButton.Enabled = true;
                        break;
                    case 3:
                        after_process();
                        break;
                    default:
                        break;
                }
            }
        }
        public void CheckData(int status, string message)
        {
            if (this.InvokeRequired)
                this.Invoke(new CheckDataMethod(CheckData), status, message);
            else
            {
                switch (status)
                {
                    case 1:

                        MessageBoxButtons messButton = MessageBoxButtons.OKCancel;
                        DialogResult dr = MessageBox.Show(message + "\n仍然继续嘛？", "是否继续", messButton);
                        if (dr == DialogResult.Cancel)
                        {
                            thread.Abort();
                            ShowPro(100, 2);
                        }
                        break;
                    
                    default:
                        break;
                }
            }
        }
        private void after_process()
        {
            if (subject.SelectedItem.ToString().Equals("总分") || subject.SelectedItem.ToString().Contains("行政版"))
            {
                basic_gridView.DataSource = ld.basic_data;
                this.radWizard1.SelectedPage = this.radWizard1.Pages[2];
            }
            else
            {
                basic_gridView.DataSource = ld.basic_data;
                group_gridView.DataSource = ld.group_data;
                this.radWizard1.SelectedPage = this.radWizard1.Pages[2];
            }
        }
        private void exam_type_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (exam.SelectedIndex)
            {
                case 0:
                    
                    subject.Items.Clear();
                    subject.Items.Add("语文");
                    subject.Items.Add("数学");
                    subject.Items.Add("英语");
                    subject.Items.Add("物理");
                    subject.Items.Add("化学");
                    subject.Items.Add("生物");
                    subject.Items.Add("地理");
                    subject.Items.Add("历史");
                    subject.Items.Add("思想品德");
                    subject.Items.Add("中考行政版");
                    subject.Items.Add("总分");
                    subject.ResetText();


                    break;
                case 1:
                    
                    subject.Items.Clear();
                    subject.Items.Add("语文");
                    subject.Items.Add("数学");
                    subject.Items.Add("英语");
                    subject.Items.Add("物理");
                    subject.Items.Add("化学");
                    subject.Items.Add("生物");
                    subject.Items.Add("政治");
                    subject.Items.Add("历史");
                    subject.Items.Add("地理");
                    subject.ResetText();

                    break;
                case 2:

                    subject.Items.Clear();
                    subject.Items.Add("总分");
                    subject.Items.Add("高考行政版");
                    subject.Items.Add("语文");
                    subject.Items.Add("英语");
                    subject.Items.Add("数学理");
                    subject.Items.Add("数学文");
                    subject.Items.Add("理综-物理");
                    subject.Items.Add("理综-化学");
                    subject.Items.Add("理综-生物");
                    subject.Items.Add("文综-政治");
                    subject.Items.Add("文综-地理");
                    subject.Items.Add("文综-历史");
                    subject.ResetText();

                    break;
                case 3:
                    subject.Items.Clear();
                    subject.Items.Add("总分");
                    subject.Items.Add("高考行政版");
                    subject.Items.Add("语文");
                    subject.Items.Add("数学");
                    subject.Items.Add("英语");
                    subject.Items.Add("物理");
                    subject.Items.Add("化学");
                    subject.Items.Add("生物");
                    subject.Items.Add("政治");
                    subject.Items.Add("历史");
                    subject.Items.Add("地理");
                    subject.ResetText();
                    break;
                default:
                    break;

            }
        }
        public void ErrorM(string message)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new ErrorMessage(ErrorM), message);
            }
            else
            {
                Error(message);
                if (thread.IsAlive)
                {
                    thread.Abort();
                    ShowPro(100, 2);
                }

            }
        }

        private bool Error(string errormessage)
        {
            MessageBox.Show(errormessage, "提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return false;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.InitialDirectory = "C://";
            openFileDialog1.Filter = "Excel files (*.xls,*.xlsx)|*.xls;*.xlsx|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.RestoreDirectory = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
                zh_addr.Text = openFileDialog1.FileName;
        }

        private void subject_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(subject.SelectedItem.ToString().Contains("理综") || subject.SelectedItem.ToString().Contains("文综"))
            {
                zh_panel2.Show();
                zh_panel.Show();
                zongfen_enable();
            }
            else if (subject.SelectedItem.ToString().Equals("总分") || subject.SelectedItem.ToString().Contains("行政版"))
            {
                zongfen_disable();
                zh_panel.Hide();
                zh_panel2.Hide();
            }
            else
            {
                zh_panel.Hide();
                zh_panel2.Hide();
                zongfen_enable();
            }
        }

        private void zongfen_disable()
        {
            ans_addr.Enabled = false;
            group_addr.Enabled = false;
            button2.Enabled = false;
            button3.Enabled = false;
            popu_num.Enabled = false;
            remark_num.Enabled = false;

        }
        private void zongfen_enable()
        {
            ans_addr.Enabled = true;
            group_addr.Enabled = true;
            button2.Enabled = true;
            button3.Enabled = true;
            popu_num.Enabled = true;
            remark_num.Enabled = true;

        }
    
    }
}
