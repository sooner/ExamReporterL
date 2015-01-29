namespace ExamReport
{
    partial class MyWizard
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MyWizard));
            this.radWizard1 = new Telerik.WinControls.UI.RadWizard();
            this.wizardCompletionPage1 = new Telerik.WinControls.UI.WizardCompletionPage();
            this.panel3 = new System.Windows.Forms.Panel();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.panel1 = new System.Windows.Forms.Panel();
            this.is_batch_import = new System.Windows.Forms.CheckBox();
            this.exam_date = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.exam = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.panel5 = new System.Windows.Forms.Panel();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.button3 = new System.Windows.Forms.Button();
            this.group_addr = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.button2 = new System.Windows.Forms.Button();
            this.ans_addr = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.database_addr = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.subject = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.wizardWelcomePage1 = new Telerik.WinControls.UI.WizardWelcomePage();
            this.wizardPage3 = new Telerik.WinControls.UI.WizardPage();
            this.label7 = new System.Windows.Forms.Label();
            this.fullmark = new System.Windows.Forms.NumericUpDown();
            this.zf_panel = new System.Windows.Forms.Panel();
            this.zh_panel = new System.Windows.Forms.Panel();
            this.huaxue_dili = new System.Windows.Forms.NumericUpDown();
            this.label43 = new System.Windows.Forms.Label();
            this.wuli_lishi = new System.Windows.Forms.NumericUpDown();
            this.label42 = new System.Windows.Forms.Label();
            this.shengwu_zhengzhi = new System.Windows.Forms.NumericUpDown();
            this.label41 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label44 = new System.Windows.Forms.Label();
            this.PartialRight = new System.Windows.Forms.NumericUpDown();
            this.label33 = new System.Windows.Forms.Label();
            this.fullmark_iszero = new System.Windows.Forms.CheckBox();
            this.sub_iszero = new System.Windows.Forms.CheckBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.label19 = new System.Windows.Forms.Label();
            this.remark_num = new System.Windows.Forms.NumericUpDown();
            this.label18 = new System.Windows.Forms.Label();
            this.label17 = new System.Windows.Forms.Label();
            this.popu_num = new System.Windows.Forms.NumericUpDown();
            this.label16 = new System.Windows.Forms.Label();
            this.Mark_choice = new System.Windows.Forms.RadioButton();
            this.Popu_choice = new System.Windows.Forms.RadioButton();
            ((System.ComponentModel.ISupportInitialize)(this.radWizard1)).BeginInit();
            this.radWizard1.SuspendLayout();
            this.panel3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.panel1.SuspendLayout();
            this.panel5.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.fullmark)).BeginInit();
            this.zf_panel.SuspendLayout();
            this.zh_panel.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.huaxue_dili)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.wuli_lishi)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.shengwu_zhengzhi)).BeginInit();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.PartialRight)).BeginInit();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.remark_num)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.popu_num)).BeginInit();
            this.SuspendLayout();
            // 
            // radWizard1
            // 
            this.radWizard1.CompletionPage = this.wizardCompletionPage1;
            this.radWizard1.Controls.Add(this.panel1);
            this.radWizard1.Controls.Add(this.panel3);
            this.radWizard1.Controls.Add(this.panel5);
            this.radWizard1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.radWizard1.Location = new System.Drawing.Point(0, 0);
            this.radWizard1.Name = "radWizard1";
            this.radWizard1.PageHeaderIcon = null;
            this.radWizard1.Pages.Add(this.wizardWelcomePage1);
            this.radWizard1.Pages.Add(this.wizardPage3);
            this.radWizard1.Pages.Add(this.wizardCompletionPage1);
            this.radWizard1.Size = new System.Drawing.Size(642, 524);
            this.radWizard1.TabIndex = 0;
            this.radWizard1.ThemeName = "ControlDefault";
            this.radWizard1.WelcomePage = this.wizardWelcomePage1;
            // 
            // wizardCompletionPage1
            // 
            this.wizardCompletionPage1.ContentArea = this.panel3;
            this.wizardCompletionPage1.Header = "";
            this.wizardCompletionPage1.Name = "wizardCompletionPage1";
            this.wizardCompletionPage1.Title = "导入数据";
            this.wizardCompletionPage1.Visibility = Telerik.WinControls.ElementVisibility.Visible;
            // 
            // panel3
            // 
            this.panel3.BackColor = System.Drawing.Color.White;
            this.panel3.Controls.Add(this.dataGridView1);
            this.panel3.Location = new System.Drawing.Point(150, 56);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(471, 418);
            this.panel3.TabIndex = 2;
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(2, 3);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 23;
            this.dataGridView1.Size = new System.Drawing.Size(468, 412);
            this.dataGridView1.TabIndex = 0;
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.White;
            this.panel1.Controls.Add(this.is_batch_import);
            this.panel1.Controls.Add(this.exam_date);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.exam);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Location = new System.Drawing.Point(150, 56);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(492, 420);
            this.panel1.TabIndex = 0;
            // 
            // is_batch_import
            // 
            this.is_batch_import.AutoSize = true;
            this.is_batch_import.Location = new System.Drawing.Point(33, 149);
            this.is_batch_import.Name = "is_batch_import";
            this.is_batch_import.Size = new System.Drawing.Size(72, 16);
            this.is_batch_import.TabIndex = 4;
            this.is_batch_import.Text = "批量录入";
            this.is_batch_import.UseVisualStyleBackColor = true;
            // 
            // exam_date
            // 
            this.exam_date.FormattingEnabled = true;
            this.exam_date.Location = new System.Drawing.Point(99, 109);
            this.exam_date.Name = "exam_date";
            this.exam_date.Size = new System.Drawing.Size(121, 20);
            this.exam_date.TabIndex = 3;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(31, 111);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(65, 12);
            this.label2.TabIndex = 2;
            this.label2.Text = "考试日期：";
            // 
            // exam
            // 
            this.exam.FormattingEnabled = true;
            this.exam.Items.AddRange(new object[] {
            "中考",
            "会考",
            "高考"});
            this.exam.Location = new System.Drawing.Point(99, 66);
            this.exam.Name = "exam";
            this.exam.Size = new System.Drawing.Size(121, 20);
            this.exam.TabIndex = 1;
            this.exam.SelectedIndexChanged += new System.EventHandler(this.exam_type_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(29, 70);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(65, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "考试类型：";
            // 
            // panel5
            // 
            this.panel5.BackColor = System.Drawing.Color.White;
            this.panel5.Controls.Add(this.zf_panel);
            this.panel5.Controls.Add(this.groupBox2);
            this.panel5.Controls.Add(this.groupBox1);
            this.panel5.Controls.Add(this.zh_panel);
            this.panel5.Controls.Add(this.progressBar);
            this.panel5.Controls.Add(this.button3);
            this.panel5.Controls.Add(this.group_addr);
            this.panel5.Controls.Add(this.label6);
            this.panel5.Controls.Add(this.button2);
            this.panel5.Controls.Add(this.ans_addr);
            this.panel5.Controls.Add(this.label5);
            this.panel5.Controls.Add(this.button1);
            this.panel5.Controls.Add(this.database_addr);
            this.panel5.Controls.Add(this.label4);
            this.panel5.Controls.Add(this.subject);
            this.panel5.Controls.Add(this.label3);
            this.panel5.Location = new System.Drawing.Point(0, 64);
            this.panel5.Name = "panel5";
            this.panel5.Size = new System.Drawing.Size(642, 412);
            this.panel5.TabIndex = 4;
            // 
            // progressBar
            // 
            this.progressBar.Location = new System.Drawing.Point(26, 380);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(580, 17);
            this.progressBar.TabIndex = 11;
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(534, 139);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(75, 23);
            this.button3.TabIndex = 10;
            this.button3.Text = "打开";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // group_addr
            // 
            this.group_addr.Location = new System.Drawing.Point(90, 141);
            this.group_addr.Name = "group_addr";
            this.group_addr.Size = new System.Drawing.Size(423, 21);
            this.group_addr.TabIndex = 9;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(29, 144);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(65, 12);
            this.label6.TabIndex = 8;
            this.label6.Text = "数据需求：";
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(534, 105);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 23);
            this.button2.TabIndex = 7;
            this.button2.Text = "打开";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // ans_addr
            // 
            this.ans_addr.Location = new System.Drawing.Point(90, 105);
            this.ans_addr.Name = "ans_addr";
            this.ans_addr.Size = new System.Drawing.Size(423, 21);
            this.ans_addr.TabIndex = 6;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(29, 108);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(65, 12);
            this.label5.TabIndex = 5;
            this.label5.Text = "标准答案：";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(534, 68);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 4;
            this.button1.Text = "打开";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // database_addr
            // 
            this.database_addr.Location = new System.Drawing.Point(90, 68);
            this.database_addr.Name = "database_addr";
            this.database_addr.Size = new System.Drawing.Size(423, 21);
            this.database_addr.TabIndex = 3;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(29, 71);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(65, 12);
            this.label4.TabIndex = 2;
            this.label4.Text = "数据文件：";
            // 
            // subject
            // 
            this.subject.FormattingEnabled = true;
            this.subject.Location = new System.Drawing.Point(90, 26);
            this.subject.Name = "subject";
            this.subject.Size = new System.Drawing.Size(121, 20);
            this.subject.TabIndex = 1;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(28, 30);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(65, 12);
            this.label3.TabIndex = 0;
            this.label3.Text = "考试科目：";
            // 
            // wizardWelcomePage1
            // 
            this.wizardWelcomePage1.ContentArea = this.panel1;
            this.wizardWelcomePage1.Header = "";
            this.wizardWelcomePage1.Name = "wizardWelcomePage1";
            this.wizardWelcomePage1.Title = "考试类型";
            this.wizardWelcomePage1.Visibility = Telerik.WinControls.ElementVisibility.Visible;
            // 
            // wizardPage3
            // 
            this.wizardPage3.ContentArea = this.panel5;
            this.wizardPage3.Header = "";
            this.wizardPage3.Name = "wizardPage3";
            this.wizardPage3.Title = "科目信息";
            this.wizardPage3.Visibility = Telerik.WinControls.ElementVisibility.Visible;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(3, 10);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(65, 12);
            this.label7.TabIndex = 12;
            this.label7.Text = "科目总分：";
            // 
            // fullmark
            // 
            this.fullmark.Location = new System.Drawing.Point(64, 7);
            this.fullmark.Maximum = new decimal(new int[] {
            800,
            0,
            0,
            0});
            this.fullmark.Name = "fullmark";
            this.fullmark.Size = new System.Drawing.Size(120, 21);
            this.fullmark.TabIndex = 13;
            // 
            // zf_panel
            // 
            this.zf_panel.Controls.Add(this.fullmark);
            this.zf_panel.Controls.Add(this.label7);
            this.zf_panel.Location = new System.Drawing.Point(31, 314);
            this.zf_panel.Name = "zf_panel";
            this.zf_panel.Size = new System.Drawing.Size(245, 39);
            this.zf_panel.TabIndex = 14;
            // 
            // zh_panel
            // 
            this.zh_panel.Controls.Add(this.huaxue_dili);
            this.zh_panel.Controls.Add(this.label43);
            this.zh_panel.Controls.Add(this.wuli_lishi);
            this.zh_panel.Controls.Add(this.label42);
            this.zh_panel.Controls.Add(this.shengwu_zhengzhi);
            this.zh_panel.Controls.Add(this.label41);
            this.zh_panel.Location = new System.Drawing.Point(297, 314);
            this.zh_panel.Name = "zh_panel";
            this.zh_panel.Size = new System.Drawing.Size(312, 39);
            this.zh_panel.TabIndex = 15;
            // 
            // huaxue_dili
            // 
            this.huaxue_dili.Location = new System.Drawing.Point(254, 9);
            this.huaxue_dili.Maximum = new decimal(new int[] {
            500,
            0,
            0,
            0});
            this.huaxue_dili.Name = "huaxue_dili";
            this.huaxue_dili.Size = new System.Drawing.Size(52, 21);
            this.huaxue_dili.TabIndex = 11;
            // 
            // label43
            // 
            this.label43.AutoSize = true;
            this.label43.Location = new System.Drawing.Point(214, 13);
            this.label43.Name = "label43";
            this.label43.Size = new System.Drawing.Size(41, 12);
            this.label43.TabIndex = 10;
            this.label43.Text = "化学：";
            // 
            // wuli_lishi
            // 
            this.wuli_lishi.Location = new System.Drawing.Point(152, 9);
            this.wuli_lishi.Maximum = new decimal(new int[] {
            500,
            0,
            0,
            0});
            this.wuli_lishi.Name = "wuli_lishi";
            this.wuli_lishi.Size = new System.Drawing.Size(54, 21);
            this.wuli_lishi.TabIndex = 9;
            // 
            // label42
            // 
            this.label42.AutoSize = true;
            this.label42.Location = new System.Drawing.Point(115, 13);
            this.label42.Name = "label42";
            this.label42.Size = new System.Drawing.Size(41, 12);
            this.label42.TabIndex = 8;
            this.label42.Text = "物理：";
            // 
            // shengwu_zhengzhi
            // 
            this.shengwu_zhengzhi.Location = new System.Drawing.Point(51, 9);
            this.shengwu_zhengzhi.Maximum = new decimal(new int[] {
            500,
            0,
            0,
            0});
            this.shengwu_zhengzhi.Name = "shengwu_zhengzhi";
            this.shengwu_zhengzhi.Size = new System.Drawing.Size(55, 21);
            this.shengwu_zhengzhi.TabIndex = 7;
            // 
            // label41
            // 
            this.label41.AutoSize = true;
            this.label41.Location = new System.Drawing.Point(13, 13);
            this.label41.Name = "label41";
            this.label41.Size = new System.Drawing.Size(41, 12);
            this.label41.TabIndex = 6;
            this.label41.Text = "生物：";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label44);
            this.groupBox1.Controls.Add(this.PartialRight);
            this.groupBox1.Controls.Add(this.label33);
            this.groupBox1.Controls.Add(this.fullmark_iszero);
            this.groupBox1.Controls.Add(this.sub_iszero);
            this.groupBox1.Location = new System.Drawing.Point(31, 192);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(245, 96);
            this.groupBox1.TabIndex = 72;
            this.groupBox1.TabStop = false;
            // 
            // label44
            // 
            this.label44.AutoSize = true;
            this.label44.Location = new System.Drawing.Point(181, 64);
            this.label44.Name = "label44";
            this.label44.Size = new System.Drawing.Size(17, 12);
            this.label44.TabIndex = 74;
            this.label44.Text = "分";
            // 
            // PartialRight
            // 
            this.PartialRight.Location = new System.Drawing.Point(136, 59);
            this.PartialRight.Name = "PartialRight";
            this.PartialRight.Size = new System.Drawing.Size(37, 21);
            this.PartialRight.TabIndex = 73;
            // 
            // label33
            // 
            this.label33.AutoSize = true;
            this.label33.Location = new System.Drawing.Point(55, 64);
            this.label33.Name = "label33";
            this.label33.Size = new System.Drawing.Size(77, 12);
            this.label33.TabIndex = 72;
            this.label33.Text = "多选题少选得";
            // 
            // fullmark_iszero
            // 
            this.fullmark_iszero.AutoSize = true;
            this.fullmark_iszero.Location = new System.Drawing.Point(16, 26);
            this.fullmark_iszero.Name = "fullmark_iszero";
            this.fullmark_iszero.Size = new System.Drawing.Size(96, 16);
            this.fullmark_iszero.TabIndex = 69;
            this.fullmark_iszero.Text = "删除总分为零";
            this.fullmark_iszero.UseVisualStyleBackColor = true;
            // 
            // sub_iszero
            // 
            this.sub_iszero.AutoSize = true;
            this.sub_iszero.Location = new System.Drawing.Point(122, 25);
            this.sub_iszero.Name = "sub_iszero";
            this.sub_iszero.Size = new System.Drawing.Size(108, 16);
            this.sub_iszero.TabIndex = 68;
            this.sub_iszero.Text = "删除主观题为零";
            this.sub_iszero.UseVisualStyleBackColor = true;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.label19);
            this.groupBox2.Controls.Add(this.remark_num);
            this.groupBox2.Controls.Add(this.label18);
            this.groupBox2.Controls.Add(this.label17);
            this.groupBox2.Controls.Add(this.popu_num);
            this.groupBox2.Controls.Add(this.label16);
            this.groupBox2.Controls.Add(this.Mark_choice);
            this.groupBox2.Controls.Add(this.Popu_choice);
            this.groupBox2.Location = new System.Drawing.Point(297, 192);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(309, 96);
            this.groupBox2.TabIndex = 75;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "分组类型";
            // 
            // label19
            // 
            this.label19.AutoSize = true;
            this.label19.Location = new System.Drawing.Point(217, 61);
            this.label19.Name = "label19";
            this.label19.Size = new System.Drawing.Size(41, 12);
            this.label19.TabIndex = 7;
            this.label19.Text = "分一组";
            // 
            // remark_num
            // 
            this.remark_num.Location = new System.Drawing.Point(129, 58);
            this.remark_num.Name = "remark_num";
            this.remark_num.Size = new System.Drawing.Size(77, 21);
            this.remark_num.TabIndex = 6;
            // 
            // label18
            // 
            this.label18.AutoSize = true;
            this.label18.Location = new System.Drawing.Point(107, 61);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(17, 12);
            this.label18.TabIndex = 5;
            this.label18.Text = "每";
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.Location = new System.Drawing.Point(215, 25);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(17, 12);
            this.label17.TabIndex = 4;
            this.label17.Text = "组";
            // 
            // popu_num
            // 
            this.popu_num.Location = new System.Drawing.Point(129, 21);
            this.popu_num.Name = "popu_num";
            this.popu_num.Size = new System.Drawing.Size(77, 21);
            this.popu_num.TabIndex = 3;
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Location = new System.Drawing.Point(99, 25);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(29, 12);
            this.label16.TabIndex = 2;
            this.label16.Text = "平分";
            // 
            // Mark_choice
            // 
            this.Mark_choice.AutoSize = true;
            this.Mark_choice.Location = new System.Drawing.Point(17, 61);
            this.Mark_choice.Name = "Mark_choice";
            this.Mark_choice.Size = new System.Drawing.Size(83, 16);
            this.Mark_choice.TabIndex = 1;
            this.Mark_choice.TabStop = true;
            this.Mark_choice.Text = "按成绩分：";
            this.Mark_choice.UseVisualStyleBackColor = true;
            // 
            // Popu_choice
            // 
            this.Popu_choice.AutoSize = true;
            this.Popu_choice.Location = new System.Drawing.Point(17, 23);
            this.Popu_choice.Name = "Popu_choice";
            this.Popu_choice.Size = new System.Drawing.Size(83, 16);
            this.Popu_choice.TabIndex = 0;
            this.Popu_choice.TabStop = true;
            this.Popu_choice.Text = "按人数分：";
            this.Popu_choice.UseVisualStyleBackColor = true;
            // 
            // MyWizard
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(642, 524);
            this.Controls.Add(this.radWizard1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "MyWizard";
            this.Text = "数据导入";
            ((System.ComponentModel.ISupportInitialize)(this.radWizard1)).EndInit();
            this.radWizard1.ResumeLayout(false);
            this.panel3.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel5.ResumeLayout(false);
            this.panel5.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.fullmark)).EndInit();
            this.zf_panel.ResumeLayout(false);
            this.zf_panel.PerformLayout();
            this.zh_panel.ResumeLayout(false);
            this.zh_panel.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.huaxue_dili)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.wuli_lishi)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.shengwu_zhengzhi)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.PartialRight)).EndInit();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.remark_num)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.popu_num)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private Telerik.WinControls.UI.RadWizard radWizard1;
        private Telerik.WinControls.UI.WizardCompletionPage wizardCompletionPage1;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Panel panel1;
        private Telerik.WinControls.UI.WizardWelcomePage wizardWelcomePage1;
        private System.Windows.Forms.CheckBox is_batch_import;
        private System.Windows.Forms.ComboBox exam_date;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox exam;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Panel panel5;
        private System.Windows.Forms.Label label3;
        private Telerik.WinControls.UI.WizardPage wizardPage3;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.TextBox group_addr;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.TextBox ans_addr;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox database_addr;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ComboBox subject;
        private System.Windows.Forms.NumericUpDown fullmark;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Panel zf_panel;
        private System.Windows.Forms.Panel zh_panel;
        private System.Windows.Forms.NumericUpDown huaxue_dili;
        private System.Windows.Forms.Label label43;
        private System.Windows.Forms.NumericUpDown wuli_lishi;
        private System.Windows.Forms.Label label42;
        private System.Windows.Forms.NumericUpDown shengwu_zhengzhi;
        private System.Windows.Forms.Label label41;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label44;
        private System.Windows.Forms.NumericUpDown PartialRight;
        private System.Windows.Forms.Label label33;
        private System.Windows.Forms.CheckBox fullmark_iszero;
        private System.Windows.Forms.CheckBox sub_iszero;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Label label19;
        private System.Windows.Forms.NumericUpDown remark_num;
        private System.Windows.Forms.Label label18;
        private System.Windows.Forms.Label label17;
        private System.Windows.Forms.NumericUpDown popu_num;
        private System.Windows.Forms.Label label16;
        private System.Windows.Forms.RadioButton Mark_choice;
        private System.Windows.Forms.RadioButton Popu_choice;
    }
}