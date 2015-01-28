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
            this.panel1 = new System.Windows.Forms.Panel();
            this.is_batch_import = new System.Windows.Forms.CheckBox();
            this.exam_date = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.exam_type = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.panel5 = new System.Windows.Forms.Panel();
            this.label3 = new System.Windows.Forms.Label();
            this.wizardWelcomePage1 = new Telerik.WinControls.UI.WizardWelcomePage();
            this.wizardPage3 = new Telerik.WinControls.UI.WizardPage();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.button3 = new System.Windows.Forms.Button();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.label7 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.radWizard1)).BeginInit();
            this.radWizard1.SuspendLayout();
            this.panel3.SuspendLayout();
            this.panel1.SuspendLayout();
            this.panel5.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
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
            this.radWizard1.Mode = Telerik.WinControls.UI.WizardMode.Aero;
            this.radWizard1.Name = "radWizard1";
            this.radWizard1.PageHeaderIcon = ((System.Drawing.Image)(resources.GetObject("radWizard1.PageHeaderIcon")));
            this.radWizard1.Pages.Add(this.wizardWelcomePage1);
            this.radWizard1.Pages.Add(this.wizardPage3);
            this.radWizard1.Pages.Add(this.wizardCompletionPage1);
            this.radWizard1.Size = new System.Drawing.Size(587, 440);
            this.radWizard1.TabIndex = 0;
            this.radWizard1.Text = "radWizard1";
            this.radWizard1.WelcomePage = this.wizardWelcomePage1;
            // 
            // wizardCompletionPage1
            // 
            this.wizardCompletionPage1.ContentArea = this.panel3;
            this.wizardCompletionPage1.Name = "wizardCompletionPage1";
            this.wizardCompletionPage1.Visibility = Telerik.WinControls.ElementVisibility.Visible;
            // 
            // panel3
            // 
            this.panel3.BackColor = System.Drawing.Color.White;
            this.panel3.Controls.Add(this.label7);
            this.panel3.Controls.Add(this.dataGridView1);
            this.panel3.Location = new System.Drawing.Point(150, 41);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(437, 351);
            this.panel3.TabIndex = 2;
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.White;
            this.panel1.Controls.Add(this.is_batch_import);
            this.panel1.Controls.Add(this.exam_date);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.exam_type);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Location = new System.Drawing.Point(150, 41);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(437, 351);
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
            // exam_type
            // 
            this.exam_type.FormattingEnabled = true;
            this.exam_type.Items.AddRange(new object[] {
            "中考",
            "会考",
            "高考"});
            this.exam_type.Location = new System.Drawing.Point(99, 66);
            this.exam_type.Name = "exam_type";
            this.exam_type.Size = new System.Drawing.Size(121, 20);
            this.exam_type.TabIndex = 1;
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
            this.panel5.Controls.Add(this.progressBar1);
            this.panel5.Controls.Add(this.button3);
            this.panel5.Controls.Add(this.textBox3);
            this.panel5.Controls.Add(this.label6);
            this.panel5.Controls.Add(this.button2);
            this.panel5.Controls.Add(this.textBox2);
            this.panel5.Controls.Add(this.label5);
            this.panel5.Controls.Add(this.button1);
            this.panel5.Controls.Add(this.textBox1);
            this.panel5.Controls.Add(this.label4);
            this.panel5.Controls.Add(this.comboBox1);
            this.panel5.Controls.Add(this.label3);
            this.panel5.Location = new System.Drawing.Point(0, 41);
            this.panel5.Name = "panel5";
            this.panel5.Size = new System.Drawing.Size(587, 351);
            this.panel5.TabIndex = 4;
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
            this.wizardWelcomePage1.Name = "wizardWelcomePage1";
            this.wizardWelcomePage1.Visibility = Telerik.WinControls.ElementVisibility.Visible;
            // 
            // wizardPage3
            // 
            this.wizardPage3.ContentArea = this.panel5;
            this.wizardPage3.Name = "wizardPage3";
            this.wizardPage3.Visibility = Telerik.WinControls.ElementVisibility.Visible;
            // 
            // comboBox1
            // 
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(90, 26);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(121, 20);
            this.comboBox1.TabIndex = 1;
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
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(90, 68);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(283, 21);
            this.textBox1.TabIndex = 3;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(390, 66);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 4;
            this.button1.Text = "打开";
            this.button1.UseVisualStyleBackColor = true;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(390, 103);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 23);
            this.button2.TabIndex = 7;
            this.button2.Text = "打开";
            this.button2.UseVisualStyleBackColor = true;
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(90, 105);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(283, 21);
            this.textBox2.TabIndex = 6;
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
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(390, 139);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(75, 23);
            this.button3.TabIndex = 10;
            this.button3.Text = "打开";
            this.button3.UseVisualStyleBackColor = true;
            // 
            // textBox3
            // 
            this.textBox3.Location = new System.Drawing.Point(90, 141);
            this.textBox3.Name = "textBox3";
            this.textBox3.Size = new System.Drawing.Size(283, 21);
            this.textBox3.TabIndex = 9;
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
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(30, 197);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(435, 17);
            this.progressBar1.TabIndex = 11;
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(3, 29);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 23;
            this.dataGridView1.Size = new System.Drawing.Size(431, 306);
            this.dataGridView1.TabIndex = 0;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(6, 11);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(77, 12);
            this.label7.TabIndex = 1;
            this.label7.Text = "已导入数据：";
            // 
            // MyWizard
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(587, 440);
            this.Controls.Add(this.radWizard1);
            this.Name = "MyWizard";
            this.Text = "数据导入";
            ((System.ComponentModel.ISupportInitialize)(this.radWizard1)).EndInit();
            this.radWizard1.ResumeLayout(false);
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel5.ResumeLayout(false);
            this.panel5.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
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
        private System.Windows.Forms.ComboBox exam_type;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Panel panel5;
        private System.Windows.Forms.Label label3;
        private Telerik.WinControls.UI.WizardPage wizardPage3;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.TextBox textBox3;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ComboBox comboBox1;
    }
}