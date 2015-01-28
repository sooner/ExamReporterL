using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Telerik.WinControls;

namespace ExamReport
{
    public partial class mainform : Telerik.WinControls.UI.RadForm
    {
        public mainform()
        {
            InitializeComponent();
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
    }
}
