﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExamReport
{
    public class Configuration
    {
        public string subject;
        public string exam;
        public string save_address;
        public string report_style;
        public string CurrentDirectory;
        public bool isVisible = true;

        public decimal shengwu_zhengzhi;
        public decimal wuli_lishi;
        public decimal huaxue_dili;

        public string year = "2015";
        public string month = "6月";

        public string QX = "";
        public string school = "";

        public bool WSLG = false;
        public bool OnlyQZT = false;

        public List<decimal> GroupMark = new List<decimal>();
        public decimal fullmark;

        public List<decimal> sub_groupMark = new List<decimal>();
        public decimal sub_fullmark;

        public int smooth_degree = 10;
        public bool is_sub_cor = false;

        public void change()
        {
            fullmark = sub_fullmark;
            GroupMark = sub_groupMark;
        }
    } 
}