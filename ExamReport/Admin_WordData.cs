using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace ExamReport
{
    class Admin_WordData
    {
        public class basic_stat
        {
            public int totalnum;
            public decimal fullmark;
            public decimal max;
            public decimal min;
            public decimal avg;
            public decimal stDev;
            public decimal Dfactor;
            public decimal difficulty;
        }
        
        public basic_stat total;

        public DataTable total_dist;
        public DataTable total_freq;

        public DataTable total_level;
        public DataTable sub_diff;

        public basic_stat urban;
        public basic_stat country;

        public DataTable urban_sub;
        public DataTable country_sub;

        public List<DataTable> districts;

        public Admin_WordData()
        {
            total = new basic_stat();
            urban = new basic_stat();
            country = new basic_stat();


        }


    }
}
