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
            public decimal skewness;
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

        public DataTable districts;

        public Admin_WordData()
        {
            total = new basic_stat();
            urban = new basic_stat();
            country = new basic_stat();

            total_dist = new DataTable();
            total_freq = new DataTable();
            total_level = new DataTable();
            sub_diff = new DataTable();
            urban_sub = new DataTable();
            country_sub = new DataTable();
            districts = new DataTable();

            total_dist.Columns.Add("mark", typeof(int));
            total_dist.Columns.Add("count", typeof(int));

            total_freq.Columns.Add("totalmark", typeof(string));
            total_freq.Columns.Add("frequency", typeof(int));
            total_freq.Columns.Add("rate", typeof(decimal));
            total_freq.Columns.Add("accumulateFreq", typeof(int));
            total_freq.Columns.Add("accumulateRate", typeof(decimal));

            total_level.Columns.Add("text", typeof(string));
            total_level.Columns.Add("level", typeof(int));
            total_level.Columns.Add("frequency", typeof(int));
            total_level.Columns.Add("rate", typeof(decimal));

            sub_diff.Columns.Add("sub", typeof(string));
            sub_diff.Columns.Add("diff", typeof(decimal));
            sub_diff.Columns.Add("total", typeof(decimal));
            sub_diff.Columns.Add("avg", typeof(decimal));

            urban_sub.Columns.Add("sub", typeof(string));
            urban_sub.Columns.Add("diff", typeof(decimal));

            country_sub.Columns.Add("sub", typeof(string));
            country_sub.Columns.Add("diff", typeof(decimal));

            districts.Columns.Add("total", typeof(decimal));
            districts.Columns.Add("urban", typeof(decimal));
            districts.Columns.Add("country", typeof(decimal));
            districts.Columns.Add("dc", typeof(decimal));
            districts.Columns.Add("xc", typeof(decimal));
            districts.Columns.Add("hd", typeof(decimal));
            districts.Columns.Add("cy", typeof(decimal));
            districts.Columns.Add("sjs", typeof(decimal));
            districts.Columns.Add("ft", typeof(decimal));
            districts.Columns.Add("ys", typeof(decimal));
            districts.Columns.Add("tz", typeof(decimal));
            districts.Columns.Add("sy", typeof(decimal));
            districts.Columns.Add("cp", typeof(decimal));
            districts.Columns.Add("mtg", typeof(decimal));
            districts.Columns.Add("fs", typeof(decimal));
            districts.Columns.Add("dx", typeof(decimal));
            districts.Columns.Add("hr", typeof(decimal));
            districts.Columns.Add("pg", typeof(decimal));
            districts.Columns.Add("my", typeof(decimal));
            districts.Columns.Add("yq", typeof(decimal));
        }


    }
}
