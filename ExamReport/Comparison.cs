using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using System.Data;


namespace ExamReport
{
    public class Comparison
    {

        string[] subs = {"yw","sxw","sxl", "yy", "wl","hx","sw","zz","dl","ls" };
        string[] subs_total = { "yww", "yyw", "wz", "ywl", "sxl", "yyl", "lz" };

        string[] sum_wk = { "yww", "sxw", "yyw", "wz", "ls", "dl", "zz" };
        string[] sum_lk = { "ywl", "sxl", "yyl", "lz", "wl", "hx", "sw" };

        DataTable summary;

        public Comparison()
        {
            summary = new DataTable();

            summary.Columns.Add("sub", typeof(string));
            summary.Columns.Add("year", typeof(string));
            summary.Columns.Add("total_num", typeof(int));
            summary.Columns.Add("fullmark", typeof(decimal));
            summary.Columns.Add("max", typeof(decimal));
            summary.Columns.Add("min", typeof(decimal));
            summary.Columns.Add("avg", typeof(decimal));
            summary.Columns.Add("stDev", typeof(decimal));
            summary.Columns.Add("Dfactor", typeof(decimal));
            summary.Columns.Add("difficulty", typeof(decimal));

            summary.PrimaryKey = new DataColumn[] { summary.Columns["sub"] };

            
        }

        public void start(string year1, string year2, string exam)
        {
            Dictionary<string, WordData> year1_data = new Dictionary<string, WordData>();
            Dictionary<string, WordData> year2_data = new Dictionary<string, WordData>();

            CacheData cachedata = new CacheData();

            ZF_worddata year1_w_data = new ZF_worddata();
            ZF_worddata year1_l_data = new ZF_worddata();

            ZF_worddata year2_w_data = new ZF_worddata();
            ZF_worddata year2_l_data = new ZF_worddata();

            cachedata.load_zf_data(year1, "gk", "wk", year1_w_data);
            cachedata.load_zf_data(year1, "gk", "lk", year1_l_data);

            cachedata.load_zf_data(year2, "gk", "wk", year2_w_data);
            cachedata.load_zf_data(year2, "gk", "lk", year2_l_data);


            Partition_wordcreator.ChartCombine year1_comb = new Partition_wordcreator.ChartCombine();
            year1_comb.Add(year1_w_data.dist, "文科");
            year1_comb.Add(year1_l_data.dist, "理科");

            Partition_wordcreator.ChartCombine year2_comb = new Partition_wordcreator.ChartCombine();
            year2_comb.Add(year2_w_data.dist, "文科");
            year2_comb.Add(year2_l_data.dist, "理科");



            foreach (string sub in subs)
            {
                WordData sub_data1 = new WordData(null);
                cachedata.load_totaldata(year1, exam, sub, sub_data1);

                WordData sub_data2 = new WordData(null);
                cachedata.load_totaldata(year2, exam, sub, sub_data2);

                year1_data.Add(sub, sub_data1);
                year2_data.Add(sub, sub_data2);

                insert_sub_to_summary(sub, year1, sub_data1);
                insert_sub_to_summary(sub, year2, sub_data2);

            }


            insert_zf_to_summary("wk", year1, year1_w_data);
            insert_zf_to_summary("wk", year2, year2_w_data);

            foreach (string sub in sum_wk)
            {
                if (!year1_data.Keys.Contains(sub))
                {
                    Part
                }
                insert_sub_to_summary(sub, year1, year1_data[sub]);
            }
            insert_zf_to_summary("lk", year1, year1_l_data);
            insert_zf_to_summary("lk", year2, year2_l_data);



        }

        public void insert_zf_to_summary(string name, string year, ZF_worddata data)
        {
            DataRow dr = summary.NewRow();

            dr["sub"] = name;
            dr["year"] = year;
            dr["total_num"] = data.total_num;
            dr["fullmark"] = data.fullmark;
            dr["max"] = data.max;
            dr["min"] = data.min;
            dr["avg"] = data.avg;
            dr["stDev"] = data.stDev;
            dr["Dfactor"] = data.Dfactor;
            dr["difficulty"] = data.difficulty;

            summary.Rows.Add(dr);
        }

        public void insert_sub_to_summary(string name, string year, WordData data)
        {
            DataRow dr = summary.NewRow();

            dr["sub"] = name;
            dr["year"] = year;
            dr["total_num"] = data.total_num;
            dr["fullmark"] = data.fullmark;
            dr["max"] = data.max;
            dr["min"] = data.min;
            dr["avg"] = data.avg;
            dr["stDev"] = data.stDev;
            dr["Dfactor"] = data.Dfactor;
            dr["difficulty"] = data.difficulty;

            summary.Rows.Add(dr);
        }

        
    }
}
