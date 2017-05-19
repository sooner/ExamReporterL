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

        string[] subs = {"yw","sxw","sxl", "yy", "lz", "wl","hx","sw", "wz", "zz","dl","ls" };
        string[] subs_total = { "yww", "yyw", "wz", "ywl", "sxl", "yyl", "lz" };

        string[] sum_wk = { "yww", "sxw", "yyw", "wz", "ls", "dl", "zz" };
        string[] sum_lk = { "ywl", "sxl", "yyl", "lz", "wl", "hx", "sw" };

        public DataTable summary;
        public Partition_wordcreator.ChartCombine year1_comb = new Partition_wordcreator.ChartCombine();
        public Partition_wordcreator.ChartCombine year2_comb = new Partition_wordcreator.ChartCombine();
        public Dictionary<string, WordData> year1_data = new Dictionary<string, WordData>();
        public Dictionary<string, WordData> year2_data = new Dictionary<string, WordData>();
        
        
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
            summary.Columns.Add("diff", typeof(decimal));
            
        }

        public void start(string year1, string year2, string exam)
        {

            CacheData cachedata = new CacheData();

            ZF_worddata year1_w_data = new ZF_worddata();
            ZF_worddata year1_l_data = new ZF_worddata();

            ZF_worddata year2_w_data = new ZF_worddata();
            ZF_worddata year2_l_data = new ZF_worddata();

            cachedata.load_zf_data(year1, "gk", "wk", year1_w_data);
            cachedata.load_zf_data(year1, "gk", "lk", year1_l_data);

            cachedata.load_zf_data(year2, "gk", "wk", year2_w_data);
            cachedata.load_zf_data(year2, "gk", "lk", year2_l_data);


            
            year1_comb.Add(year1_w_data.dist, "文科");
            year1_comb.Add(year1_l_data.dist, "理科");

            
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

                //insert_sub_to_summary(sub, year1, sub_data1);
                //insert_sub_to_summary(sub, year2, sub_data2);

            }

            decimal diff = year2_w_data.difficulty - year1_w_data.difficulty; 
            insert_zf_to_summary("wk", year1, year1_w_data, diff);
            insert_zf_to_summary("wk", year2, year2_w_data, diff);

            foreach (string sub in sum_wk)
            {
                decimal temp_diff;
                if (!year1_data.Keys.Contains(sub))
                {
                    PartitionData pdata1 = new PartitionData("");
                    cachedata.load_partitiondata(year1, exam, sub, pdata1);
                    PartitionData pdata2 = new PartitionData("");
                    cachedata.load_partitiondata(year2, exam, sub, pdata2);

                    temp_diff = pdata2.difficulty - pdata1.difficulty;

                    insert_pt_to_summary(sub, year1, pdata1, temp_diff);
                    insert_pt_to_summary(sub, year2, pdata2, temp_diff);
                }
                else
                {
                    temp_diff = year2_data[sub].difficulty - year1_data[sub].difficulty;
                    insert_sub_to_summary(sub, year1, year1_data[sub], temp_diff);
                    insert_sub_to_summary(sub, year2, year2_data[sub], temp_diff);
                }
            }

            decimal l_diff = year2_l_data.difficulty - year1_l_data.difficulty;
            insert_zf_to_summary("lk", year1, year1_l_data, l_diff);
            insert_zf_to_summary("lk", year2, year2_l_data, l_diff);

            foreach (string sub in sum_lk)
            {
                decimal temp_diff;
                if (!year1_data.Keys.Contains(sub))
                {
                    PartitionData pdata1 = new PartitionData("");
                    cachedata.load_partitiondata(year1, exam, sub, pdata1);
                    PartitionData pdata2 = new PartitionData("");
                    cachedata.load_partitiondata(year2, exam, sub, pdata2);

                    temp_diff = pdata2.difficulty - pdata1.difficulty;
                    insert_pt_to_summary(sub, year1, pdata1, temp_diff);
                    insert_pt_to_summary(sub, year2, pdata2, temp_diff);
                }
                else
                {
                    temp_diff = year2_data[sub].difficulty - year1_data[sub].difficulty;
                    insert_sub_to_summary(sub, year1, year1_data[sub], temp_diff);
                    insert_sub_to_summary(sub, year2, year2_data[sub], temp_diff);
                }
            }

        }


        public void insert_pt_to_summary(string name, string year, PartitionData data, decimal diff)
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
            dr["diff"] = diff;

            summary.Rows.Add(dr);
        }

        public void insert_zf_to_summary(string name, string year, ZF_worddata data, decimal diff)
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
            dr["diff"] = diff;

            summary.Rows.Add(dr);
        }

        public void insert_sub_to_summary(string name, string year, WordData data, decimal diff)
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
            dr["diff"] = diff;

            summary.Rows.Add(dr);
        }

        
    }
}
