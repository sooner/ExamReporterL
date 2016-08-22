using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace ExamReport
{
    class AdminCal
    {
        DataTable _data;
        decimal _fullmark;
        //string _name;
        List<string[]> qx_code = new List<string[]>();

        public Admin_WordData w_result;
        public Admin_WordData l_result;

        public Configuration _config;

        public AdminCal(Configuration config, DataTable data, decimal fullmark)
        {
            _data = data;
            _fullmark = fullmark;

            w_result = new Admin_WordData();
            l_result = new Admin_WordData();

            _config = config;

            qx_init();
        }

        public void qx_init()
        {
            qx_code.Add(new string[2]{"01", "03"});
            qx_code.Add(new string[1] { "02" });
            qx_code.Add(new string[1] { "05" });
            qx_code.Add(new string[1] { "06" });
            qx_code.Add(new string[1] { "07" });
            qx_code.Add(new string[1] { "08" });
            qx_code.Add(new string[1] { "09" });
            qx_code.Add(new string[1] { "10" });
            qx_code.Add(new string[1] { "11" });
            qx_code.Add(new string[1] { "12" });
            qx_code.Add(new string[1] { "13" });
            qx_code.Add(new string[1] { "14" });
            qx_code.Add(new string[1] { "15" });
            qx_code.Add(new string[1] { "16" });
            qx_code.Add(new string[1] { "17" });
            qx_code.Add(new string[1] { "28" });
            qx_code.Add(new string[1] { "29" });
        }
        public void Calculate()
        {
            DataTable w_data = _data.equalfilter("type", "w");
            DataTable l_data = _data.equalfilter("type", "l");

            single_process(w_data, w_result, _config.wen_first_level, _config.wen_second_level, _config.wen_third_level);
            single_process(l_data, l_result, _config.li_first_level, _config.li_second_level, _config.li_third_level);

            wen_process(w_data, w_result);
            li_process(l_data, l_result);
        }

        public void wen_process(DataTable data, Admin_WordData result)
        {
            wen_sub(data, result.sub_diff);

            DataTable urban = data.filteredtable("qxdm", _config.urban_code);
            DataTable country = data.filteredtable("qxdm", _config.country_code);
            wen_sub(urban, result.urban_sub);
            wen_sub(country, result.country_sub);

            qx_process("zf", 750, data, result.districts);
            qx_process("yw", 150, data, result.districts);
            qx_process("sx", 150, data, result.districts);
            qx_process("yy", 150, data, result.districts);
            qx_process("sw_ls", 100, data, result.districts);
            qx_process("wl_dl", 100, data, result.districts);
            qx_process("hx_zz", 100, data, result.districts);
            qx_process("zh", 300, data, result.districts);

        }

        public void li_process(DataTable data, Admin_WordData result)
        {
            li_sub(data, result.sub_diff);
            DataTable urban = data.filteredtable("qxdm", _config.urban_code);
            DataTable country = data.filteredtable("qxdm", _config.country_code);
            li_sub(urban, result.urban_sub);
            li_sub(country, result.country_sub);

            qx_process("zf", 750, data, result.districts);
            qx_process("yw", 150, data, result.districts);
            qx_process("sx", 150, data, result.districts);
            qx_process("yy", 150, data, result.districts);
            qx_process("wl_dl", 120, data, result.districts);
            qx_process("hx_zz", 100, data, result.districts);
            qx_process("sw_ls", 80, data, result.districts);
            qx_process("zh", 300, data, result.districts);

        }
        public void li_sub(DataTable data, DataTable sub)
        {
            InsertSubDiff("语文", "yw", 150, data, sub);
            InsertSubDiff("数学理", "sx", 150, data, sub);
            InsertSubDiff("英语", "yy", 150, data, sub);
            InsertSubDiff("物理", "wl_dl", 120, data, sub);
            InsertSubDiff("化学", "hx_zz", 100, data, sub);
            InsertSubDiff("生物", "sw_ls", 80, data, sub);
            InsertSubDiff("理综", "zh", 300, data, sub);
        }
        public void qx_process(string sub, decimal fullmark, DataTable data, DataTable districts)
        {
            DataRow avg = districts.NewRow();
            DataRow diff = districts.NewRow();
            if (sub.Equals("zh"))
            {
                avg["total"] = data.AsEnumerable().Average(c => (c.Field<decimal>("sw_ls") + c.Field<decimal>("hx_zz") + c.Field<decimal>("wl_dl")));
                diff["total"] = Convert.ToDecimal(avg["total"]) / fullmark;

                avg["urban"] = data.filteredtable("qxdm", _config.urban_code).AsEnumerable().Average(c => (c.Field<decimal>("sw_ls") + c.Field<decimal>("hx_zz") + c.Field<decimal>("wl_dl")));
                diff["urban"] = Convert.ToDecimal(avg["urban"]) / fullmark;

                avg["country"] = data.filteredtable("qxdm", _config.country_code).AsEnumerable().Average(c => (c.Field<decimal>("sw_ls") + c.Field<decimal>("hx_zz") + c.Field<decimal>("wl_dl")));
                diff["country"] = Convert.ToDecimal(avg["country"]) / fullmark;

                for (int i = 0; i < qx_code.Count; i++)
                {
                    avg[i + 3] = data.filteredtable("qxdm", qx_code[i]).AsEnumerable().Average(c => (c.Field<decimal>("sw_ls") + c.Field<decimal>("hx_zz") + c.Field<decimal>("wl_dl")));
                    diff[i + 3] = Convert.ToDecimal(avg[i + 3]) / fullmark;
                }
            }
            else
            {
                avg["total"] = data.AsEnumerable().Average(c => c.Field<decimal>(sub));
                diff["total"] = Convert.ToDecimal(avg["total"]) / fullmark;

                avg["urban"] = data.filteredtable("qxdm", _config.urban_code).AsEnumerable().Average(c => c.Field<decimal>(sub));
                diff["urban"] = Convert.ToDecimal(avg["urban"]) / fullmark;

                avg["country"] = data.filteredtable("qxdm", _config.country_code).AsEnumerable().Average(c => c.Field<decimal>(sub));
                diff["country"] = Convert.ToDecimal(avg["country"]) / fullmark;

                for (int i = 0; i < qx_code.Count; i++)
                {
                    avg[i + 3] = data.filteredtable("qxdm", qx_code[i]).AsEnumerable().Average(c => c.Field<decimal>(sub));
                    diff[i + 3] = Convert.ToDecimal(avg[i + 3]) / fullmark;
                }
            }
            districts.Rows.Add(avg);
            districts.Rows.Add(diff);
        }
        public void wen_sub(DataTable data, DataTable sub)
        {
            InsertSubDiff("语文", "yw", 150, data, sub);
            InsertSubDiff("数学文", "sx", 150, data, sub);
            InsertSubDiff("英语", "yy", 150, data, sub);
            InsertSubDiff("历史", "sw_ls", 100, data, sub);
            InsertSubDiff("地理", "wl_dl", 100, data, sub);
            InsertSubDiff("政治", "hx_zz", 100, data, sub);
            InsertSubDiff("文综", "zh", 300, data, sub);
        }

        public void single_process(DataTable data, Admin_WordData result, int first, int second, int third)
        {
            total_statistic(data, result.total);

            for (int i = 0; i < _fullmark; i = i + 25)
            {
                DataRow dr = result.total_dist.NewRow();
                int min = i;
                int max = i + 25;
                dr["mark"] = i + 13;
                if (max != _fullmark)
                {

                    int count = data.AsEnumerable().Where(c => (c.Field<decimal>("zf") >= min && c.Field<decimal>("zf") < max)).Count();
                    dr["count"] = count;
                }
                else
                {
                    int count = data.AsEnumerable().Where(c => (c.Field<decimal>("zf") >= min && c.Field<decimal>("zf") <= max)).Count();
                    dr["count"] = count;
                }
                
                result.total_dist.Rows.Add(dr);

            }

            int acct_count = 0;
            for (int i = Convert.ToInt32(_fullmark); i > 0; i = i - 25)
            {
                int min = i - 25;
                int max = i;

                int count = data.AsEnumerable().Where(c => (c.Field<decimal>("zf") > min && c.Field<decimal>("zf") <= max)).Count();
                DataRow dr = result.total_freq.NewRow();
                dr["totalmark"] = (min + 1).ToString() + "～" + max.ToString();
                dr["frequency"] = count;
                dr["rate"] = count / Convert.ToDecimal(result.total.totalnum) * 100;
                dr["accumulateFreq"] = count + acct_count;
                dr["accumulateRate"] = (count + acct_count) / Convert.ToDecimal(result.total.totalnum) * 100;

                if (count == 0)
                    continue;
                result.total_freq.Rows.Add(dr);
                acct_count += count;

            }
            InsertTotalLevel("650以上", Convert.ToInt32(_fullmark), 650, data, result);
            InsertTotalLevel("600以上", Convert.ToInt32(_fullmark), 600, data, result);
            InsertTotalLevel("一本",  Convert.ToInt32(_fullmark), first, data, result);
            InsertTotalLevel("二本", first, second, data, result);
            InsertTotalLevel("三本", second, third, data, result);

            

            DataTable urban = data.filteredtable("qxdm", _config.urban_code);
            DataTable country = data.filteredtable("qxdm", _config.country_code);

            total_statistic(urban, result.urban);
            total_statistic(country, result.country);



        }

        public void total_statistic(DataTable data, Admin_WordData.basic_stat stat)
        {
            stat.totalnum = data.Rows.Count;
            stat.fullmark = _fullmark;
            stat.max = data.Max("zf");
            stat.min = data.Min("zf");
            stat.avg = data.Avg("zf");
            stat.stDev = data.StDev("zf");
            stat.Dfactor = stat.stDev / stat.avg;
            stat.difficulty = stat.avg / _fullmark;
        }

        public void InsertSubDiff(string sub_str, string sub, decimal sub_full, DataTable data, DataTable sub_dt)
        {
            DataRow dr = sub_dt.NewRow();
            dr["sub"] = sub_str;
            if (sub.Equals("zh"))
                dr["diff"] = Convert.ToDecimal(data.AsEnumerable().Average(c => (c.Field<decimal>("sw_ls") + c.Field<decimal>("wl_dl") + c.Field<decimal>("hx_zz")))) / sub_full;
            else
                dr["diff"] = Convert.ToDecimal(data.AsEnumerable().Average(c => c.Field<decimal>(sub))) / sub_full;
            sub_dt.Rows.Add(dr);
        }

        public void InsertTotalLevel(string level_str, int max, int min, DataTable data, Admin_WordData result)
        {
            int totalnum = data.Rows.Count;

            int count;
            if(level_str.Contains("以上"))
                count = data.AsEnumerable().Where(c => (c.Field<decimal>("zf") >= min && c.Field<decimal>("zf") <= max)).Count();
            else
                count = data.AsEnumerable().Where(c => (c.Field<decimal>("zf") >= min && c.Field<decimal>("zf") < max)).Count();
            DataRow dr = result.total_level.NewRow();
            dr["text"] = level_str;
            dr["level"] = min;
            dr["frequency"] = count;
            dr["rate"] = count / Convert.ToDecimal(totalnum) * 100;
            result.total_level.Rows.Add(dr);
        }

    }
}
