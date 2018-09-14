using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace ExamReport
{
    class ZF_statistic
    {
        DataTable _data;
        public Configuration _config;
        public ZF_worddata w_result;
        public ZF_worddata l_result;

        public List<DataTable> sub;
        public decimal _fullmark;
        public string _name;
        public DataTable _choice;
        public Dictionary<string, ZF_worddata> results = new Dictionary<string, ZF_worddata>();
        public ZF_statistic(Configuration config, DataTable data, decimal fullmark, string name)
        {
            _data = data;
            w_result = new ZF_worddata();
            l_result = new ZF_worddata();
            _fullmark = fullmark;
            _name = name;
            _config = config;
        }

        

        public ZF_worddata init_null_data()
        {
            ZF_worddata data = new ZF_worddata();
            data.total_num = 0;
            data.fullmark = _fullmark;
            data.max = 0;
            data.min = 0;
            data.avg = 0;
            data.Dfactor = 0;
            data.difficulty = 0;
            data.stDev = 0;
            data.dist = null;
            data.frequency = null;
            return data;
        }

        public void zk_process()
        {
            foreach (int key in Utils.sub_choice.Keys)
            {
                DataTable temp_d = _data.equalfilter("xk", key.ToString());
                if (temp_d.Rows.Count == 0)
                    results.Add(Utils.sub_choice[key], init_null_data());
                else
                    results.Add(Utils.sub_choice[key], statistic_process(temp_d));
            }
        }

        public ZF_worddata statistic_process(DataTable dt)
        {

            ZF_worddata result = new ZF_worddata();
            result.total_num = dt.Rows.Count;
            result.fullmark = _fullmark;
            result.max = Convert.ToDecimal(dt.Compute("Max([zf])", ""));
            result.min = Convert.ToDecimal(dt.Compute("Min([zf])", ""));
            result.avg = Convert.ToDecimal(dt.Compute("Avg([zf])", ""));
            Partition_statistic.stdev single_stdev = new Partition_statistic.stdev(result.total_num, result.avg);
            foreach (DataRow dr in dt.Rows)
            {
                single_stdev.add((decimal)dr["zf"]);
            }
            result.difficulty = result.avg / result.fullmark;
            result.stDev = single_stdev.get_value();
            if (result.avg == 0)
                result.Dfactor = 0;
            else
                result.Dfactor = result.stDev / result.avg;

            var freq = from row in dt.AsEnumerable()
                       group row by row.Field<decimal>("zf") into grp
                       orderby grp.Key descending
                       select new
                       {
                           totalmark = grp.Key,
                           count = grp.Count()
                           //average = grp.Average(row => row.Field<decimal>(totalmark_str)) 
                       };
            bool first = true;
            int freqency = 0;
            bool isEven = result.total_num % 2 == 0;
            decimal mid;
            if (isEven)
                mid = result.total_num / 2.0m;
            else
                mid = (result.total_num + 1) / 2.0m;
            bool midCheck = true;
            int MaxFreq = 0;
            decimal total_interval = 1.0m;
            //decimal first_interval = 0.0m;
            decimal flag = 0.0m;
            if (result.fullmark > 20.0m)
            {
                total_interval = Math.Floor(result.fullmark / 20.0m);
                flag = (total_interval + 1) / 2.0m;

                int j = 0;
                for (j = 0; j < 20; j++)
                {
                    DataRow inter_row = result.dist.NewRow();
                    inter_row["mark"] = flag;
                    inter_row["rate"] = 0;
                    flag += total_interval;
                    result.dist.Rows.Add(inter_row);
                }
                if ((result.fullmark - 20.0m * total_interval) != 0)
                {
                    DataRow last_row = result.dist.NewRow();
                    last_row["mark"] = 20.0m * total_interval + (result.fullmark - 20.0m * total_interval + 1) / 2.0m;
                    last_row["rate"] = 0;
                    result.dist.Rows.Add(last_row);
                }
            }
            else
            {
                int j = 0;
                for (j = 0; j < result.fullmark; j++)
                {
                    DataRow inter_row = result.dist.NewRow();
                    inter_row["mark"] = Convert.ToDecimal(j + 1);
                    inter_row["rate"] = 0;
                    result.dist.Rows.Add(inter_row);
                }
            }

            int dist_num = 0;
            foreach (var item in freq)
            {
                DataRow dr = result.frequency.NewRow();
                dr["totalmark"] = item.totalmark;
                dr["frequency"] = item.count;
                dr["rate"] = ((decimal)item.count / result.total_num) * 100;


                if (first)
                {
                    dr["accumulateFreq"] = dr["frequency"];
                    dr["accumulateRate"] = dr["rate"];
                    freqency = (int)dr["frequency"];
                    first = false;
                }
                else
                {
                    dr["accumulateFreq"] = freqency + item.count;
                    dr["accumulateRate"] = ((int)dr["accumulateFreq"] / Convert.ToDecimal(result.total_num)) * 100;
                    freqency = (int)dr["accumulateFreq"];

                }

                if (total_interval == 1.0m)
                    dist_num = Convert.ToInt32(Math.Floor((decimal)dr["totalmark"]));
                else
                    dist_num = Convert.ToInt32(Math.Ceiling((decimal)dr["totalmark"] / total_interval));
                if (dist_num > 20)
                    result.dist.Rows[20]["rate"] = (decimal)result.dist.Rows[20]["rate"] + Convert.ToDecimal(dr["frequency"]);
                else if (dist_num == 0)
                    result.dist.Rows[dist_num]["rate"] = (decimal)result.dist.Rows[dist_num]["rate"] + Convert.ToDecimal(dr["frequency"]);
                else
                    result.dist.Rows[dist_num - 1]["rate"] = (decimal)result.dist.Rows[dist_num - 1]["rate"] + Convert.ToDecimal(dr["frequency"]);

                if (midCheck && (int)dr["accumulateFreq"] >= mid)
                {
                    if (result.frequency.Rows.Count == 0)
                        result.mean = (decimal)dr["totalmark"];
                    else
                    {
                        DataRow midRow = result.frequency.Rows[result.frequency.Rows.Count - 1];
                        if ((int)dr["frequency"] == 1)
                            if (isEven)
                                result.mean = ((decimal)dr["totalmark"] + (decimal)midRow["totalmark"]) / 2;
                            else
                                result.mean = (decimal)dr["totalmark"];
                        else
                        {
                            int fb = (int)dr["accumulateFreq"] - (int)dr["frequency"];
                            if (isEven)
                                result.mean = (decimal)dr["totalmark"] + 0.5m - (mid - fb) * (1.0m / (int)dr["frequency"]);
                            else
                                result.mean = (decimal)dr["totalmark"] + 0.5m - (mid - fb - 0.5m) * (1.0m / (int)dr["frequency"]);
                        }
                    }
                    midCheck = false;
                }

                result.frequency.Rows.Add(dr);
            }
            DataTable new_freq = result.frequency.Clone();
            new_freq.PrimaryKey = new DataColumn[] { new_freq.Columns["totalmark"] };
            foreach (DataRow dr in result.frequency.Rows)
            {
                decimal keyMark = cus_round(1, (decimal)dr["totalmark"]);

                if (!new_freq.Rows.Contains(keyMark))
                {
                    dr["totalmark"] = keyMark;
                    new_freq.ImportRow(dr);
                }
                else
                {
                    DataRow oldrow = new_freq.Rows.Find(keyMark);
                    oldrow["frequency"] = (int)oldrow["frequency"] + (int)dr["frequency"];
                    oldrow["rate"] = ((int)oldrow["frequency"] / (decimal)result.total_num) * 100;
                    oldrow["accumulateFreq"] = (int)oldrow["accumulateFreq"] + (int)dr["frequency"];
                    oldrow["accumulateRate"] = ((int)oldrow["accumulateFreq"] / (decimal)result.total_num) * 100;
                    if (MaxFreq < (int)oldrow["frequency"])
                    {
                        MaxFreq = (int)oldrow["frequency"];
                    }

                }
                if (MaxFreq < (int)dr["frequency"])
                {
                    MaxFreq = (int)dr["frequency"];
                }
            }
            result.frequency = new_freq;

            foreach (DataRow dr in result.dist.Rows)
            {
                dr["rate"] = (decimal)dr["rate"] / Convert.ToDecimal(result.total_num) * 100;
            }
            return result;
        }
        public void partition_process()
        {
            DataTable w_data = _data.equalfilter("type", "w");
            DataTable l_data = _data.equalfilter("type", "l");

            single_process(w_data, w_result);
            single_process(l_data, l_result);

            if (_config.report_style.Equals("总体"))
            {
                sub = new List<DataTable>();
                insertSub(sub, _data, "yw", 150m);
                insertSub(sub, _data.equalfilter("type", "w"), "sx", 150m);
                insertSub(sub, _data.equalfilter("type", "l"), "sx", 150m);
                insertSub(sub, _data, "yy", 150m);
                insertSub(sub, _data.equalfilter("type", "w"), "zh", 300m);
                insertSub(sub, _data.equalfilter("type", "l"), "zh", 300m);
            }

        }
        public void insertSub(List<DataTable> list, DataTable dt, string sub, decimal fullmark)
        {
            
            DataTable data = new DataTable();
            data.Columns.Add("mark", typeof(decimal));
            data.Columns.Add("difficulty", typeof(decimal));

            decimal flag = 0m;
            decimal interval = 1.0m;
            if (_fullmark > 40.0m)
            {
                interval = Math.Floor(_fullmark / 40.0m);
                flag = (interval + 1) / 2.0m;

                int j = 0;
                for (j = 0; j < 40; j++)
                {
                    DataRow inter_row = data.NewRow();
                    inter_row["mark"] = flag;
                    inter_row["difficulty"] = 0;
                    flag += interval;
                    data.Rows.Add(inter_row);
                }
                if ((_fullmark - 40.0m * interval) != 0)
                {
                    DataRow last_row = data.NewRow();
                    last_row["mark"] = 40.0m * interval + (_fullmark - 40.0m * interval + 1) / 2.0m;
                    last_row["difficulty"] = 0;
                    data.Rows.Add(last_row);
                }
            }
            else
            {
                int j = 0;
                for (j = 0; j < _fullmark; j++)
                {
                    DataRow inter_row = data.NewRow();
                    inter_row["mark"] = Convert.ToDecimal(j + 1);
                    inter_row["difficulty"] = 0;
                    data.Rows.Add(inter_row);
                }
            }

            
            var yw = from row in dt.AsEnumerable()
                     group row by row.Field<decimal>("zf") into grp
                     orderby grp.Key ascending
                     select new
                     {
                         mark = grp.Key,
                         avg = grp.Average(row => row.Field<decimal>(sub)),
                         count = grp.Count()
                     };
            int dist_num = 0;
            int[] count = new int[data.Rows.Count];
            for (int k = 0; k < data.Rows.Count; k++)
            {
                count[k] = 0;
            }
            foreach (var item in yw)
            {
                //DataRow dr = data.NewRow();
                //dr["mark"] = item.mark;
                //dr["difficulty"] = item.avg / fullmark;
                //data.Rows.Add(dr);
                if (interval == 1.0m)
                    dist_num = Convert.ToInt32(Math.Floor(item.mark));
                else
                    dist_num = Convert.ToInt32(Math.Ceiling(item.mark / interval));
                if (dist_num > 40)
                {
                    data.Rows[40]["difficulty"] = (decimal)data.Rows[40]["difficulty"] + Convert.ToDecimal(item.count * item.avg);
                    count[40] += item.count;
                }
                else if (dist_num == 0)
                {
                    data.Rows[dist_num]["difficulty"] = (decimal)data.Rows[dist_num]["difficulty"] + Convert.ToDecimal(item.count * item.avg);
                    count[dist_num] += item.count;
                }
                else
                {
                    data.Rows[dist_num - 1]["difficulty"] = (decimal)data.Rows[dist_num - 1]["difficulty"] + Convert.ToDecimal(item.count * item.avg);
                    count[dist_num - 1] += item.count;
                }
            }
            for (int i = 0; i < data.Rows.Count; i++)
            {
                if (count[i] == 0)
                    data.Rows[i]["difficulty"] = 0;
                    
                else
                    data.Rows[i]["difficulty"] = ((decimal)data.Rows[i]["difficulty"] / count[i]) / fullmark;

            }
            while (true)
            {
                if((decimal)data.Rows[data.Rows.Count - 1]["difficulty"] == 0)
                    data.Rows.Remove(data.Rows[data.Rows.Count - 1]);
                else
                    break;
            }
            //for (int i = 0; i < data.Rows.Count; i++)
            //{
            //    if ((decimal)data.Rows[i]["difficulty"] == 0)
            //        data.Rows.Remove(data.Rows[i]);
            //}
            list.Add(data);
        }
        public void single_process(DataTable dt, ZF_worddata data)
        {
            data.total_num = dt.Rows.Count;
            data.fullmark = _fullmark;
            data.max = Convert.ToDecimal(dt.Compute("Max(zf)", ""));
            data.min = Convert.ToDecimal(dt.Compute("Min(zf)", ""));
            data.avg = Convert.ToDecimal(dt.Compute("Avg(zf)", ""));
            Partition_statistic.stdev stdev = new Partition_statistic.stdev(data.total_num, data.avg);
            foreach (DataRow dr in dt.Rows)
            {
                stdev.add(Convert.ToDecimal(dr["zf"]));
            }
            data.stDev = stdev.get_value();
            data.Dfactor = data.stDev / data.avg;
            data.difficulty = data.avg / data.fullmark;

            decimal flag = 0m;
            decimal interval = 1.0m;
            if (data.fullmark > 10.0m)
            {
                int tuple = Convert.ToInt32(Math.Floor(data.fullmark / 10.0m));
                interval = 10;
                flag = (interval + 1) / 2.0m;

                int j = 0;

                DataRow start_row = data.dist.NewRow();
                start_row["mark"] = 0;
                start_row["rate"] = 0;
                data.dist.Rows.Add(start_row);

                for (j = 0; j < tuple; j++)
                {
                    DataRow inter_row = data.dist.NewRow();
                    inter_row["mark"] = flag;
                    inter_row["rate"] = 0;
                    flag += interval;
                    data.dist.Rows.Add(inter_row);
                }
                if ((data.fullmark - tuple * interval) != 0)
                {
                    DataRow last_row = data.dist.NewRow();
                    last_row["mark"] = 20.0m * interval + (data.fullmark - 20.0m * interval + 1) / 2.0m;
                    last_row["rate"] = 0;
                    data.dist.Rows.Add(last_row);
                }
            }
            else
            {
                int j = 0;
                for (j = 0; j <= data.fullmark; j++)
                {
                    DataRow inter_row = data.dist.NewRow();
                    inter_row["mark"] = Convert.ToDecimal(j);
                    inter_row["rate"] = 0;
                    data.dist.Rows.Add(inter_row);
                }
            }
            var temp = from row in dt.AsEnumerable()
                       group row by row.Field<decimal>("zf") into grp
                       orderby grp.Key descending
                       select new
                       {
                           totalmark = grp.Key,
                           count = grp.Count()
                       };

            bool first = true;
            int lastcount = 0;
            foreach (var item in temp)
            {
                DataRow dr = data.frequency.NewRow();
                dr["totalmark"] = item.totalmark;
                dr["frequency"] = item.count;
                dr["rate"] = item.count / Convert.ToDecimal(data.total_num) * 100m;
                if (first)
                {
                    dr["accumulateFreq"] = item.count;
                    dr["accumulateRate"] = dr["rate"];
                    lastcount = item.count;
                    first = false;
                }
                else
                {
                    dr["accumulateFreq"] = item.count + lastcount;
                    dr["accumulateRate"] = (int)dr["accumulateFreq"] / Convert.ToDecimal(data.total_num) * 100m;
                    lastcount = (int)dr["accumulateFreq"];
                }
                data.frequency.Rows.Add(dr);
            }
            int dist_num = 0;
            for (int i = data.frequency.Rows.Count - 1; i >= 0; i--)
            {
                if (interval == 1.0m)
                {
                    dist_num = Convert.ToInt32(Math.Floor((decimal)data.frequency.Rows[i]["totalmark"]));
                    data.dist.Rows[dist_num]["rate"] = (decimal)data.frequency.Rows[dist_num]["rate"] + Convert.ToDecimal(data.frequency.Rows[i]["frequency"]);
                }
                else
                {
                    dist_num = Convert.ToInt32(Math.Ceiling((decimal)data.frequency.Rows[i]["totalmark"] / interval));
                    //if (dist_num > 20)
                    //    data.dist.Rows[20]["rate"] = (decimal)data.frequency.Rows[20]["rate"] + Convert.ToDecimal(data.frequency.Rows[i]["frequency"]);
                    //else if (dist_num == 0)
                        data.dist.Rows[dist_num]["rate"] = (decimal)data.dist.Rows[dist_num]["rate"] + Convert.ToDecimal(data.frequency.Rows[i]["frequency"]);
                    //else
                        //data.dist.Rows[dist_num - 1]["rate"] = (decimal)data.dist.Rows[dist_num - 1]["rate"] + Convert.ToDecimal(data.frequency.Rows[i]["frequency"]);
                }
            }
            foreach (DataRow dr in data.dist.Rows)
            {
                dr["rate"] = (decimal)dr["rate"] / data.total_num * 100;
            }
            DataTable new_freq = data.frequency.Clone();
            new_freq.PrimaryKey = new DataColumn[] { new_freq.Columns["totalmark"] };
            foreach (DataRow dr in data.frequency.Rows)
            {
                decimal keyMark = cus_round(1, (decimal)dr["totalmark"]);

                if (!new_freq.Rows.Contains(keyMark))
                {
                    dr["totalmark"] = keyMark;
                    new_freq.ImportRow(dr);
                }
                else
                {
                    DataRow oldrow = new_freq.Rows.Find(keyMark);
                    oldrow["frequency"] = (int)oldrow["frequency"] + (int)dr["frequency"];
                    oldrow["rate"] = ((int)oldrow["frequency"] / (decimal)data.total_num) * 100;
                    oldrow["accumulateFreq"] = (int)oldrow["accumulateFreq"] + (int)dr["frequency"];
                    oldrow["accumulateRate"] = ((int)oldrow["accumulateFreq"] / (decimal)data.total_num) * 100;
                }

            }
            data.frequency = new_freq;
            //foreach (DataRow dr in data.dist.Rows)
            //{
            //    dr["rate"] = (decimal)dr["rate"] / Convert.ToDecimal(data.total_num) * 100m;
            //}

        }

        public decimal cus_round(double mark, decimal num)
        {
            decimal temp = num * Convert.ToDecimal(Math.Pow(10.0, mark - 1));
            decimal floor = Convert.ToDecimal(Math.Floor(Convert.ToDouble(temp)));
            if (temp < (floor + 0.5m))
                return floor / Convert.ToDecimal(Math.Pow(10.0, mark - 1));
            else
                return (floor + 1) / Convert.ToDecimal(Math.Pow(10.0, mark - 1));
        }
    }
}
