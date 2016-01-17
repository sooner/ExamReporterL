using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.Linq;
using System.Text.RegularExpressions;
using System.Collections;
using System.Diagnostics;

namespace ExamReport
{
    class Total_statistic
    {
        DataTable _basic_data;
        DataTable _groups_table;
        DataTable _standard_ans;
        DataTable _groups_ans;
        public Configuration _config;
        public WordData result;
        decimal _fullmark;
        decimal PHN;
        decimal PLN;
        
        int _groupnum;

        string cor_col = "totalmark";

        public Total_statistic(WordData _result, DataTable dt, decimal fullmark, DataTable standard_ans, DataTable groups_table, DataTable groups_ans, int groupnum)
        {
            _basic_data = dt;
            _fullmark = fullmark;
            result = _result;
            _standard_ans = standard_ans;
            _standard_ans.PrimaryKey = new DataColumn[] { _standard_ans.Columns["th"] };
            PLN = Convert.ToDecimal(Math.Ceiling(_basic_data.Rows.Count * 0.27));
            PHN = _basic_data.Rows.Count - PLN;
            
            _groups_table = groups_table;
            _groups_ans = groups_ans;
            _groupnum = groupnum;
            result._groups_ans = _groups_ans;
            result._standard_ans = _standard_ans;
        }

        public bool statistic_process(bool isZonghe)
        {
            result.total_num = _basic_data.Rows.Count;
            result.fullmark = _fullmark;
            string totalmark_str;

            if (_basic_data.Columns.Contains("ZH_totalmark"))
                totalmark_str = "ZH_totalmark";
            else
                totalmark_str = "totalmark";

            if (!isZonghe && _config.is_sub_cor)
            {
                cor_col = "ZH_totalmark";
                
            }
            result.max = (decimal) _basic_data.Compute("Max(" + totalmark_str + ")", "");
            result.min = (decimal) _basic_data.Compute("Min(" + totalmark_str + ")", "");
            result.avg = (decimal) _basic_data.Compute("Avg(" + totalmark_str + ")", "");
            decimal ZH_avg = (decimal)_basic_data.Compute("Avg(" + cor_col + ")", "");
            //result.stDev = Convert.ToDecimal(Math.Sqrt(Convert.ToDouble(_basic_data.AsEnumerable().Select(c => (c.Field<decimal>(totalmark_str) - result.avg) * (c.Field<decimal>(totalmark_str) - result.avg)).Average())));
            Regex number = new Regex("^[Tt]\\d");
            #region total analysis table process
            //Stopwatch sw = new Stopwatch();
            //sw.Start();
            //decimal totalsum = _basic_data.AsEnumerable().Select(c => c.Field<decimal>("totalmark")).Sum();
            //decimal y2 = _basic_data.AsEnumerable().Select(c => c.Field<decimal>("totalmark") * c.Field<decimal>("totalmark")).Sum();
            //decimal y_result = y2 - totalsum * totalsum / result.total_num;
            foreach (DataColumn dc in _basic_data.Columns)
            {
                if(number.IsMatch(dc.ColumnName))
                {
                    
                    DataRow newR = result.total_analysis.NewRow();
                    string topic_num = dc.ColumnName.Substring(1);

                    newR["number"] = dc.ColumnName;
                    newR["fullmark"] = Convert.ToDecimal(_standard_ans.Rows.Find(topic_num)["fs"]);
                    newR["max"] = _basic_data.Compute("Max([" + dc.ColumnName + "])", "");
                    newR["min"] = _basic_data.Compute("Min([" + dc.ColumnName + "])", "");
                    newR["avg"] = _basic_data.Compute("Avg([" + dc.ColumnName + "])", "");
                    newR["standardErr"] = 0.0;
                    newR["dfactor"] = 0.0;
                    newR["difficulty"] = 0.0;
                    newR["correlation"] = 0.0;
                    newR["discriminant"] = 0.0;
                    newR["PHN"] = 0.0;
                    newR["PLN"] = 0.0;
                    newR["CorrectMark"] = 0.0;
                    newR["CorrectNum"] = 0.0;
                    newR["WrongMark"] = 0.0;
                    newR["WrongNum"] = 0.0;
                    newR["MultipleSum"] = 0.0;
                    newR["SquareSumX"] = 0.0;

                    if (!_standard_ans.Rows.Find(topic_num)["da"].Equals(""))
                        newR["objective"] = 1;
                    else
                        newR["objective"] = 0;
                    //newR["max"] = _basic_data.AsEnumerable().Max(c => c.Field<decimal>(dc.ColumnName));
                    //newR["min"] = _basic_data.AsEnumerable().Min(c => c.Field<decimal>(dc.ColumnName));
                    //newR["avg"] = _basic_data.AsEnumerable().Average(c => c.Field<decimal>(dc.ColumnName));
                    //newR["standardErr"] = Convert.ToDecimal(Math.Sqrt(_basic_data.AsEnumerable().Select(
                    //    c => Math.Pow(Convert.ToDouble(c.Field<decimal>(dc.ColumnName) - (decimal)newR["avg"]), 2)).Average()));
                    //newR["dfactor"] = (decimal)newR["standardErr"] / (decimal)newR["avg"];
                    //newR["difficulty"] = (decimal)newR["avg"] / (decimal)newR["fullmark"];
                    //if (!_standard_ans.Rows.Find(topic_num)["da"].Equals(""))
                    //{

                    //    newR["correlation"] = (_basic_data.AsEnumerable().Where(c => c.Field<decimal>(dc.ColumnName) > 0).Select(c => c.Field<decimal>("totalmark")).Average()
                    //        - _basic_data.AsEnumerable().Where(c => c.Field<decimal>(dc.ColumnName) == 0).Select(c => c.Field<decimal>("totalmark")).Average())
                    //        / result.stDev * Convert.ToDecimal(Math.Sqrt(Convert.ToDouble(_basic_data.AsEnumerable().Where(c => c.Field<decimal>(dc.ColumnName) > 0).Count() *
                    //        _basic_data.AsEnumerable().Where(c => c.Field<decimal>(dc.ColumnName) == 0).Count())));
                    //}
                    //else
                    //{
                    //    decimal xy = _basic_data.AsEnumerable().Select(c => c.Field<decimal>(dc.ColumnName) * c.Field<decimal>("totalmark")).Sum();
                    //    decimal xyn = (decimal)newR["avg"] * totalsum;
                    //    decimal x2 = _basic_data.AsEnumerable().Select(c => c.Field<decimal>(dc.ColumnName) * c.Field<decimal>(dc.ColumnName)).Sum();
                    //    decimal x_result = x2 - (decimal)newR["avg"] * (decimal)newR["avg"] * result.total_num;
                    //    newR["correlation"] = xy - xyn / Convert.ToDecimal(Math.Sqrt(Convert.ToDouble(y_result * x_result)));

                    //}
                    //decimal PH = _basic_data.AsEnumerable().OrderByDescending(c => c.Field<decimal>("totalmark")).Take(Convert.ToInt32(PLN)).Select(c => c.Field<decimal>(dc.ColumnName)).Average();
                    //decimal PL = _basic_data.AsEnumerable().OrderBy(c => c.Field<decimal>("totalmark")).Take(Convert.ToInt32(PLN)).Select(c => c.Field<decimal>(dc.ColumnName)).Average();
                    //newR["discriminant"] = PH - PL;


                    result.total_analysis.Rows.Add(newR);
                }
                
            }
            int count = 0;
            decimal Var = 0.0m;
            decimal ThreeMulti = 0.0m;
            decimal FourMulti = 0.0m;
            decimal SquareSumY = 0.0m;
            //decimal alfaMultiplySum = 0.0m;
            decimal alfaSquareX = 0.0m;
            decimal alfaSquareY = 0.0m;
            decimal alfaSumX = 0.0m;
            decimal alfaSumY = 0.0m;

            decimal ZH_var = 0;
            Regex oddeven = new Regex("^[Tt]\\d+$");
            foreach (DataRow dr in _basic_data.Rows)
            {
                decimal odd = 0.0m;
                decimal even = 0.0m;
                count++;
                Var += ((decimal)dr[totalmark_str] - result.avg) * ((decimal)dr[totalmark_str] - result.avg);
                if (!isZonghe)
                    ZH_var += ((decimal)dr[cor_col] - ZH_avg) * ((decimal)dr[cor_col] - ZH_avg);
                ThreeMulti += Convert.ToDecimal(Math.Pow(Convert.ToDouble((decimal)dr[totalmark_str] - result.avg), 3.0));
                FourMulti += Convert.ToDecimal(Math.Pow(Convert.ToDouble((decimal)dr[totalmark_str] - result.avg), 4.0));
                SquareSumY += (decimal)dr[cor_col] * (decimal)dr[cor_col];
                foreach (DataColumn dc in _basic_data.Columns)
                {
                    if (result.total_analysis.Rows.Contains(dc.ColumnName))
                    {
                        DataRow total_row = result.total_analysis.Rows.Find(dc.ColumnName);
                        decimal temp_avg = (decimal)total_row["avg"];
                        total_row["standardErr"] = (decimal)total_row["standardErr"] + (Convert.ToDecimal(dr[dc]) - temp_avg)
                            * (Convert.ToDecimal(dr[dc]) - temp_avg);

                        if (count <= PLN)
                            total_row["PLN"] = (decimal)total_row["PLN"] + Convert.ToDecimal(dr[dc]);
                        else if (count > PHN)
                            total_row["PHN"] = (decimal)total_row["PHN"] + Convert.ToDecimal(dr[dc]);
                        if ((int)total_row["objective"] == 1)
                        {
                            if (Convert.ToDouble(dr[dc]) > 0)
                            {
                                decimal temp_mark = (decimal)total_row["CorrectMark"] + Convert.ToDecimal(dr[cor_col]);

                                total_row["CorrectMark"] = temp_mark;
                                total_row["CorrectNum"] = Convert.ToDecimal(total_row["CorrectNum"]) + 1;
                            }
                            else
                            {
                                decimal temp_mark = (decimal)total_row["WrongMark"] + Convert.ToDecimal(dr[cor_col]);
                                total_row["WrongMark"] = temp_mark;
                                total_row["WrongNum"] = Convert.ToDecimal(total_row["WrongNum"]) + 1;
                            }
                        }
                        else
                        {
                            decimal temp_mark = (decimal)total_row["MultipleSum"] + Convert.ToDecimal(dr[cor_col]) * Convert.ToDecimal(dr[dc]);
                            total_row["MultipleSum"] = temp_mark;
                            temp_mark = (decimal)total_row["SquareSumX"] + Convert.ToDecimal(dr[dc]) * Convert.ToDecimal(dr[dc]);
                            total_row["SquareSumX"] = temp_mark;
                        }
                        if (oddeven.IsMatch(dc.ColumnName))
                        {
                            int topic = Convert.ToInt16(dc.ColumnName.Substring(1));
                            if (topic % 2 != 0)
                                odd += (decimal)dr[dc];
                            else
                                even += (decimal)dr[dc];

                        }

                    }

                }
                //alfaMultiplySum += odd * even;
                alfaSquareX += odd * odd;
                alfaSquareY += even * even;
                alfaSumX += odd;
                alfaSumY += even;
            }
            result.skewness = (ThreeMulti / result.total_num) / Convert.ToDecimal(Math.Pow(Convert.ToDouble(Var / result.total_num), 1.5));
            result.kertosis = (FourMulti / result.total_num) / Convert.ToDecimal(Math.Pow(Convert.ToDouble(Var / result.total_num), 2.0)) - 3m;
            decimal oddDev = alfaSquareX / result.total_num - (alfaSumX / result.total_num) * (alfaSumX / result.total_num);
            decimal evenDev = alfaSquareY / result.total_num - (alfaSumY / result.total_num) * (alfaSumY / result.total_num);
            //decimal alfaNumerator = alfaMultiplySum - alfaSumX * alfaSumY / result.total_num;
            //decimal alfaPart1 = alfaSquareX - alfaSumX * alfaSumX / result.total_num;
            //decimal alfaPart2 = alfaSquareY - alfaSumY * alfaSumY / result.total_num;
            //result.alfa = alfaNumerator / Convert.ToDecimal(Math.Sqrt(Convert.ToDouble(alfaPart1 * alfaPart2)));
            result.alfa = 2 * (1 - (oddDev + evenDev) / (Var / result.total_num));
            result.stDev = Convert.ToDecimal(Math.Sqrt(Convert.ToDouble(Var / result.total_num)));
            decimal ZH_stDev = result.stDev;
            if (!isZonghe)
                ZH_stDev = Convert.ToDecimal(Math.Sqrt(Convert.ToDouble(ZH_var / result.total_num)));
            decimal part2 = SquareSumY - ZH_avg * ZH_avg * result.total_num;
            foreach (DataRow dr in result.total_analysis.Rows)
            {
                decimal temp = Convert.ToDecimal(Math.Sqrt(Convert.ToDouble((decimal)dr["standardErr"] / result.total_num)));
                dr["standardErr"] = temp;
                if ((decimal)dr["avg"] == 0)
                    dr["dfactor"] = 0;
                else
                    dr["dfactor"] = temp / (decimal)dr["avg"];
                dr["difficulty"] = (decimal)dr["avg"] / (decimal)dr["fullmark"];
                if ((int)dr["objective"] == 1)
                {
                    decimal p = (decimal)dr["CorrectNum"] / result.total_num;
                    decimal q = (decimal)dr["WrongNum"] / result.total_num;
                    if (p == 0 || q == 0)
                        dr["correlation"] = 0;
                    else
                        dr["correlation"] = (((decimal)dr["CorrectMark"] / (decimal)dr["CorrectNum"]) - ((decimal)dr["WrongMark"] / (decimal)dr["WrongNum"])) * Convert.ToDecimal(Math.Sqrt(Convert.ToDouble(p * q))) / ZH_stDev;
                }
                else
                {
                    decimal numerator = (decimal)dr["MultipleSum"] - ZH_avg * (decimal)dr["avg"] * result.total_num;
                    decimal part1 = (decimal)dr["SquareSumX"] - (decimal)dr["avg"] * (decimal)dr["avg"] * result.total_num;
                    if (part1 == 0)
                        dr["correlation"] = 0;
                    else
                        dr["correlation"] = numerator / Convert.ToDecimal(Math.Sqrt(Convert.ToDouble(part1 * part2)));
                }
                dr["discriminant"] = (((decimal)dr["PHN"] - (decimal)dr["PLN"]) / PLN) / (decimal)dr["fullmark"];

            }
            //sw.Stop();
            //string time = sw.ElapsedMilliseconds.ToString();
            //decimal part2 = 0;
            //decimal ZH_stDev = 0;
            #endregion
            #region frequency table process
            var freq = from row in _basic_data.AsEnumerable()
                       group row by row.Field<decimal>(totalmark_str) into grp
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
                    DataRow inter_row = result.totalmark_dist.NewRow();
                    inter_row["mark"] = flag;
                    inter_row["num"] = 0;
                    flag += total_interval;
                    result.totalmark_dist.Rows.Add(inter_row);
                }
                if ((result.fullmark - 20.0m * total_interval) != 0)
                {
                    DataRow last_row = result.totalmark_dist.NewRow();
                    last_row["mark"] = 20.0m * total_interval + (result.fullmark - 20.0m * total_interval + 1) / 2.0m;
                    last_row["num"] = 0;
                    result.totalmark_dist.Rows.Add(last_row);
                }
            }
            else
            {
                int j = 0;
                for (j = 0; j < result.fullmark; j++)
                {
                    DataRow inter_row = result.totalmark_dist.NewRow();
                    inter_row["mark"] = Convert.ToDecimal(j + 1);
                    inter_row["num"] = 0;
                    result.totalmark_dist.Rows.Add(inter_row);
                }
            }

            //if (first_interval == 0.0m)
            //    flag = result.fullmark - total_interval + 1;
            //else
            //    flag = result.fullmark - first_interval + 1;

            //int last_freq = 0;
            //decimal last_mark = result.fullmark + 1;

            int dist_num = 0;
            foreach (var item in freq)
            {
                DataRow dr = result.frequency_dist.NewRow();
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
                    result.totalmark_dist.Rows[20]["num"] = (int)result.totalmark_dist.Rows[20]["num"] + (int)dr["frequency"];
                else if (dist_num == 0)
                    result.totalmark_dist.Rows[dist_num]["num"] = (int)result.totalmark_dist.Rows[dist_num]["num"] + (int)dr["frequency"];
                else
                    result.totalmark_dist.Rows[dist_num - 1]["num"] = (int)result.totalmark_dist.Rows[dist_num - 1]["num"] + (int)dr["frequency"];

                if (midCheck && (int)dr["accumulateFreq"] >= mid)
                {

                    DataRow midRow = result.frequency_dist.Rows[result.frequency_dist.Rows.Count - 1];
                    if ((int)dr["frequency"] == 1)
                        if (isEven)
                            result.mean = ((decimal)dr["totalmark"] + ( decimal)midRow["totalmark"]) / 2;
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
                    midCheck = false;
                }
                
                result.frequency_dist.Rows.Add(dr);
            }
            DataTable new_freq = result.frequency_dist.Clone();
            new_freq.PrimaryKey = new DataColumn[] { new_freq.Columns["totalmark"] };
            foreach (DataRow dr in result.frequency_dist.Rows)
            {
                decimal keyMark = Convert.ToDecimal(Math.Floor(Convert.ToDouble(dr["totalmark"])));

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
                        result.mode = (decimal)oldrow["totalmark"];
                    }
                    
                }
                if (MaxFreq < (int)dr["frequency"])
                {
                    MaxFreq = (int)dr["frequency"];
                    result.mode = (decimal)dr["totalmark"];
                }
            }
            result.frequency_dist = new_freq;

            #endregion

            Total_tuple_analysis(result, totalmark_str);
            #region groups table process
            int group_dc;
            string spattern = "^\\d+~\\d+$";

            //边界问题
            int ans_count = 0;
            for (group_dc = 0; group_dc < _groups_table.Columns.Count; group_dc++)
            {
                if (_groups_table.Columns[group_dc].ColumnName.StartsWith("FZ"))
                {
                    DataRow groups_row = result.group_analysis.NewRow();
                    groups_row["number"] = _groups_table.Columns[group_dc].ColumnName;
                    groups_row["max"] = _groups_table.Compute("Max([" + _groups_table.Columns[group_dc].ColumnName + "])", "");
                    groups_row["min"] = _groups_table.Compute("Min([" + _groups_table.Columns[group_dc].ColumnName + "])", "");
                    groups_row["avg"] = _groups_table.Compute("Avg([" + _groups_table.Columns[group_dc].ColumnName + "])", "");
                    groups_row["standardErr"] = 0.0;
                    groups_row["dfactor"] = 0.0;
                    groups_row["difficulty"] = 0.0;
                    groups_row["correlation"] = 0.0;
                    groups_row["discriminant"] = 0.0;
                    groups_row["PHN"] = 0.0;
                    groups_row["PLN"] = 0.0;
                    groups_row["MultipleSum"] = 0.0;
                    groups_row["SquareSumX"] = 0.0;
                    groups_row["fullmark"] = 0.0;

                    if (groups_row["number"].ToString().Equals("生物") || groups_row["number"].ToString().Equals("政治"))
                    {
                        groups_row["fullmark"] = _config.shengwu_zhengzhi;
                    }
                    else if (groups_row["number"].ToString().Equals("物理") || groups_row["number"].ToString().Equals("历史"))
                    {
                        groups_row["fullmark"] = _config.wuli_lishi;
                    }
                    else if (groups_row["number"].ToString().Equals("化学") || groups_row["number"].ToString().Equals("地理"))
                    {
                        groups_row["fullmark"] = _config.huaxue_dili;
                    }
                    else
                    {
                        string org = _groups_ans.Rows[ans_count][1].ToString().Trim();
                        string[] org_char = org.Split(new char[2] { ',', '，' });
                        foreach (string th in org_char)
                        {

                            if (System.Text.RegularExpressions.Regex.IsMatch(th, spattern))
                            //if(th.Contains('~'))
                            {
                                string[] num = th.Split('~');
                                int j;
                                int size = Convert.ToInt32(num[0]) < Convert.ToInt32(num[1]) ? Convert.ToInt32(num[1]) : Convert.ToInt32(num[0]);
                                int start = Convert.ToInt32(num[0]) > Convert.ToInt32(num[1]) ? Convert.ToInt32(num[1]) : Convert.ToInt32(num[0]);
                                //此处需判断size和start的边界问题
                                for (j = start; j < size + 1; j++)
                                {
                                    DataRow dr = result.total_analysis.Rows.Find("t" + j.ToString());
                                    groups_row["fullmark"] = (decimal)groups_row["fullmark"] + (decimal)dr["fullmark"];
                                }

                            }
                            else
                            {
                                DataRow dr = result.total_analysis.Rows.Find("t" + th.Trim());
                                groups_row["fullmark"] = (decimal)groups_row["fullmark"] + (decimal)dr["fullmark"];
                            }
                        }
                    }
                    ans_count++;
                    result.group_analysis.Rows.Add(groups_row);
                }
            }
            count = 0;
            foreach (DataRow dr in _groups_table.Rows)
            {
                count++;
                foreach (DataColumn dc in _groups_table.Columns)
                {
                    if (result.group_analysis.Rows.Contains(dc.ColumnName))
                    {
                        DataRow groups_dr = result.group_analysis.Rows.Find(dc.ColumnName);
                        decimal temp = (decimal)groups_dr["avg"];
                        groups_dr["standardErr"] = (decimal)groups_dr["standardErr"] + (Convert.ToDecimal(dr[dc]) - temp) * (Convert.ToDecimal(dr[dc]) - temp);
                        if (count <= PLN)
                            groups_dr["PLN"] = (decimal)groups_dr["PLN"] + Convert.ToDecimal(dr[dc]);
                        else if (count >= PHN)
                            groups_dr["PHN"] = (decimal)groups_dr["PHN"] + Convert.ToDecimal(dr[dc]);

                        decimal temp_mark = (decimal)groups_dr["MultipleSum"] + Convert.ToDecimal(_basic_data.Rows[count - 1][cor_col]) * Convert.ToDecimal(dr[dc]);
                        groups_dr["MultipleSum"] = temp_mark;
                        temp_mark = (decimal)groups_dr["SquareSumX"] + Convert.ToDecimal(dr[dc]) * Convert.ToDecimal(dr[dc]);
                        groups_dr["SquareSumX"] = temp_mark;
                    }
                }
            }
            foreach (DataRow dr in result.group_analysis.Rows)
            {
                decimal temp = Convert.ToDecimal(Math.Sqrt(Convert.ToDouble((decimal)dr["standardErr"] / result.total_num)));
                dr["standardErr"] = temp;
                if ((decimal)dr["avg"] == 0)
                    dr["dfactor"] = 0;
                else
                    dr["dfactor"] = temp / (decimal)dr["avg"];
                dr["difficulty"] = (decimal)dr["avg"] / (decimal)dr["fullmark"];
                decimal numerator = (decimal)dr["MultipleSum"] - ZH_avg * (decimal)dr["avg"] * result.total_num;
                decimal part1 = (decimal)dr["SquareSumX"] - (decimal)dr["avg"] * (decimal)dr["avg"] * result.total_num;
                if (part1 == 0 || part2 == 0)
                    dr["correlation"] = 0;
                else
                    dr["correlation"] = numerator / Convert.ToDecimal(Math.Sqrt(Convert.ToDouble(part1 * part2)));
                dr["discriminant"] = (((decimal)dr["PHN"] - (decimal)dr["PLN"]) / PLN) / (decimal)dr["fullmark"];
            }
            #endregion

            

            
            result.Dfactor = result.stDev / result.avg;
            result.difficulty = result.avg / result.fullmark;
            result.standardErr = result.stDev / Convert.ToDecimal(Math.Sqrt(Convert.ToDouble(result.total_num)));

            #region  single group detail analysis
            int i;
            foreach (DataRow dr in result.group_analysis.Rows)
            {
                WordData.group_data temp = new WordData.group_data();

                temp.group_dist.Columns.Add("mark", typeof(decimal));
                temp.group_dist.Columns.Add("num", typeof(int));

                decimal interval = Math.Ceiling((decimal)dr["fullmark"] / 20.0m);
                int tuple_num = Convert.ToInt32(Math.Floor((decimal)dr["fullmark"] / interval));
                
                    flag = (interval + 1) / 2.0m;

                    DataRow inter_row = temp.group_dist.NewRow();
                    inter_row["mark"] = 0;
                    inter_row["num"] = 0;
                    temp.group_dist.Rows.Add(inter_row);

                    int j = 0;
                    for (j = 0; j < tuple_num; j++)
                    {
                        inter_row = temp.group_dist.NewRow();
                        inter_row["mark"] = flag;
                        inter_row["num"] = 0;
                        flag += interval;
                        temp.group_dist.Rows.Add(inter_row);
                    }
                    if (((decimal)dr["fullmark"] - tuple_num * interval) != 0)
                    {
                        DataRow last_row = temp.group_dist.NewRow();
                        last_row["mark"] = tuple_num * interval + 1 + ((decimal)dr["fullmark"] - tuple_num * interval - 1) / 2.0m;
                        last_row["num"] = 0;
                        temp.group_dist.Rows.Add(last_row);
                    }
                    
                
               
                var single_group = from row in _groups_table.AsEnumerable()
                                   group row by row.Field<decimal>(dr["number"].ToString()) into grp
                                   orderby grp.Key ascending
                                   select new
                                   {
                                       totalmark = grp.Key,
                                       count = grp.Count(),
                                       average = grp.Average(row => row.Field<decimal>(cor_col)),
                                       
                                   };
                
                temp.group_detail.Columns.Add("mark", typeof(string));
                for (i = 1; i <= _groupnum; i++)
                    temp.group_detail.Columns.Add("G" + i.ToString(), typeof(decimal));
                temp.group_detail.Columns.Add("frequency", typeof(int));
                temp.group_detail.Columns.Add("rate", typeof(decimal));
                temp.group_detail.Columns.Add("avg", typeof(decimal));
                temp.group_detail.PrimaryKey = new DataColumn[] { temp.group_detail.Columns["mark"] };

                foreach (var item in single_group)
                {
                    if (!temp.group_detail.Rows.Contains(Convert.ToInt32(Math.Ceiling(item.totalmark)).ToString() + "～"))
                    {
                        DataRow newrow = temp.group_detail.NewRow();
                        newrow["mark"] = Convert.ToInt32(Math.Ceiling(item.totalmark)).ToString() + "～";
                        newrow["frequency"] = item.count;
                        newrow["rate"] = 0;
                        newrow["avg"] = item.average * item.count;
                        for (i = 1; i <= _groupnum; i++)
                            newrow["G" + i.ToString()] = 0.0m;
                        temp.group_detail.Rows.Add(newrow);
                    }
                    else
                    {
                        DataRow oldrow = temp.group_detail.Rows.Find(Convert.ToInt32(Math.Ceiling(item.totalmark)).ToString() + "～");
                        oldrow["frequency"] = (int)oldrow["frequency"] + item.count;
                        oldrow["avg"] = (decimal)oldrow["avg"] + item.average * item.count;
                    }
                    dist_num = Convert.ToInt32(Math.Ceiling(item.totalmark / interval));

                    
                    //if (dist_num == 0)
                        temp.group_dist.Rows[dist_num]["num"] = (int)temp.group_dist.Rows[dist_num]["num"] + item.count;
                    //else
                    //    temp.group_dist.Rows[dist_num - 1]["num"] = (int)temp.group_dist.Rows[dist_num - 1]["num"] + item.count;
                   


                }
                foreach (DataRow row in temp.group_detail.Rows)
                {
                    row["rate"] = ((int)row["frequency"] / Convert.ToDecimal(result.total_num)) * 100;
                    row["avg"] = (decimal)row["avg"] / (int)row["frequency"];
                }
                var tuple = from row in _groups_table.AsEnumerable()
                            group row by new
                            {
                                name = row.Field<decimal>(dr["number"].ToString()),
                                groups = row.Field<string>("Groups"),
                            } into grp
                            orderby grp.Key.name ascending
                            select new
                                   {
                                       totalmark = grp.Key.name,
                                       groups = grp.Key.groups,
                                       count = grp.Count()
                                   };

                foreach (var item in tuple)
                {
                    DataRow tuple_row = temp.group_detail.Rows.Find(Convert.ToInt32(Math.Ceiling(item.totalmark)).ToString() + "～");
                    tuple_row[item.groups.ToString().Trim()] = (decimal)tuple_row[item.groups.ToString().Trim()] + Convert.ToDecimal(item.count);
                }

                var difficulty = from row in _groups_table.AsEnumerable()
                                 group row by row.Field<decimal>(cor_col) into grp
                                 orderby grp.Key ascending
                                 select new
                                 {
                                     totalmark = grp.Key,
                                     diff = grp.Average(row => row.Field<decimal>(dr["number"].ToString()))
                                 };
                temp.group_difficulty.Columns.Add("totalmark", typeof(decimal));
                temp.group_difficulty.Columns.Add("difficulty", typeof(decimal));

                foreach (var item in difficulty)
                {
                    DataRow diff_row = temp.group_difficulty.NewRow();
                    diff_row["totalmark"] = item.totalmark;
                    diff_row["difficulty"] = item.diff / (decimal)dr["fullmark"];
                    temp.group_difficulty.Rows.Add(diff_row);
                }

                var gdata = from row in _groups_table.AsEnumerable()
                            group row by row.Field<string>("Groups") into grp
                            select new
                                {
                                    gtype = grp.Key,
                                    count = grp.Count(),
                                    avg = grp.Average(row => row.Field<decimal>(dr["number"].ToString()))
                                };
                DataRow total = temp.group_detail.NewRow();
                DataRow avg = temp.group_detail.NewRow();
                for (int k = 1; k <= _groupnum; k++)
                {
                    total["G" + k.ToString()] = 0;
                    avg["G" + k.ToString()] = 0;
                }

                foreach (var item in gdata)
                {
                    total[item.gtype.ToString().Trim()] = item.count;
                    avg[item.gtype.ToString().Trim()] = item.avg / (decimal)dr["fullmark"];
                }
                total["mark"] = "合计";
                total["frequency"] = result.total_num;
                total["rate"] = 100.0m;
                total["avg"] = ZH_avg;

                avg["mark"] = "得分率";
                avg["frequency"] = 0;
                avg["rate"] = 0.0m;
                avg["avg"] = 0.0m;

                temp.group_detail.Rows.Add(total);
                temp.group_detail.Rows.Add(avg);

                result.single_group_analysis.Add(temp);

            }
            #endregion
            if (!isZonghe)
            {
                #region single topic analysis
                string[] single_topic = { "1", "2", "4", "8", "P", "p" };
                int topic_count = 0;

                foreach (DataRow dr in result.total_analysis.Rows)
                {
                    WordData.single_data temp = new WordData.single_data();
                    temp.single_difficulty.Columns.Add("totalmark", typeof(decimal));
                    temp.single_difficulty.Columns.Add("difficulty", typeof(decimal));

                    var diff = from row in _basic_data.AsEnumerable()
                               group row by row.Field<decimal>(cor_col) into grp
                               orderby grp.Key ascending
                               select new
                               {
                                   totalmark = grp.Key,
                                   avg = grp.Average(row => row.Field<decimal>(dr["number"].ToString().Trim()))
                               };

                    foreach (var item in diff)
                    {
                        DataRow diff_dr = temp.single_difficulty.NewRow();
                        diff_dr["totalmark"] = item.totalmark;
                        diff_dr["difficulty"] = item.avg / (decimal)dr["fullmark"];
                        temp.single_difficulty.Rows.Add(diff_dr);
                    }


                    if (!_standard_ans.Rows.Find(dr["number"].ToString().Substring(1))["da"].Equals(""))
                    {
                        temp.single_detail.Columns.Add("mark", typeof(string));
                        for (i = 1; i <= _groupnum; i++)
                            temp.single_detail.Columns.Add("G" + i.ToString().Trim(), typeof(decimal));
                        temp.single_detail.Columns.Add("frequency", typeof(int));
                        temp.single_detail.Columns.Add("rate", typeof(decimal));
                        temp.single_detail.Columns.Add("correlation", typeof(decimal));
                        temp.single_detail.Columns.Add("avg", typeof(decimal));

                        temp.single_detail.PrimaryKey = new DataColumn[] { temp.single_detail.Columns["mark"] };

                        var single_avg = from row in _basic_data.AsEnumerable()
                                         group row by row.Field<string>("D" + dr["number"].ToString().Substring(1)) into grp
                                         select new
                                         {
                                             choice = grp.Key,
                                             count = grp.Count(),
                                             avg = grp.Average(row => row.Field<decimal>(cor_col)),
                                             //var = grp.Average(row => row.Field<decimal>("totalmark") * row.Field<decimal>("totalmark"))
                                         };
                        foreach (var item in single_avg)
                        {
                            DataRow single_row = temp.single_detail.NewRow();
                            single_row["mark"] = choiceTransfer(item.choice.ToString());
                            single_row["frequency"] = item.count;
                            single_row["rate"] = item.count / Convert.ToDecimal(result.total_num) * 100;
                            single_row["avg"] = item.avg;

                            single_row["correlation"] = correlation(item.count, item.avg, ZH_avg, result.total_num, ZH_stDev);

                            for (i = 1; i <= _groupnum; i++)
                            {
                                single_row["G" + i.ToString().Trim()] = 0m;
                            }

                            temp.single_detail.Rows.Add(single_row);

                        }

                        DataView dv = temp.single_detail.DefaultView;
                        dv.Sort = "mark";
                        temp.single_detail = dv.ToTable();
                        temp.single_detail.PrimaryKey = new DataColumn[] { temp.single_detail.Columns["mark"] };


                        var groups = from row in _basic_data.AsEnumerable()
                                     group row by new
                                     {
                                         groups = row.Field<string>("Groups"),
                                         choice = row.Field<string>("D" + dr["number"].ToString().Substring(1))
                                     } into grp
                                     select new
                                     {
                                         groups = grp.Key.groups,
                                         choice = grp.Key.choice,
                                         count = grp.Count(),

                                     };
                        foreach (var item in groups)
                        {
                            DataRow groups_row = temp.single_detail.Rows.Find(choiceTransfer(item.choice.ToString()));
                            groups_row[item.groups.ToString().Trim()] = Convert.ToDecimal(item.count);
                        }

                        var vertical = from row in _basic_data.AsEnumerable()
                                       group row by row.Field<string>("Groups") into grp
                                       select new
                                       {
                                           groups = grp.Key,
                                           count = grp.Count(),
                                           avg = grp.Average(row => row.Field<decimal>(dr["number"].ToString().Trim()))
                                       };
                        DataRow single_total_row = temp.single_detail.NewRow();
                        DataRow single_avg_row = temp.single_detail.NewRow();
                        single_total_row["mark"] = "合计";
                        single_avg_row["mark"] = "得分率";

                        for (int k = 1; k <= _groupnum; k++)
                        {
                            single_total_row["G" + k.ToString()] = 0;
                            single_avg_row["G" + k.ToString()] = 0;
                        }
                        foreach (var item in vertical)
                        {
                            single_total_row[item.groups.ToString().Trim()] = Convert.ToDecimal(item.count);
                            single_avg_row[item.groups.ToString().Trim()] = item.avg / (decimal)dr["fullmark"];
                        }
                        single_total_row["frequency"] = result.total_num;
                        single_total_row["rate"] = 100.0m;
                        single_total_row["correlation"] = 0m;

                        single_avg_row["frequency"] = 0;
                        single_avg_row["rate"] = 0m;
                        single_avg_row["correlation"] = 0m;

                        temp.single_detail.Rows.Add(single_total_row);
                        temp.single_detail.Rows.Add(single_avg_row);



                        if (_standard_ans.Rows.Find(dr["number"].ToString().Substring(1))["da"].ToString().Trim().Length == 1)
                        {


                            temp.stype = WordData.single_type.single;

                            DataTable _single_detail = temp.single_detail.Clone();
                            insertRow(temp.single_detail.Rows.Find("合计"), _single_detail, 0);
                            insertRow(temp.single_detail.Rows.Find("得分率"), _single_detail, 1);

                            temp.single_detail.Rows.Find("合计").Delete();
                            temp.single_detail.Rows.Find("得分率").Delete();
                            if (temp.single_detail.Rows.Contains("G"))
                            {
                                insertRow(temp.single_detail.Rows.Find("G"), _single_detail, 0);
                                temp.single_detail.Rows.Find("G").Delete();
                            }
                            if (temp.single_detail.Rows.Contains("F"))
                            {
                                insertRow(temp.single_detail.Rows.Find("F"), _single_detail, 0);
                                temp.single_detail.Rows.Find("F").Delete();
                            }
                            if (temp.single_detail.Rows.Contains("E"))
                            {
                                insertRow(temp.single_detail.Rows.Find("E"), _single_detail, 0);
                                temp.single_detail.Rows.Find("E").Delete();
                            }
                            if (temp.single_detail.Rows.Contains("D"))
                            {
                                insertRow(temp.single_detail.Rows.Find("D"), _single_detail, 0);
                                temp.single_detail.Rows.Find("D").Delete();
                            }
                            if (temp.single_detail.Rows.Contains("C"))
                            {
                                insertRow(temp.single_detail.Rows.Find("C"), _single_detail, 0);
                                temp.single_detail.Rows.Find("C").Delete();
                            }
                            if (temp.single_detail.Rows.Contains("B"))
                            {
                                insertRow(temp.single_detail.Rows.Find("B"), _single_detail, 0);
                                temp.single_detail.Rows.Find("B").Delete();
                            }
                            if (temp.single_detail.Rows.Contains("A"))
                            {
                                insertRow(temp.single_detail.Rows.Find("A"), _single_detail, 0);
                                temp.single_detail.Rows.Find("A").Delete();
                            }
                            temp.single_detail.AcceptChanges();
                            DataRow nochoice_row = _single_detail.NewRow();
                            nochoice_row["mark"] = "未选或多选";
                            for (i = 1; i <= _groupnum; i++)
                                nochoice_row["G" + i.ToString().Trim()] = 0m;
                            nochoice_row["frequency"] = 0;
                            nochoice_row["rate"] = 0m;
                            nochoice_row["correlation"] = 0m;
                            nochoice_row["avg"] = 0m;
                            foreach (DataRow temp_dr in temp.single_detail.Rows)
                            {
                                nochoice_row["avg"] = (decimal)nochoice_row["avg"] + (decimal)temp_dr["avg"] * (int)temp_dr["frequency"];
                                nochoice_row["frequency"] = (int)nochoice_row["frequency"] + (int)temp_dr["frequency"];
                                for (i = 1; i <= _groupnum; i++)
                                    nochoice_row["G" + i.ToString().Trim()] = (decimal)nochoice_row["G" + i.ToString().Trim()] + (decimal)temp_dr["G" + i.ToString().Trim()];

                            }
                            nochoice_row["rate"] = (int)nochoice_row["frequency"] / Convert.ToDecimal(result.total_num) * 100m;
                            if ((int)nochoice_row["frequency"] != 0)
                                nochoice_row["avg"] = (decimal)nochoice_row["avg"] / (int)nochoice_row["frequency"];
                            else
                                nochoice_row["avg"] = 0;
                            nochoice_row["correlation"] = correlation((int)nochoice_row["frequency"], (decimal)nochoice_row["avg"], ZH_avg, result.total_num, ZH_stDev);
                            _single_detail.Rows.InsertAt(nochoice_row, _single_detail.Rows.Count - 2);
                            temp.single_detail = _single_detail;


                            temp.single_dist.Columns.Add("groups", typeof(string));
                            foreach (DataRow detail_row in temp.single_detail.Rows)
                            {
                                if (!(detail_row["mark"].ToString().Trim().Equals("未选或多选") ||
                                     detail_row["mark"].ToString().Trim().Equals("合计") ||
                                     detail_row["mark"].ToString().Trim().Equals("得分率")))
                                    temp.single_dist.Columns.Add(detail_row["mark"].ToString().Trim(), typeof(decimal));
                            }
                            temp.single_dist.PrimaryKey = new DataColumn[] { temp.single_dist.Columns["groups"] };

                            for (i = 1; i <= _groupnum; i++)
                            {
                                DataRow dist_row = temp.single_dist.NewRow();
                                dist_row["groups"] = "G" + i.ToString().Trim();
                                decimal total = (decimal)temp.single_detail.Rows.Find("合计")["G" + i.ToString().Trim()];
                                foreach (DataRow temp_dr in temp.single_detail.Rows)
                                {
                                    if (!(temp_dr["mark"].ToString().Trim().Equals("未选或多选") ||
                                     temp_dr["mark"].ToString().Trim().Equals("合计") ||
                                     temp_dr["mark"].ToString().Trim().Equals("得分率")))
                                        if (total != 0)
                                            dist_row[temp_dr["mark"].ToString().Trim()] = (decimal)temp_dr["G" + i.ToString().Trim()] / total;
                                        else
                                            dist_row[temp_dr["mark"].ToString().Trim()] = 0;
                                }

                                temp.single_dist.Rows.Add(dist_row);
                            }


                        }
                        else
                        {
                            temp.stype = WordData.single_type.multiple;
                            temp.single_dist.Columns.Add("groups", typeof(string));
                            temp.single_dist.Columns.Add("difficulty", typeof(decimal));

                            for (i = 1; i <= _groupnum; i++)
                            {
                                DataRow temp_dr = temp.single_dist.NewRow();
                                temp_dr["groups"] = "G" + i.ToString().Trim();
                                temp_dr["difficulty"] = temp.single_detail.Rows.Find("得分率")["G" + i.ToString().Trim()];
                                temp.single_dist.Rows.Add(temp_dr);
                            }

                        }

                        DataRow ans_row = temp.single_detail.Rows.Find(choiceTransfer(_standard_ans.Rows.Find(dr["number"].ToString().Substring(1))["da"].ToString().Trim()));
                        if(ans_row != null)
                            ans_row["mark"] = "*" + ans_row["mark"];
                        temp.single_detail.Columns.Remove(temp.single_detail.Columns["avg"]);
                        temp.single_detail.AcceptChanges();

                    }
                    else
                    {
                        temp.stype = WordData.single_type.sub;
                        temp.single_detail.Columns.Add("mark", typeof(string));
                        for (i = 1; i <= _groupnum; i++)
                            temp.single_detail.Columns.Add("G" + i.ToString().Trim(), typeof(decimal));
                        temp.single_detail.Columns.Add("frequency", typeof(int));
                        temp.single_detail.Columns.Add("rate", typeof(decimal));
                        temp.single_detail.Columns.Add("correlation", typeof(decimal));

                        temp.single_detail.PrimaryKey = new DataColumn[] { temp.single_detail.Columns["mark"] };

                        var single_avg = from row in _basic_data.AsEnumerable()
                                         group row by row.Field<decimal>(dr["number"].ToString().Trim()) into grp
                                         orderby grp.Key ascending
                                         select new
                                         {
                                             mark = grp.Key,
                                             count = grp.Count(),
                                             avg = grp.Average(row => row.Field<decimal>(cor_col))
                                         };
                        foreach (var item in single_avg)
                        {
                            //if (!temp.single_detail.Rows.Contains(Convert.ToInt32(Math.Floor(item.mark)).ToString() + "～"))
                            if (!temp.single_detail.Rows.Contains(string.Format("{0:F1}",item.mark) + "～"))
                            {
                                DataRow temp_dr = temp.single_detail.NewRow();
                                temp_dr["mark"] = string.Format("{0:F1}", item.mark) + "～";
                                temp_dr["frequency"] = item.count;
                                temp_dr["rate"] = 0;
                                temp_dr["correlation"] = item.avg * item.count;
                                for (i = 1; i <= _groupnum; i++)
                                {
                                    temp_dr["G" + i.ToString().Trim()] = 0m;
                                }
                                temp.single_detail.Rows.Add(temp_dr);
                            }
                            else
                            {
                                DataRow oldrow = temp.single_detail.Rows.Find(string.Format("{0:F1}", item.mark) + "～");
                                oldrow["frequency"] = (int)oldrow["frequency"] + item.count;
                                oldrow["correlation"] = (decimal)oldrow["correlation"] + item.avg * item.count;
                            }
                        }
                        foreach (DataRow row in temp.single_detail.Rows)
                        {
                            row["rate"] = ((int)row["frequency"] / Convert.ToDecimal(result.total_num)) * 100;
                            row["correlation"] = (decimal)row["correlation"] / (int)row["frequency"];
                        }
                        var gdata = from row in _basic_data.AsEnumerable()
                                    group row by new
                                    {
                                        groups = row.Field<string>("Groups"),
                                        mark = row.Field<decimal>(dr["number"].ToString().Trim())
                                    } into grp
                                    select new
                                    {
                                        groups = grp.Key.groups,
                                        mark = grp.Key.mark,
                                        count = grp.Count()
                                    };
                        foreach (var item in gdata)
                        {
                            DataRow temp_dr = temp.single_detail.Rows.Find(string.Format("{0:F1}", item.mark) + "～");
                            temp_dr[item.groups.ToString().Trim()] = (decimal)temp_dr[item.groups.ToString().Trim()] + Convert.ToDecimal(item.count);

                        }
                        var vertical = from row in _basic_data.AsEnumerable()
                                       group row by row.Field<string>("Groups") into grp
                                       select new
                                       {
                                           groups = grp.Key,
                                           count = grp.Count(),
                                           avg = grp.Average(row => row.Field<decimal>(dr["number"].ToString().Trim()))
                                       };
                        DataRow total_dr = temp.single_detail.NewRow();
                        DataRow avg_dr = temp.single_detail.NewRow();

                        total_dr["mark"] = "合计";
                        total_dr["frequency"] = result.total_num;
                        total_dr["rate"] = 100.0m;
                        total_dr["correlation"] = ZH_avg;

                        avg_dr["mark"] = "得分率";
                        avg_dr["frequency"] = 0;
                        avg_dr["rate"] = 0m;
                        avg_dr["correlation"] = 0m;

                        for (int k = 1; k <= _groupnum; k++)
                        {
                            total_dr["G" + k.ToString()] = 0;
                            avg_dr["G" + k.ToString()] = 0;
                        }
                        foreach (var item in vertical)
                        {
                            total_dr[item.groups.ToString().Trim()] = Convert.ToDecimal(item.count);
                            avg_dr[item.groups.ToString().Trim()] = item.avg / (decimal)dr["fullmark"];
                        }

                        temp.single_detail.Rows.Add(total_dr);
                        temp.single_detail.Rows.Add(avg_dr);


                        temp.single_dist.Columns.Add("groups", typeof(string));
                        temp.single_dist.Columns.Add("difficulty", typeof(decimal));

                        for (i = 1; i <= _groupnum; i++)
                        {
                            DataRow temp_dr = temp.single_dist.NewRow();
                            temp_dr["groups"] = "G" + i.ToString().Trim();
                            temp_dr["difficulty"] = temp.single_detail.Rows.Find("得分率")["G" + i.ToString().Trim()];
                            temp.single_dist.Rows.Add(temp_dr);
                        }

                    }
                    topic_count++;
                    result.single_topic_analysis.Add(temp);
                }
                #endregion

                
            }
            if(!(isZonghe || _config.report_style.Equals("学校")) )
                group_correlation();
            if (isZonghe)
                group_mark(_basic_data);
            else if (!(_config.subject.Contains("理综") || _config.subject.Contains("文综")))
                group_mark(_basic_data);
            else
                zh_group_mark(_basic_data);
            return true;
        }
        public void Total_tuple_analysis(WordData wd, string totalmarkstr)
        {
            wd.Total_tuple_analysis.Columns.Add("name", typeof(string));
            wd.Total_tuple_analysis.Columns.Add("ScoreRange", typeof(string));
            wd.Total_tuple_analysis.Columns.Add("Average", typeof(decimal));
            wd.Total_tuple_analysis.Columns.Add("difficulty", typeof(decimal));

            wd.Total_tuple_analysis.PrimaryKey = new DataColumn[] { wd.Total_tuple_analysis.Columns["name"]};

            var tuples = from row in _basic_data.AsEnumerable()
                      group row by row.Field<string>("Groups") into grp
                      select new
                      {
                          name = grp.Key,
                          max = grp.Max(row => row.Field<decimal>(totalmarkstr)),
                          min = grp.Min(row => row.Field<decimal>(totalmarkstr)),
                          avg = grp.Average(row => row.Field<decimal>(totalmarkstr))
                      };
            foreach (var tuple in tuples)
            {
                DataRow dr = wd.Total_tuple_analysis.NewRow();
                dr["name"] = tuple.name.Trim();
                dr["ScoreRange"] = tuple.min + "～" + tuple.max;
                dr["Average"] = tuple.avg;
                dr["difficulty"] = tuple.avg / wd.fullmark;
                wd.Total_tuple_analysis.Rows.Add(dr);
            }

        }
        public void xz_single_postprocess(string th)
        {
            #region 选做题部分
            DataTable xz_data = new DataTable();
            List<DataTable> xz_total = new List<DataTable>();
            List<List<WordData.single_data>> xz_single = new List<List<WordData.single_data>>();
            List<string> xz_name = new List<string>();

            Utils.XZ_group_separate(_basic_data, _config, "X" + th);
            xz_data.Columns.Add("totalmark", typeof(decimal));
            xz_data.Columns.Add("X" + th, typeof(string));
            xz_data.Columns.Add("xz_groups", typeof(string));
            xz_data.Columns.Add("T" + th, typeof(decimal));
            DataRow ans_dr = _standard_ans.Rows.Find(th);

            if (!ans_dr["da"].Equals(""))
                xz_data.Columns.Add("D" + th, typeof(string));
            foreach (DataRow dr in _basic_data.Rows)
            {
                DataRow newrow = xz_data.NewRow();
                foreach (DataColumn dc in xz_data.Columns)
                    newrow[dc.ColumnName] = dr[dc.ColumnName];
                xz_data.Rows.Add(newrow);
            }
            var xz_tuple = from row in xz_data.AsEnumerable()
                           group row by row.Field<string>("X" + th) into grp
                           orderby grp.Key ascending
                           select new
                           {
                               name = grp.Key,
                               count = grp.Count()
                           };
            foreach (var item in xz_tuple)
            {
                DataView dv = xz_data.equalfilter("X"+th, item.name).DefaultView;
                dv.Sort = "totalmark";
                xz_name.Add(item.name);
                xz_group_analysis(dv.ToTable(), item.count, xz_total, xz_single);
            }

            for (int i = 0; i < xz_total[0].Rows.Count; i++)
            {
                for (int j = 0; j < xz_total.Count; j++)
                {
                    DataRow dr = xz_total[j].Rows[i];
                    dr["number"] = (string)dr["number"] + "选" + xz_name[j];
                    result.total_analysis.ImportRow(dr);
                    result.single_topic_analysis.Add(xz_single[j][i]);
                }
            }

            #endregion
        }

        public void xz_postprocess(List<string> xz_th)
        {
            foreach (string th in xz_th)
                xz_single_postprocess(th);
        }

        public void xz_group_analysis(DataTable dt, int count, List<DataTable> xz_total, List<List<WordData.single_data>> xz_single)
        {
            //dt.SeperateGroups(_config.group_type, _groupnum);
            DataTable xz_total_analysis = result.total_analysis.Clone();
            Regex number = new Regex("^[Tt]\\d");
            #region total analysis table process
            foreach (DataColumn dc in dt.Columns)
            {
                if (number.IsMatch(dc.ColumnName))
                {

                    DataRow newR = xz_total_analysis.NewRow();
                    string topic_num = dc.ColumnName.Substring(1);

                    newR["number"] = dc.ColumnName;
                    newR["fullmark"] = Convert.ToDecimal(_standard_ans.Rows.Find(topic_num)["fs"]);
                    newR["max"] = dt.Compute("Max([" + dc.ColumnName + "])", "");
                    newR["min"] = dt.Compute("Min([" + dc.ColumnName + "])", "");
                    newR["avg"] = dt.Compute("Avg([" + dc.ColumnName + "])", "");
                    newR["standardErr"] = 0.0;
                    newR["dfactor"] = 0.0;
                    newR["difficulty"] = 0.0;
                    newR["correlation"] = 0.0;
                    newR["discriminant"] = 0.0;
                    newR["PHN"] = 0.0;
                    newR["PLN"] = 0.0;
                    newR["CorrectMark"] = 0.0;
                    newR["CorrectNum"] = 0.0;
                    newR["WrongMark"] = 0.0;
                    newR["WrongNum"] = 0.0;
                    newR["MultipleSum"] = 0.0;
                    newR["SquareSumX"] = 0.0;

                    if (!_standard_ans.Rows.Find(topic_num)["da"].Equals(""))
                        newR["objective"] = 1;
                    else
                        newR["objective"] = 0;
                    xz_total_analysis.Rows.Add(newR);
                }

            }
            int count_tm = 0;
            int my_pln = Convert.ToInt32(Math.Ceiling(count * 0.27));
            int my_phn = count - my_pln;

            decimal Var = 0.0m;
            decimal avg = (decimal)dt.Compute("Avg(totalmark)", "");
            decimal SquareSumY = 0.0m;
            
            Regex oddeven = new Regex("^[Tt]\\d+$");
            foreach (DataRow dr in dt.Rows)
            {
                Var += ((decimal)dr["totalmark"] - avg) * ((decimal)dr["totalmark"] - avg);
                count_tm++;
                SquareSumY += (decimal)dr["totalmark"] * (decimal)dr["totalmark"];
                foreach (DataColumn dc in dt.Columns)
                {
                    if (xz_total_analysis.Rows.Contains(dc.ColumnName))
                    {
                        DataRow total_row = xz_total_analysis.Rows.Find(dc.ColumnName);
                        decimal temp_avg = (decimal)total_row["avg"];
                        total_row["standardErr"] = (decimal)total_row["standardErr"] + (Convert.ToDecimal(dr[dc]) - temp_avg)
                            * (Convert.ToDecimal(dr[dc]) - temp_avg);

                        if (count_tm <= my_pln)
                            total_row["PLN"] = (decimal)total_row["PLN"] + Convert.ToDecimal(dr[dc]);
                        else if (count_tm > my_phn)
                            total_row["PHN"] = (decimal)total_row["PHN"] + Convert.ToDecimal(dr[dc]);
                        if ((int)total_row["objective"] == 1)
                        {
                            if (Convert.ToDouble(dr[dc]) > 0)
                            {
                                decimal temp_mark = (decimal)total_row["CorrectMark"] + Convert.ToDecimal(dr["totalmark"]);
                                total_row["CorrectMark"] = temp_mark;
                                total_row["CorrectNum"] = Convert.ToDecimal(total_row["CorrectNum"]) + 1;
                            }
                            else
                            {
                                decimal temp_mark = (decimal)total_row["WrongMark"] + Convert.ToDecimal(dr["totalmark"]);
                                total_row["WrongMark"] = temp_mark;
                                total_row["WrongNum"] = Convert.ToDecimal(total_row["WrongNum"]) + 1;
                            }
                        }
                        else
                        {
                            decimal temp_mark = (decimal)total_row["MultipleSum"] + Convert.ToDecimal(dr["totalmark"]) * Convert.ToDecimal(dr[dc]);
                            total_row["MultipleSum"] = temp_mark;
                            temp_mark = (decimal)total_row["SquareSumX"] + Convert.ToDecimal(dr[dc]) * Convert.ToDecimal(dr[dc]);
                            total_row["SquareSumX"] = temp_mark;
                        }
                        

                    }

                }
                
            }

            decimal stDev = Convert.ToDecimal(Math.Sqrt(Convert.ToDouble(Var / count)));
            decimal part2 = SquareSumY - avg * avg * count;
            foreach (DataRow dr in xz_total_analysis.Rows)
            {
                decimal temp = 0;
                if(count != 0)
                    temp = Convert.ToDecimal(Math.Sqrt(Convert.ToDouble((decimal)dr["standardErr"] / count)));
                dr["standardErr"] = temp;
                if ((decimal)dr["avg"] == 0)
                    dr["dfactor"] = 0;
                else
                    dr["dfactor"] = temp / (decimal)dr["avg"];
                dr["difficulty"] = (decimal)dr["avg"] / (decimal)dr["fullmark"];
                if ((int)dr["objective"] == 1)
                {
                    decimal p = (decimal)dr["CorrectNum"] / count;
                    decimal q = (decimal)dr["WrongNum"] / count;
                    dr["correlation"] = (((decimal)dr["CorrectMark"] / (decimal)dr["CorrectNum"]) - ((decimal)dr["WrongMark"] / (decimal)dr["WrongNum"])) * Convert.ToDecimal(Math.Sqrt(Convert.ToDouble(p * q))) / stDev;
                }
                else
                {
                    decimal numerator = (decimal)dr["MultipleSum"] - avg * (decimal)dr["avg"] * count;
                    decimal part1 = (decimal)dr["SquareSumX"] - (decimal)dr["avg"] * (decimal)dr["avg"] * count;
                    if (part1 == 0 || part2 == 0)
                        dr["correlation"] = 0;
                    else
                        dr["correlation"] = numerator / Convert.ToDecimal(Math.Sqrt(Convert.ToDouble(part1 * part2)));
                }
                dr["discriminant"] = (((decimal)dr["PHN"] - (decimal)dr["PLN"]) / my_pln) / (decimal)dr["fullmark"];

            }
            #endregion
            xz_total.Add(xz_total_analysis);

            
            #region single topic analysis
            List<WordData.single_data> xz_table = new List<WordData.single_data>();
            string[] single_topic = { "1", "2", "4", "8", "P", "p" };

            foreach (DataRow dr in xz_total_analysis.Rows)
            {
                WordData.single_data temp = new WordData.single_data();
                temp.single_difficulty.Columns.Add("totalmark", typeof(decimal));
                temp.single_difficulty.Columns.Add("difficulty", typeof(decimal));

                var diff = from row in dt.AsEnumerable()
                           group row by row.Field<decimal>("totalmark") into grp
                           orderby grp.Key ascending
                           select new
                           {
                               totalmark = grp.Key,
                               avg = grp.Average(row => row.Field<decimal>(dr["number"].ToString().Trim()))
                           };

                foreach (var item in diff)
                {
                    DataRow diff_dr = temp.single_difficulty.NewRow();
                    diff_dr["totalmark"] = item.totalmark;
                    diff_dr["difficulty"] = item.avg / (decimal)dr["fullmark"];
                    temp.single_difficulty.Rows.Add(diff_dr);
                }


                if (!_standard_ans.Rows.Find(dr["number"].ToString().Substring(1))["da"].Equals(""))
                {
                    temp.single_detail.Columns.Add("mark", typeof(string));
                    for (int i = 1; i <= _groupnum; i++)
                        temp.single_detail.Columns.Add("G" + i.ToString().Trim(), typeof(decimal));
                    temp.single_detail.Columns.Add("frequency", typeof(int));
                    temp.single_detail.Columns.Add("rate", typeof(decimal));
                    temp.single_detail.Columns.Add("correlation", typeof(decimal));
                    temp.single_detail.Columns.Add("avg", typeof(decimal));

                    temp.single_detail.PrimaryKey = new DataColumn[] { temp.single_detail.Columns["mark"] };

                    var single_avg = from row in dt.AsEnumerable()
                                     group row by row.Field<string>("D" + dr["number"].ToString().Substring(1)) into grp
                                     select new
                                     {
                                         choice = grp.Key,
                                         count = grp.Count(),
                                         avg = grp.Average(row => row.Field<decimal>("totalmark")),
                                         //var = grp.Average(row => row.Field<decimal>("totalmark") * row.Field<decimal>("totalmark"))
                                     };
                    foreach (var item in single_avg)
                    {
                        DataRow single_row = temp.single_detail.NewRow();
                        single_row["mark"] = choiceTransfer(item.choice.ToString());
                        single_row["frequency"] = item.count;
                        single_row["rate"] = item.count / Convert.ToDecimal(count) * 100;
                        single_row["avg"] = item.avg;

                        single_row["correlation"] = correlation(item.count, item.avg, avg, count, stDev);

                        for (int i = 1; i <= _groupnum; i++)
                        {
                            single_row["G" + i.ToString().Trim()] = 0m;
                        }

                        temp.single_detail.Rows.Add(single_row);

                    }

                    DataView dv = temp.single_detail.DefaultView;
                    dv.Sort = "mark";
                    temp.single_detail = dv.ToTable();
                    temp.single_detail.PrimaryKey = new DataColumn[] { temp.single_detail.Columns["mark"] };


                    var groups = from row in dt.AsEnumerable()
                                 group row by new
                                 {
                                     groups = row.Field<string>("xz_groups"),
                                     choice = row.Field<string>("D" + dr["number"].ToString().Substring(1))
                                 } into grp
                                 select new
                                 {
                                     groups = grp.Key.groups,
                                     choice = grp.Key.choice,
                                     count = grp.Count(),

                                 };
                    foreach (var item in groups)
                    {
                        DataRow groups_row = temp.single_detail.Rows.Find(choiceTransfer(item.choice.ToString()));
                        groups_row[item.groups.ToString().Trim()] = Convert.ToDecimal(item.count);
                    }

                    var vertical = from row in dt.AsEnumerable()
                                   group row by row.Field<string>("xz_groups") into grp
                                   select new
                                   {
                                       groups = grp.Key,
                                       count = grp.Count(),
                                       avg = grp.Average(row => row.Field<decimal>(dr["number"].ToString().Trim()))
                                   };
                    DataRow single_total_row = temp.single_detail.NewRow();
                    DataRow single_avg_row = temp.single_detail.NewRow();
                    single_total_row["mark"] = "合计";
                    single_avg_row["mark"] = "得分率";
                    foreach (var item in vertical)
                    {
                        single_total_row[item.groups.ToString().Trim()] = Convert.ToDecimal(item.count);
                        single_avg_row[item.groups.ToString().Trim()] = item.avg / (decimal)dr["fullmark"];
                    }
                    single_total_row["frequency"] = count;
                    single_total_row["rate"] = 100.0m;
                    single_total_row["correlation"] = 0m;

                    single_avg_row["frequency"] = 0;
                    single_avg_row["rate"] = 0m;
                    single_avg_row["correlation"] = 0m;

                    temp.single_detail.Rows.Add(single_total_row);
                    temp.single_detail.Rows.Add(single_avg_row);



                    if (_standard_ans.Rows.Find(dr["number"].ToString().Substring(1))["da"].ToString().Trim().Length == 1)
                    {

                        temp.stype = WordData.single_type.single;

                        DataTable _single_detail = temp.single_detail.Clone();
                        insertRow(temp.single_detail.Rows.Find("合计"), _single_detail, 0);
                        insertRow(temp.single_detail.Rows.Find("得分率"), _single_detail, 1);

                        temp.single_detail.Rows.Find("合计").Delete();
                        temp.single_detail.Rows.Find("得分率").Delete();
                        if (temp.single_detail.Rows.Contains("G"))
                        {
                            insertRow(temp.single_detail.Rows.Find("G"), _single_detail, 0);
                            temp.single_detail.Rows.Find("G").Delete();
                        }
                        if (temp.single_detail.Rows.Contains("F"))
                        {
                            insertRow(temp.single_detail.Rows.Find("F"), _single_detail, 0);
                            temp.single_detail.Rows.Find("F").Delete();
                        }
                        if (temp.single_detail.Rows.Contains("E"))
                        {
                            insertRow(temp.single_detail.Rows.Find("E"), _single_detail, 0);
                            temp.single_detail.Rows.Find("E").Delete();
                        }
                        if (temp.single_detail.Rows.Contains("D"))
                        {
                            insertRow(temp.single_detail.Rows.Find("D"), _single_detail, 0);
                            temp.single_detail.Rows.Find("D").Delete();
                        }
                        if (temp.single_detail.Rows.Contains("C"))
                        {
                            insertRow(temp.single_detail.Rows.Find("C"), _single_detail, 0);
                            temp.single_detail.Rows.Find("C").Delete();
                        }
                        if (temp.single_detail.Rows.Contains("B"))
                        {
                            insertRow(temp.single_detail.Rows.Find("B"), _single_detail, 0);
                            temp.single_detail.Rows.Find("B").Delete();
                        }
                        if (temp.single_detail.Rows.Contains("A"))
                        {
                            insertRow(temp.single_detail.Rows.Find("A"), _single_detail, 0);
                            temp.single_detail.Rows.Find("A").Delete();
                        }
                        temp.single_detail.AcceptChanges();
                        DataRow nochoice_row = _single_detail.NewRow();
                        nochoice_row["mark"] = "未选或多选";
                        for (int i = 1; i <= _groupnum; i++)
                            nochoice_row["G" + i.ToString().Trim()] = 0m;
                        nochoice_row["frequency"] = 0;
                        nochoice_row["rate"] = 0m;
                        nochoice_row["correlation"] = 0m;
                        nochoice_row["avg"] = 0m;
                        foreach (DataRow temp_dr in temp.single_detail.Rows)
                        {
                            nochoice_row["avg"] = (decimal)nochoice_row["avg"] + (decimal)temp_dr["avg"] * (int)temp_dr["frequency"];
                            nochoice_row["frequency"] = (int)nochoice_row["frequency"] + (int)temp_dr["frequency"];
                            for (int i = 1; i <= _groupnum; i++)
                                nochoice_row["G" + i.ToString().Trim()] = (decimal)nochoice_row["G" + i.ToString().Trim()] + (decimal)temp_dr["G" + i.ToString().Trim()];

                        }
                        nochoice_row["rate"] = (int)nochoice_row["frequency"] / Convert.ToDecimal(count) * 100m;
                        if ((int)nochoice_row["frequency"] != 0)
                            nochoice_row["avg"] = (decimal)nochoice_row["avg"] / (int)nochoice_row["frequency"];
                        else
                            nochoice_row["avg"] = 0;
                        nochoice_row["correlation"] = correlation((int)nochoice_row["frequency"], (decimal)nochoice_row["avg"], avg, count, stDev);
                        _single_detail.Rows.InsertAt(nochoice_row, _single_detail.Rows.Count - 2);
                        temp.single_detail = _single_detail;


                        temp.single_dist.Columns.Add("groups", typeof(string));
                        foreach (DataRow detail_row in temp.single_detail.Rows)
                        {
                            if (!(detail_row["mark"].ToString().Trim().Equals("未选或多选") ||
                                 detail_row["mark"].ToString().Trim().Equals("合计") ||
                                 detail_row["mark"].ToString().Trim().Equals("得分率")))
                                temp.single_dist.Columns.Add(detail_row["mark"].ToString().Trim(), typeof(decimal));
                        }
                        temp.single_dist.PrimaryKey = new DataColumn[] { temp.single_dist.Columns["groups"] };

                        for (int i = 1; i <= _groupnum; i++)
                        {
                            DataRow dist_row = temp.single_dist.NewRow();
                            dist_row["groups"] = "G" + i.ToString().Trim();
                            decimal total = (decimal)temp.single_detail.Rows.Find("合计")["G" + i.ToString().Trim()];
                            foreach (DataRow temp_dr in temp.single_detail.Rows)
                            {
                                if (!(temp_dr["mark"].ToString().Trim().Equals("未选或多选") ||
                                 temp_dr["mark"].ToString().Trim().Equals("合计") ||
                                 temp_dr["mark"].ToString().Trim().Equals("得分率")))
                                    dist_row[temp_dr["mark"].ToString().Trim()] = (decimal)temp_dr["G" + i.ToString().Trim()] / total;
                            }

                            temp.single_dist.Rows.Add(dist_row);
                        }


                    }
                    else
                    {
                        temp.stype = WordData.single_type.multiple;
                        temp.single_dist.Columns.Add("groups", typeof(string));
                        temp.single_dist.Columns.Add("difficulty", typeof(decimal));

                        for (int i = 1; i <= _groupnum; i++)
                        {
                            DataRow temp_dr = temp.single_dist.NewRow();
                            temp_dr["groups"] = "G" + i.ToString().Trim();
                            temp_dr["difficulty"] = temp.single_detail.Rows.Find("得分率")["G" + i.ToString().Trim()];
                            temp.single_dist.Rows.Add(temp_dr);
                        }

                    }

                    DataRow ans_row = temp.single_detail.Rows.Find(choiceTransfer(_standard_ans.Rows.Find(dr["number"].ToString().Substring(1))["da"].ToString().Trim()));
                    ans_row["mark"] = "*" + ans_row["mark"];
                    temp.single_detail.Columns.Remove(temp.single_detail.Columns["avg"]);
                    temp.single_detail.AcceptChanges();

                }
                else
                {
                    temp.stype = WordData.single_type.sub;
                    temp.single_detail.Columns.Add("mark", typeof(string));
                    for (int i = 1; i <= _groupnum; i++)
                        temp.single_detail.Columns.Add("G" + i.ToString().Trim(), typeof(decimal));
                    temp.single_detail.Columns.Add("frequency", typeof(int));
                    temp.single_detail.Columns.Add("rate", typeof(decimal));
                    temp.single_detail.Columns.Add("correlation", typeof(decimal));

                    temp.single_detail.PrimaryKey = new DataColumn[] { temp.single_detail.Columns["mark"] };

                    var single_avg = from row in dt.AsEnumerable()
                                     group row by row.Field<decimal>(dr["number"].ToString().Trim()) into grp
                                     orderby grp.Key ascending
                                     select new
                                     {
                                         mark = grp.Key,
                                         count = grp.Count(),
                                         avg = grp.Average(row => row.Field<decimal>("totalmark"))
                                     };
                    foreach (var item in single_avg)
                    {
                        if (!temp.single_detail.Rows.Contains(string.Format("{0:F1}", item.mark) + "～"))
                        {
                            DataRow temp_dr = temp.single_detail.NewRow();
                            temp_dr["mark"] = string.Format("{0:F1}", item.mark) + "～";
                            temp_dr["frequency"] = item.count;
                            temp_dr["rate"] = 0;
                            temp_dr["correlation"] = item.avg * item.count;
                            for (int i = 1; i <= _groupnum; i++)
                            {
                                temp_dr["G" + i.ToString().Trim()] = 0m;
                            }
                            temp.single_detail.Rows.Add(temp_dr);
                        }
                        else
                        {
                            DataRow oldrow = temp.single_detail.Rows.Find(string.Format("{0:F1}", item.mark) + "～");
                            oldrow["frequency"] = (int)oldrow["frequency"] + item.count;
                            oldrow["correlation"] = (decimal)oldrow["correlation"] + item.avg * item.count;
                        }
                    }
                    foreach (DataRow row in temp.single_detail.Rows)
                    {
                        row["rate"] = ((int)row["frequency"] / Convert.ToDecimal(count)) * 100;
                        row["correlation"] = (decimal)row["correlation"] / (int)row["frequency"];
                    }

                    var gdata = from row in dt.AsEnumerable()
                                group row by new
                                {
                                    groups = row.Field<string>("xz_groups"),
                                    mark = row.Field<decimal>(dr["number"].ToString().Trim())
                                } into grp
                                select new
                                {
                                    groups = grp.Key.groups,
                                    mark = grp.Key.mark,
                                    count = grp.Count()
                                };
                    foreach (var item in gdata)
                    {
                        DataRow temp_dr = temp.single_detail.Rows.Find(string.Format("{0:F1}", item.mark) + "～");
                        temp_dr[item.groups.ToString().Trim()] = (decimal)temp_dr[item.groups.ToString().Trim()] + Convert.ToDecimal(item.count);

                    }
                    var vertical = from row in dt.AsEnumerable()
                                   group row by row.Field<string>("xz_groups") into grp
                                   select new
                                   {
                                       groups = grp.Key,
                                       count = grp.Count(),
                                       avg = grp.Average(row => row.Field<decimal>(dr["number"].ToString().Trim()))
                                   };
                    DataRow total_dr = temp.single_detail.NewRow();
                    DataRow avg_dr = temp.single_detail.NewRow();

                    total_dr["mark"] = "合计";
                    total_dr["frequency"] = count;
                    total_dr["rate"] = 100.0m;
                    total_dr["correlation"] = avg;

                    avg_dr["mark"] = "得分率";
                    avg_dr["frequency"] = 0;
                    avg_dr["rate"] = 0m;
                    avg_dr["correlation"] = 0m;

                    foreach (var item in vertical)
                    {
                        total_dr[item.groups.ToString().Trim()] = Convert.ToDecimal(item.count);
                        avg_dr[item.groups.ToString().Trim()] = item.avg / (decimal)dr["fullmark"];
                    }

                    temp.single_detail.Rows.Add(total_dr);
                    temp.single_detail.Rows.Add(avg_dr);


                    temp.single_dist.Columns.Add("groups", typeof(string));
                    temp.single_dist.Columns.Add("difficulty", typeof(decimal));

                    for (int i = 1; i <= _groupnum; i++)
                    {
                        DataRow temp_dr = temp.single_dist.NewRow();
                        temp_dr["groups"] = "G" + i.ToString().Trim();
                        temp_dr["difficulty"] = temp.single_detail.Rows.Find("得分率")["G" + i.ToString().Trim()];
                        temp.single_dist.Rows.Add(temp_dr);
                    }

                }
                xz_table.Add(temp);
            }
            xz_single.Add(xz_table);
            #endregion
        }

        public void group_correlation()
        {
            int group_count = 0;
            foreach (string key in result.groups_group.Keys)
            {
                result.groups_group[key].Add("totalmark");
                DataTable cor_table = new DataTable();
                cor_table.Columns.Add("name", typeof(string));
                foreach (string name in result.groups_group[key])
                {
                    cor_table.Columns.Add(name, typeof(decimal));
                }
                foreach (string name in result.groups_group[key])
                {
                    DataRow temp = cor_table.NewRow();
                    temp["name"] = name;
                    foreach (DataColumn dc in cor_table.Columns)
                    {
                        if (dc.ColumnName.Equals("name"))
                            continue;
                        else if (name.Equals(dc.ColumnName))
                            temp[dc] = 1;
                        else
                            temp[dc] = 0;
                    }
                    cor_table.Rows.Add(temp);
                }
                for (int i = 0; i < cor_table.Rows.Count; i++)
                {
                    for (int j = i + 2; j < cor_table.Columns.Count; j++)
                    {
                        string corA,corB;
                        if (((string)cor_table.Rows[i]["name"]).Equals("totalmark"))
                            corA = cor_col;
                        else
                            corA = "FZ" + (i + 1 + group_count).ToString();

                        if (cor_table.Columns[j].ColumnName.Equals("totalmark"))
                            corB = cor_col;
                        else
                            corB = "FZ" + (j + group_count).ToString();
                        decimal cor = _groups_table.CalCor(corA, corB);
                        cor_table.Rows[i][j] = cor;
                        cor_table.Rows[j - 1][i + 1] = cor;
                    }
                }
                result.group_cor.Add(cor_table);
                group_count += cor_table.Rows.Count - 1;
            }
        }

        public decimal correlation(int frequency, decimal avg, decimal total_avg, decimal num, decimal stDev)
        {
            //decimal stDev;
            //if (var2 != 0)
            //    stDev = var2 - avg * avg;
            //else
            //    stDev = result.stDev;
            if (num == 0 || num == frequency)
                return 0;
            decimal xq = (total_avg * num - avg * frequency) / (num - frequency);
            decimal p = frequency / Convert.ToDecimal(num);
            decimal q = (num - frequency) / Convert.ToDecimal(num);
            decimal right = Convert.ToDecimal(Math.Sqrt(Convert.ToDouble(p * q)));
            return ((avg - xq) / stDev) * right;
        }

        public void insertRow(DataRow insert_row, DataTable target, int pos)
        {
            DataRow dr = target.NewRow();
            dr.ItemArray = insert_row.ItemArray;
            target.Rows.InsertAt(dr, pos);
        }

        public string choiceTransfer(string choice)
        {
            Regex reg = new Regex("^[A-Za-z]+$");
            if (reg.IsMatch(choice))
                return Utils.ToSBC(choice);
            else if (choice.Trim().Equals("0"))
                return "未选";
            else
                return choice.Trim();
        }
        
        public void HK_postprocess(Analysis.HK_hierarchy HK_hierarchy)
        {
            Regex number = new Regex("^[Tt]\\d");
            HK_worddata data = (HK_worddata)result;
            _basic_data.Columns.Add("rank", typeof(string));
            _groups_table.Columns.Add("rank", typeof(string));

            for(int i = 0; i < _basic_data.Rows.Count; i++)
            {
                decimal totalmark = (decimal)_basic_data.Rows[i]["totalmark"];
                if (totalmark >= HK_hierarchy.excellent_low && totalmark <= HK_hierarchy.excellent_high)
                {
                    _basic_data.Rows[i]["rank"] = "1";//outstanding
                    _groups_table.Rows[i]["rank"] = "1";
                }
                else if (totalmark >= HK_hierarchy.well_low && totalmark < HK_hierarchy.well_high)
                {
                    _basic_data.Rows[i]["rank"] = "2";//good
                    _groups_table.Rows[i]["rank"] = "2";
                }
                else if (totalmark >= HK_hierarchy.pass_low && totalmark < HK_hierarchy.pass_high)
                {
                    _basic_data.Rows[i]["rank"] = "3";//pass
                    _groups_table.Rows[i]["rank"] = "3";
                }
                else if (totalmark >= HK_hierarchy.fail_low && totalmark < HK_hierarchy.fail_high)
                {
                    _basic_data.Rows[i]["rank"] = "4";//fail
                    _groups_table.Rows[i]["rank"] = "4";
                }
                else
                {
                    _basic_data.Rows[i]["rank"] = "-1";//error
                    _groups_table.Rows[i]["rank"] = "-1";
                }
            }

            string[] ranks = { "1", "2", "3", "4" };
            foreach (string rank in ranks)
            {
                DataTable temp = _basic_data.equalfilter("rank", rank);
                DataRow dr = data.total.NewRow();
                CalculateTotalRank(temp, dr);
                dr["difficulty"] = (decimal)dr["avg"] / data.fullmark;
                dr["percent"] = (int)dr["totalnum"] / Convert.ToDecimal(data.total_num) * 100;
                dr["rank"] = RankConverter(rank, HK_hierarchy);
                data.total.Rows.Add(dr);
            }
            DataRow qt = data.total.NewRow();
            qt["rank"] = "全体";
            qt["totalnum"] = data.total_num;
            qt["percent"] = 100.00m;
            qt["avg"] = data.avg;
            qt["stDev"] = data.stDev;
            qt["Dfactor"] = data.Dfactor;
            qt["difficulty"] = data.difficulty;
            data.total.Rows.Add(qt);
            int count = 1;
            foreach (DataColumn dc in _basic_data.Columns)
            {
                if (number.IsMatch(dc.ColumnName))
                {
                    DataRow newrow = data.total_topic_rank.NewRow();
                    var temp = from row in _basic_data.AsEnumerable()
                               group row by row.Field<string>("rank") into grp
                               select new
                               {
                                   rank = grp.Key,
                                   avg = grp.Average(row => row.Field<decimal>(dc.ColumnName))
                               };
                    foreach (var item in temp)
                    {
                        string colname = RankConverter2(item.rank);
                        if (colname != null)
                            newrow[colname] = item.avg / (decimal)data.total_analysis.Rows.Find(dc.ColumnName)["fullmark"];
                    }
                    newrow["number"] = "第" + dc.ColumnName.Substring(1) + "题";
                    data.total_topic_rank.Rows.Add(newrow);
                    count++;
                }
            }

            for (int i = 0; i < _groups_ans.Rows.Count; i++)
            {
                DataRow newrow = data.total_topic_rank.NewRow();
                string th = _groups_ans.Rows[i][0].ToString().Trim();
                decimal fm = (decimal)data.group_analysis.Rows[i]["fullmark"];
                newrow["number"] = th;
                var temp = from row in _groups_table.AsEnumerable()
                           group row by row.Field<string>("rank") into grp
                           select new
                           {
                               rank = grp.Key,
                               avg = grp.Average(row => row.Field<decimal>((string)data.group_analysis.Rows[i]["number"]))
                           };

                foreach (var item in temp)
                {
                    string colname = RankConverter2(item.rank);
                    if (colname != null)
                        newrow[colname] = item.avg / fm;
                }

                data.total_topic_rank.Rows.Add(newrow);
            }


            TotalCal totalcal = new TotalCal(_groups_table, data.total_num);

            for (int i = 0; i < data.group_analysis.Rows.Count; i++ )
            {
                string name = (string)data.group_analysis.Rows[i]["number"];
                WordData.group_data single_group = (WordData.group_data)data.single_group_analysis[i];
                DataTable group = new DataTable();
                group.Columns.Add("mark", typeof(string));
                group.Columns.Add("outstanding", typeof(int));
                group.Columns.Add("out_percent", typeof(decimal));
                group.Columns.Add("good", typeof(int));
                group.Columns.Add("good_percent", typeof(decimal));
                group.Columns.Add("pass", typeof(int));
                group.Columns.Add("pass_percent", typeof(decimal));
                group.Columns.Add("fail", typeof(int));
                group.Columns.Add("fail_percent", typeof(decimal));
                group.Columns.Add("total", typeof(int));
                group.Columns.Add("total_percent", typeof(decimal));
                group.PrimaryKey = new DataColumn[]{group.Columns["mark"]};
                for(int j = 0; j < single_group.group_detail.Rows.Count - 1; j++)
                {
                    DataRow dr = group.NewRow();
                    dr["mark"] = single_group.group_detail.Rows[j]["mark"];
                    dr["total"] = single_group.group_detail.Rows[j]["frequency"];
                    dr["total_percent"] = 100.00m;

                    for (int k = 1; k < group.Columns.Count - 2; k++)
                        dr[k] = 0;
                    group.Rows.Add(dr);
                }

                totalcal.AddTotalRow(group);


                var temp = from row in _groups_table.AsEnumerable()
                           group row by new
                           {
                               rank = row.Field<string>("rank"),
                               mark = row.Field<decimal>(name)
                           } into grp
                           orderby grp.Key.mark ascending
                           select new
                           {
                               rank = grp.Key.rank,
                               mark = grp.Key.mark,
                               count = grp.Count()
                           };
                foreach (var item in temp)
                {
                    string col = RankConverter2(item.rank);
                    DataRow dr = group.Rows.Find(Convert.ToInt32(Math.Ceiling(item.mark)).ToString() + "～");
                    dr[col] = (int)dr[col] + item.count;
                    dr[group.Columns[col].Ordinal + 1] = (int)dr[col] / Convert.ToDecimal(dr["total"]) * 100;
                }


                data.single_group_rank.Add(group);

            }
            for (int i = 0; i < data.total_analysis.Rows.Count; i++)
            {
                string name = (string)data.total_analysis.Rows[i]["number"];
                WordData.single_data single_topic = (WordData.single_data)data.single_topic_analysis[i];

                DataTable topic = new DataTable();
                topic.Columns.Add("mark", typeof(string));
                topic.Columns.Add("outstanding", typeof(int));
                topic.Columns.Add("out_percent", typeof(decimal));
                topic.Columns.Add("good", typeof(int));
                topic.Columns.Add("good_percent", typeof(decimal));
                topic.Columns.Add("pass", typeof(int));
                topic.Columns.Add("pass_percent", typeof(decimal));
                topic.Columns.Add("fail", typeof(int));
                topic.Columns.Add("fail_percent", typeof(decimal));
                topic.Columns.Add("total", typeof(int));
                topic.Columns.Add("total_percent", typeof(decimal));
                topic.PrimaryKey = new DataColumn[] { topic.Columns["mark"] };

                for (int j = 0; j < single_topic.single_detail.Rows.Count - 1; j++)
                {
                    DataRow dr = topic.NewRow();
                    dr["mark"] = single_topic.single_detail.Rows[j]["mark"];
                    dr["total"] = single_topic.single_detail.Rows[j]["frequency"];
                    dr["total_percent"] = 100.0m;

                    for (int k = 1; k < topic.Columns.Count - 2; k++)
                        dr[k] = 0;
                    topic.Rows.Add(dr);
                }

                totalcal.AddTotalRow(topic);

                if (!_standard_ans.Rows[i]["da"].Equals(""))
                {
                    var temp = from row in _basic_data.AsEnumerable()
                               group row by new
                                   {
                                       mark = row.Field<string>("D" + data.total_analysis.Rows[i]["number"].ToString().Substring(1)),
                                       rank = row.Field<string>("rank")
                                   } into grp
                               select new
                               {
                                   mark = grp.Key.mark,
                                   rank = grp.Key.rank,
                                   count = grp.Count()
                               };
                    foreach (var item in temp)
                    {
                        DataRow dr;
                        string col = RankConverter2(item.rank);
                        if (choiceTransfer((string)_standard_ans.Rows[i]["da"]).Length == 1)
                        {
                            if (item.mark.ToString().Equals(_standard_ans.Rows[i]["da"]))
                            {
                                dr = topic.Rows.Find("*" + choiceTransfer(item.mark));
                            }
                            else if (choiceTransfer(item.mark).Length > 1)
                                dr = topic.Rows.Find("未选或多选");
                            else
                                dr = topic.Rows.Find(choiceTransfer(item.mark));
                            
                        }
                        else
                        {
                            if (item.mark.ToString().Equals(_standard_ans.Rows[i]["da"]))
                            {
                                dr = topic.Rows.Find("*" + choiceTransfer(item.mark));
                            }
                            else
                                dr = topic.Rows.Find(choiceTransfer(item.mark));
                        }
                        dr[col] = (int)dr[col] + item.count;
                        dr[topic.Columns[col].Ordinal + 1] = (int)dr[col] / Convert.ToDecimal(dr["total"]) * 100;
                    }
                }
                else
                {
                    var temp = from row in _basic_data.AsEnumerable()
                               group row by new
                               {
                                   mark = row.Field<decimal>(data.total_analysis.Rows[i]["number"].ToString()),
                                   rank = row.Field<string>("rank")
                               } into grp
                               select new
                               {
                                   mark = grp.Key.mark,
                                   rank = grp.Key.rank,
                                   count = grp.Count()
                               };
                    foreach (var item in temp)
                    {
                        string col = RankConverter2(item.rank);
                        DataRow dr = topic.Rows.Find(string.Format("{0:F1}", item.mark) + "～");
                        dr[col] = (int)dr[col] + item.count;
                        dr[topic.Columns[col].Ordinal + 1] = (int)dr[col] / Convert.ToDecimal(dr["total"]) * 100;
                    }
                }

                data.single_topic_rank.Add(topic);
            }


        }
        public class TotalCal
        {
            int outstanding;
            decimal out_percent;
            int good;
            decimal good_percent;
            int pass;
            decimal pass_percent;
            int fail;
            decimal fail_percent;
            public TotalCal(DataTable dt, int totalnum)
            {
                var mark = from row in dt.AsEnumerable()
                           group row by row.Field<string>("rank") into grp
                           select new
                           {
                               rank = grp.Key,
                               count = grp.Count()
                           };
                foreach (var item in mark)
                {
                    if (item.rank == "1")
                    {
                        outstanding = item.count;
                        out_percent = outstanding / Convert.ToDecimal(totalnum) * 100;
                    }
                    else if (item.rank == "2")
                    {
                        good = item.count;
                        good_percent = good / Convert.ToDecimal(totalnum) * 100;
                    }
                    else if (item.rank == "3")
                    {
                        pass = item.count;
                        pass_percent = pass / Convert.ToDecimal(totalnum) * 100;
                    }
                    else if (item.rank == "4")
                    {
                        fail = item.count;
                        fail_percent = fail / Convert.ToDecimal(totalnum) * 100;
                    }

                }


            }

            public void AddTotalRow(DataTable dt)
            {
                DataRow dr = dt.Rows[dt.Rows.Count - 1];
                dr["outstanding"] = outstanding;
                dr["out_percent"] = out_percent;
                dr["good"] = good;
                dr["good_percent"] = good_percent;
                dr["pass"] = pass;
                dr["pass_percent"] = pass_percent;
                dr["fail"] = fail;
                dr["fail_percent"] = fail_percent;
            }
        }
        public string RankConverter2(string number)
        {
            switch (number)
            {
                case "1":
                    return "outstanding";
                case "2":
                    return "good";
                case "3":
                    return "pass";
                case "4":
                    return "fail";
                default:
                    return null;
            }
        }
        public string RankConverter(string number, Analysis.HK_hierarchy HK_hierarchy)
        {
            switch (number)
            {
                case "1":
                    return "优秀" + " (" + Convert.ToInt32(HK_hierarchy.excellent_low).ToString() + " - " + Convert.ToInt32(HK_hierarchy.excellent_high).ToString() + ")";
                case "2":
                    return "良好" + " (" + Convert.ToInt32(HK_hierarchy.well_low).ToString() + " - " + Convert.ToInt32(HK_hierarchy.well_high).ToString() + ")";
                case "3":
                    return "及格" + " (" + Convert.ToInt32(HK_hierarchy.pass_low).ToString() + " - " + Convert.ToInt32(HK_hierarchy.pass_high).ToString() + ")";
                case "4":
                    return "不及格" + " (" + Convert.ToInt32(HK_hierarchy.fail_low).ToString() + " - " + Convert.ToInt32(HK_hierarchy.fail_high).ToString() + ")";
                default:
                    return "";
            }
        }

        public void CalculateTotalRank(DataTable dt, DataRow newrow)
        {

            int total_num = dt.Rows.Count;
            decimal avg = (decimal)dt.Compute("Avg(totalmark)", "");

            newrow["totalnum"] = total_num;
            newrow["avg"] = avg;
            Partition_statistic.stdev total_stdev = new Partition_statistic.stdev(total_num, avg);
            

            foreach (DataRow dr in dt.Rows)
            {
                total_stdev.add((decimal)dr["totalmark"]);
                
            }
            newrow["stDev"] = total_stdev.get_value();
            if ((decimal)newrow["avg"] == 0)
                newrow["Dfactor"] = 0;
            else
                newrow["Dfactor"] = (decimal)newrow["stDev"] / (decimal)newrow["avg"];
        }

        public void group_mark(DataTable dt)
        {
            var mark = from row in dt.AsEnumerable()
                       group row by row.Field<string>("Groups") into grp
                       select new
                       {
                           name = grp.Key,
                           max = grp.Max(row => row.Field<decimal>("totalmark"))
                       };
            _config.GroupMark.Clear();
            foreach (var temp in mark)
            {
                _config.GroupMark.Add(temp.max);
            }
        }

        public void zh_group_mark(DataTable dt)
        {
            var mark = from row in dt.AsEnumerable()
                       group row by row.Field<string>("Groups") into grp
                       select new
                       {
                           name = grp.Key,
                           max = grp.Max(row => row.Field<decimal>(cor_col))
                       };
            _config.sub_groupMark.Clear();
            foreach (var temp in mark)
            {
                _config.sub_groupMark.Add(temp.max);
            }
        }
    }
}
