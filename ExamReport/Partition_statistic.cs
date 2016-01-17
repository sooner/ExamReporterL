using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Text.RegularExpressions;
namespace ExamReport
{
    class Partition_statistic
    {
        public PartitionData result;

        public DataTable _basic_data;
        public DataTable _groups_data;

        public Configuration _config;

        public decimal _fullmark;
        
        public int _groupnum;

        public DataTable _standard_ans;
        public DataTable _groups_ans;
        decimal ZH_avg;
        string totalmark_str;
        string _title;

        string cor_col = "totalmark";

        public Partition_statistic(string title, DataTable dt, decimal fullmark, DataTable standard_ans, DataTable groups_table, DataTable groups_ans, int groupnum)
        {
            _basic_data = dt;
            _groups_data = groups_table;
            _fullmark = fullmark;
            _groupnum = groupnum;
            _standard_ans = standard_ans;
            _groups_ans = groups_ans;
            
            _standard_ans.PrimaryKey = new DataColumn[] { _standard_ans.Columns["th"] };
            _title = title;
            

        }

        public void statistic_process(bool isZonghe)
        {
            if (_config.WSLG)
            {
                result = new WSLG_partitiondata(_title);

            }
            else
            {
                result = new PartitionData(_title);
            }

            result._standard_ans = _standard_ans;
            result._group_ans = _groups_ans;
            ArrayList stdevlist = new ArrayList();
            result.total_num = _basic_data.Rows.Count;
            result.PLN = Convert.ToInt32(Math.Ceiling(_basic_data.Rows.Count * 0.27));
            result.PHN = _basic_data.Rows.Count - result.PLN + 1;
            
            if (_basic_data.Columns.Contains("ZH_totalmark"))
                totalmark_str = "ZH_totalmark";
            else
                totalmark_str = "totalmark";

            if (!isZonghe && _config.is_sub_cor)
                cor_col = "ZH_totalmark";

            result.fullmark = _fullmark;
            result.max = (decimal)_basic_data.Compute("Max(" + totalmark_str + ")", "");
            result.min = (decimal)_basic_data.Compute("Min(" + totalmark_str + ")", "");
            result.avg = (decimal)_basic_data.Compute("Avg(" + totalmark_str + ")", "");
            ZH_avg = (decimal)_basic_data.Compute("Avg(" + cor_col + ")", "");
            stdev total_stdev = new stdev(result.total_num, result.avg);
            stdevlist.Add(total_stdev);

            result.difficulty = result.avg / result.fullmark;
            Regex number = new Regex("^[Tt]\\d+");
            result.total.Add(new PartitionData.Disc(result.total_num, result.fullmark));

            foreach (DataColumn dc in _basic_data.Columns)
            {
                if (number.IsMatch(dc.ColumnName))
                {
                    DataRow dr = result.total_analysis.NewRow();
                    string topic_num = dc.ColumnName.Substring(1);
                    dr["number"] = dc.ColumnName;
                    dr["total_num"] = result.total_num;
                    dr["fullmark"] = Convert.ToDecimal(_standard_ans.Rows.Find(topic_num)["fs"]);
                    dr["max"] = _basic_data.Compute("Max([" + dc.ColumnName + "])", "");
                    dr["min"] = _basic_data.Compute("Min([" + dc.ColumnName + "])", "");
                    dr["avg"] = _basic_data.Compute("Avg([" + dc.ColumnName + "])", "");
                    stdev single_stdev = new stdev(result.total_num, (decimal)dr["avg"]);
                    stdevlist.Add(single_stdev);
                    dr["difficulty"] = (decimal)dr["avg"] / (decimal)dr["fullmark"];
                    result.total.Add(new PartitionData.Disc(result.total_num, (decimal)dr["fullmark"]));
                    
                    result.total_analysis.Rows.Add(dr);
                }
            }

            int row = 1;
            foreach (DataRow dr in _basic_data.Rows)
            {
                ((stdev)stdevlist[0]).add((decimal)dr[totalmark_str]);
                
                if (row <= result.PLN)
                        result.total[0].AddData((decimal)dr[totalmark_str], true);
                else if (row >= result.PHN)
                        result.total[0].AddData((decimal)dr[totalmark_str], false);
                
                int CoCount = 1;
                foreach (DataColumn dc in _basic_data.Columns)
                {
                    if (number.IsMatch(dc.ColumnName))
                    {
                        ((stdev)stdevlist[CoCount]).add((decimal)dr[dc]);

                        if (row <= result.PLN)
                                result.total[CoCount].AddData((decimal)dr[dc], true);
                        else if (row >= result.PHN)
                                result.total[CoCount].AddData((decimal)dr[dc], false);
                        
                        CoCount++;
                    }
                }
                row++;
            }

            result.stDev = ((stdev)stdevlist[0]).get_value();
            result.Dfactor = result.stDev / result.avg;

            result.discriminant = result.total[0].GetAns();

            total_tuple_analysis(result, totalmark_str);
            int count = 1;
            foreach (DataRow dr in result.total_analysis.Rows)
            {
                dr["stDev"] = ((stdev)stdevlist[count]).get_value();
                if ((decimal)dr["avg"] == 0)
                    dr["dfactor"] = 0m;
                else
                    dr["dfactor"] = (decimal)dr["stDev"] / (decimal)dr["avg"];
                result.total_discriminant.Add(result.total[count].GetAns());
                count++;
            }
            //此处groups表增加列时需要更改上限

            #region group table
            ArrayList groupStdev = new ArrayList();
            int ans_count = 0;
            for (int i = 3; i < _groups_data.Columns.Count - 2; i++)
            {
                if (_groups_data.Columns[i].ColumnName.StartsWith("FZ"))
                {
                    DataRow dr = result.groups_analysis.NewRow();
                    dr["number"] = _groups_data.Columns[i].ColumnName;
                    dr["fullmark"] = group_fullmark(dr["number"].ToString(), ans_count);
                    dr["max"] = _groups_data.Compute("Max([" + _groups_data.Columns[i].ColumnName + "])", "");
                    dr["min"] = _groups_data.Compute("Min([" + _groups_data.Columns[i].ColumnName + "])", "");
                    dr["avg"] = _groups_data.Compute("Avg([" + _groups_data.Columns[i].ColumnName + "])", "");
                    stdev temp = new stdev(result.total_num, (decimal)dr["avg"]);
                    groupStdev.Add(temp);
                    dr["difficulty"] = (decimal)dr["avg"] / (decimal)dr["fullmark"];

                    result.group.Add(new PartitionData.Disc(result.total_num, (decimal)dr["fullmark"]));
                    
                    result.groups_analysis.Rows.Add(dr);
                    ans_count++;
                }
            }
            //修改上限
            row = 1;
            
            foreach (DataRow dr in _groups_data.Rows)
            {
                ans_count = 0;
                for (int i = 3; i < _groups_data.Columns.Count - 2; i++)
                {
                    if (_groups_data.Columns[i].ColumnName.StartsWith("FZ"))
                    {
                        ((stdev)groupStdev[ans_count]).add((decimal)dr[_groups_data.Columns[i]]);
                        if (row <= result.PLN)
                                result.group[ans_count].AddData((decimal)dr[i], true);
                            else if (row >= result.PHN)
                                result.group[ans_count].AddData((decimal)dr[i], false);
                        
                        ans_count++;
                    }

                }
                row++;
            }
            count = 0;
            foreach (DataRow dr in result.groups_analysis.Rows)
            {
                dr["stDev"] = ((stdev)groupStdev[count]).get_value();
                if ((decimal)dr["avg"] == 0)
                    dr["dfactor"] = 0m;
                else
                    dr["dfactor"] = (decimal)dr["stDev"] / (decimal)dr["avg"];

                result.group_discriminant.Add(result.group[count].GetAns());

                count++;
            }


            #endregion
            frequency_table();
            single_groups_analysis();
            if (!isZonghe)
            {
                single_topic_analysis();
            }
            group_mark(_basic_data);
        }
        public void total_tuple_analysis(PartitionData wd, string totalmarkstr)
        {
            wd.Total_tuple_analysis.Columns.Add("name", typeof(string));
            wd.Total_tuple_analysis.Columns.Add("ScoreRange", typeof(string));
            wd.Total_tuple_analysis.Columns.Add("Average", typeof(string));
            wd.Total_tuple_analysis.Columns.Add("difficulty", typeof(string));

            wd.Total_tuple_analysis.PrimaryKey = new DataColumn[] { wd.Total_tuple_analysis.Columns["name"] };

            for (int i = 1; i < _groupnum + 1; i++)
            {
                DataRow dr = wd.Total_tuple_analysis.NewRow();
                dr["name"] = "G" + i;
                dr["ScoreRange"] = "0.0～0.0";
                dr["Average"] = "0.0";
                dr["difficulty"] = "0.0";
                wd.Total_tuple_analysis.Rows.Add(dr);
            }

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
                DataRow dr = wd.Total_tuple_analysis.Rows.Find(tuple.name.Trim());
                dr["ScoreRange"] = tuple.min + "～" + tuple.max;
                dr["Average"] = string.Format("{0:F1}", tuple.avg);
                dr["difficulty"] = string.Format("{0:F2}", tuple.avg / wd.fullmark);
            }

        }
        public void single_topic_analysis()
        {
            int i = 0;
            foreach (DataRow dr in result.total_analysis.Rows)
            {
                PartitionData.single_data temp = new PartitionData.single_data();
                temp.single_detail = new DataTable();
                if (!_standard_ans.Rows.Find(dr["number"].ToString().Substring(1))["da"].ToString().Trim().Equals(""))
                {
                    temp.single_detail.Columns.Add("mark", typeof(string));
                    for (i = 1; i <= _groupnum; i++)
                        temp.single_detail.Columns.Add("G" + i.ToString().Trim(), typeof(decimal));
                    temp.single_detail.Columns.Add("frequency", typeof(int));
                    temp.single_detail.Columns.Add("rate", typeof(decimal));

                    temp.single_detail.Columns.Add("avg", typeof(decimal));

                    temp.single_detail.PrimaryKey = new DataColumn[] { temp.single_detail.Columns["mark"] };

                    var single_avg = from row in _basic_data.AsEnumerable()
                                     group row by row.Field<string>("D" + dr["number"].ToString().Substring(1)) into grp
                                     select new
                                     {
                                         choice = grp.Key,
                                         count = grp.Count(),
                                         avg = grp.Average(row => row.Field<decimal>(cor_col))
                                     };
                    foreach (var item in single_avg)
                    {
                        DataRow single_row = temp.single_detail.NewRow();
                        try
                        {
                            single_row["mark"] = choiceTransfer(item.choice.ToString());
                        }
                        catch (Exception e)
                        {
                            throw new ArgumentException("第" + dr["number"].ToString().Substring(1) + "题存在未知答案" + item.choice.ToString());
                        }
                        single_row["frequency"] = item.count;
                        single_row["rate"] = item.count / Convert.ToDecimal(result.total_num) * 100;
                        single_row["avg"] = item.avg;



                        for (i = 1; i <= _groupnum; i++)
                        {
                            single_row["G" + i.ToString().Trim()] = 0m;
                        }

                        temp.single_detail.Rows.Add(single_row);

                    }



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
                        groups_row[item.groups.ToString().Trim()] = item.count;
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
                    for (i = 1; i <= _groupnum; i++)
                    {
                        single_total_row["G" + i.ToString().Trim()] = 0m;
                        single_avg_row["G" + i.ToString().Trim()] = 0m;
                    }
                    foreach (var item in vertical)
                    {
                        single_total_row[item.groups.ToString().Trim()] = Convert.ToDecimal(item.count);
                        single_avg_row[item.groups.ToString().Trim()] = item.avg / (decimal)dr["fullmark"];
                    }
                    single_total_row["frequency"] = result.total_num;
                    single_total_row["rate"] = 100.0m;
                    single_total_row["avg"] = ZH_avg;

                    single_avg_row["frequency"] = 0;
                    single_avg_row["rate"] = 0m;
                    single_avg_row["avg"] = 0m;

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
                        nochoice_row["avg"] = 0m;
                        foreach (DataRow temp_dr in temp.single_detail.Rows)
                        {
                            nochoice_row["avg"] = (decimal)nochoice_row["avg"] + (decimal)temp_dr["avg"] * (int)temp_dr["frequency"];
                            nochoice_row["frequency"] = (int)nochoice_row["frequency"] + (int)temp_dr["frequency"];
                            for (i = 1; i <= _groupnum; i++)
                                nochoice_row["G" + i.ToString().Trim()] = (decimal)nochoice_row["G" + i.ToString().Trim()] + (decimal)temp_dr["G" + i.ToString().Trim()];

                        }
                        nochoice_row["rate"] = (int)nochoice_row["frequency"] / Convert.ToDecimal(result.total_num) * 100m;
                        if ((int)nochoice_row["frequency"] == 0)
                            nochoice_row["avg"] = 0;
                        else
                            nochoice_row["avg"] = (decimal)nochoice_row["avg"] / (int)nochoice_row["frequency"];


                        _single_detail.Rows.InsertAt(nochoice_row, _single_detail.Rows.Count - 2);

                        
                        temp.single_detail = _single_detail;




                    }
                    else
                    {
                        temp.stype = WordData.single_type.multiple;
                        

                    }
                    for (i = 0; i < temp.single_detail.Rows.Count - 2; i++)
                    {
                        for (int j = 1; j <= _groupnum; j++)
                        {
                            if ((int)temp.single_detail.Rows[i]["frequency"] != 0)
                                temp.single_detail.Rows[i]["G" + j.ToString()] = (decimal)temp.single_detail.Rows[i]["G" + j.ToString()] / (int)temp.single_detail.Rows[i]["frequency"] * 100;
                            else
                                temp.single_detail.Rows[i]["G" + j.ToString()] = 0;
                        }
                    }
                    DataRow ans_row = temp.single_detail.Rows.Find(choiceTransfer(_standard_ans.Rows.Find(dr["number"].ToString().Substring(1))["da"].ToString().Trim()));
                    if(ans_row != null)
                        ans_row["mark"] = "*" + ans_row["mark"];


                }
                else
                {
                    temp.stype = WordData.single_type.sub;
                    temp.single_detail.Columns.Add("mark", typeof(string));
                    for (i = 1; i <= _groupnum; i++)
                        temp.single_detail.Columns.Add("G" + i.ToString().Trim(), typeof(decimal));
                    temp.single_detail.Columns.Add("frequency", typeof(decimal));
                    temp.single_detail.Columns.Add("rate", typeof(decimal));
                    temp.single_detail.Columns.Add("avg", typeof(decimal));

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
                        if (!temp.single_detail.Rows.Contains(string.Format("{0:F1}", item.mark) + "～"))
                        {
                            DataRow temp_dr = temp.single_detail.NewRow();
                            temp_dr["mark"] = string.Format("{0:F1}", item.mark) + "～";
                            temp_dr["frequency"] = item.count;
                            temp_dr["rate"] = 0;
                            temp_dr["avg"] = item.avg * item.count;
                            for (i = 1; i <= _groupnum; i++)
                            {
                                temp_dr["G" + i.ToString().Trim()] = 0m;
                            }
                            temp.single_detail.Rows.Add(temp_dr);
                        }
                        else
                        {
                            DataRow oldrow = temp.single_detail.Rows.Find(string.Format("{0:F1}", item.mark) + "～");
                            oldrow["frequency"] = (decimal)oldrow["frequency"] + item.count;
                            oldrow["avg"] = (decimal)oldrow["avg"] + item.avg * item.count;
                        }
                    }
                    foreach (DataRow row in temp.single_detail.Rows)
                    {
                        row["rate"] = ((decimal)row["frequency"] / result.total_num) * 100;
                        row["avg"] = (decimal)row["avg"] / (decimal)row["frequency"];
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
                        temp_dr[item.groups.ToString().Trim()] = (decimal)temp_dr[item.groups.ToString().Trim()] + item.count;

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
                    total_dr["avg"] = ZH_avg;

                    avg_dr["mark"] = "得分率";
                    avg_dr["frequency"] = 0;
                    avg_dr["rate"] = 0m;
                    avg_dr["avg"] = 0m;

                    for (i = 1; i <= _groupnum; i++)
                    {
                        total_dr["G" + i.ToString().Trim()] = 0m;
                        avg_dr["G" + i.ToString().Trim()] = 0m;
                    }
                    foreach (var item in vertical)
                    {
                        total_dr[item.groups.ToString().Trim()] = Convert.ToDecimal(item.count);
                        avg_dr[item.groups.ToString().Trim()] = item.avg / (decimal)dr["fullmark"];
                    }

                    temp.single_detail.Rows.Add(total_dr);
                    temp.single_detail.Rows.Add(avg_dr);

                    for (i = 0; i < temp.single_detail.Rows.Count - 2; i++)
                    {
                        for (int j = 1; j <= _groupnum; j++)
                        {
                            if ((decimal)temp.single_detail.Rows[i]["frequency"] != 0)
                                temp.single_detail.Rows[i]["G" + j.ToString()] = (decimal)temp.single_detail.Rows[i]["G" + j.ToString()] / (decimal)temp.single_detail.Rows[i]["frequency"] * 100;
                            else
                                temp.single_detail.Rows[i]["G" + j.ToString()] = 0;
                        }
                    }


                }
                result.single_topic_analysis.Add(temp);
            }


        }
        public void xz_single_postprocess(string th)
        {
            #region 选做题部分
            DataTable xz_data = new DataTable();
            List<DataTable> xz_total = new List<DataTable>();
            List<List<PartitionData.single_data>> xz_single = new List<List<PartitionData.single_data>>();
            List<string> xz_name = new List<string>();
            List<List<decimal>> xz_total_disc = new List<List<decimal>>();

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
                DataView dv = xz_data.equalfilter("X" + th, item.name).DefaultView;
                dv.Sort = "totalmark";
                xz_name.Add(item.name);
                xz_group_analysis(dv.ToTable(), item.count, xz_total, xz_single, xz_total_disc);
            }

            for (int i = 0; i < xz_total[0].Rows.Count; i++)
            {
                for (int j = 0; j < xz_total.Count; j++)
                {
                    DataRow dr = xz_total[j].Rows[i];
                    dr["number"] = (string)dr["number"] + "选" + xz_name[j];
                    result.total_analysis.ImportRow(dr);
                    result.single_topic_analysis.Add(xz_single[j][i]);

                    result.total_discriminant.Add(xz_total_disc[j][i]);

                    
                }
            }

            #endregion
        }

        public void xz_postprocess(List<string> xz_th)
        {
            foreach (string th in xz_th)
                xz_single_postprocess(th);
        }
        //public void xz_postprocess(List<string> xz_th)
        //{
        //    DataTable xz_data = new DataTable();
        //    List<DataTable> xz_total = new List<DataTable>();
        //    List<List<PartitionData.single_data>> xz_single = new List<List<PartitionData.single_data>>();
        //    List<string> xz_name = new List<string>();
        //    List<List<decimal>> xz_total_disc = new List<List<decimal>>();
        //    xz_data.Columns.Add("totalmark", typeof(decimal));
        //    xz_data.Columns.Add("XZ", typeof(string));
        //    xz_data.Columns.Add("xz_groups", typeof(string));
        //    foreach (string th in xz_th)
        //    {
        //        xz_data.Columns.Add("T" + th, typeof(decimal));
        //        DataRow dr = _standard_ans.Rows.Find(th);

        //        if (!dr["da"].Equals(""))
        //            xz_data.Columns.Add("D" + th, typeof(string));
        //    }
        //    foreach (DataRow dr in _basic_data.Rows)
        //    {
        //        DataRow newrow = xz_data.NewRow();
        //        foreach (DataColumn dc in xz_data.Columns)
        //            newrow[dc.ColumnName] = dr[dc.ColumnName];
        //        xz_data.Rows.Add(newrow);
        //    }
        //    var xz_tuple = from row in xz_data.AsEnumerable()
        //                   group row by row.Field<string>("XZ") into grp
        //                   orderby grp.Key ascending
        //                   select new
        //                   {
        //                       name = grp.Key,
        //                       count = grp.Count()
        //                   };
        //    foreach (var item in xz_tuple)
        //    {
        //        DataView dv = xz_data.equalfilter("XZ", item.name).DefaultView;
        //        dv.Sort = "totalmark";
        //        xz_name.Add(item.name);
        //        xz_group_analysis(dv.ToTable(), item.count, xz_total, xz_single, xz_total_disc);
        //    }

        //    for (int i = 0; i < xz_total[0].Rows.Count; i++)
        //    {
        //        for (int j = 0; j < xz_total.Count; j++)
        //        {
        //            DataRow dr = xz_total[j].Rows[i];
        //            dr["number"] = (string)dr["number"] + "选" + xz_name[j];
        //            result.total_analysis.ImportRow(dr);
        //            result.single_topic_analysis.Add(xz_single[j][i]);
        //            if (_config.WSLG)
        //            {
        //                ((WSLG_partitiondata)result).total_discriminant.Add(xz_total_disc[j][i]);
        //            }
        //        }
        //    }
        //}
        public void xz_group_analysis(DataTable dt, int xz_count, List<DataTable> xz_total, List<List<PartitionData.single_data>> xz_single, List<List<decimal>> xz_total_disc)
        {
            //dt.SeperateGroups(Utils.group_type, _groupnum);
            DataTable xz_total_analysis = result.total_analysis.Clone();
            List<stdev> stdevlist = new List<stdev>();
            List<WSLG_partitiondata.Disc> xz_disc = new List<WSLG_partitiondata.Disc>();
            List<decimal> total_discriminant = new List<decimal>();
            int PLN = Convert.ToInt32(Math.Ceiling(xz_count * 0.27));
            int PHN = xz_count - PLN + 1;
            Regex number = new Regex("^[Tt]\\d+");
 
            foreach (DataColumn dc in dt.Columns)
            {
                if (number.IsMatch(dc.ColumnName))
                {
                    DataRow dr = xz_total_analysis.NewRow();
                    string topic_num = dc.ColumnName.Substring(1);
                    dr["number"] = dc.ColumnName;
                    dr["total_num"] = xz_count;
                    dr["fullmark"] = Convert.ToDecimal(_standard_ans.Rows.Find(topic_num)["fs"]);
                    dr["max"] = dt.Compute("Max([" + dc.ColumnName + "])", "");
                    dr["min"] = dt.Compute("Min([" + dc.ColumnName + "])", "");
                    dr["avg"] = dt.Compute("Avg([" + dc.ColumnName + "])", "");
                    stdev single_stdev = new stdev(xz_count, (decimal)dr["avg"]);
                    stdevlist.Add(single_stdev);
                    dr["difficulty"] = (decimal)dr["avg"] / (decimal)dr["fullmark"];
                    xz_disc.Add(new PartitionData.Disc(xz_count, (decimal)dr["fullmark"]));
                    
                    xz_total_analysis.Rows.Add(dr);
                }
            }

            int row_num = 0;
            foreach (DataRow dr in dt.Rows)
            {
                int CoCount = 0;
                foreach (DataColumn dc in dt.Columns)
                {
                    if (number.IsMatch(dc.ColumnName))
                    {
                        stdevlist[CoCount].add((decimal)dr[dc]);
                        if (_config.WSLG)
                        {
                            if (row_num < PLN)
                                xz_disc[CoCount].AddData((decimal)dr[dc], true);
                            else if (row_num >= PHN)
                                xz_disc[CoCount].AddData((decimal)dr[dc], false);
                        }
                        CoCount++;
                    }
                }
                row_num++;
            }

            int count = 0;
            foreach (DataRow dr in xz_total_analysis.Rows)
            {
                dr["stDev"] = stdevlist[count].get_value();
                if ((decimal)dr["avg"] == 0)
                    dr["dfactor"] = 0m;
                else
                    dr["dfactor"] = (decimal)dr["stDev"] / (decimal)dr["avg"];
                if (_config.WSLG)
                    total_discriminant.Add(xz_disc[count].GetAns());
                count++;
            }
            xz_total.Add(xz_total_analysis);
            if (_config.WSLG)
                xz_total_disc.Add(total_discriminant);
            List<PartitionData.single_data> xz_single_data = new List<PartitionData.single_data>();
            int i = 0;
            foreach (DataRow dr in xz_total_analysis.Rows)
            {
                PartitionData.single_data temp = new PartitionData.single_data();
                temp.single_detail = new DataTable();
                if (!_standard_ans.Rows.Find(dr["number"].ToString().Substring(1))["da"].ToString().Trim().Equals(""))
                {
                    temp.single_detail.Columns.Add("mark", typeof(string));
                    for (i = 1; i <= _groupnum; i++)
                        temp.single_detail.Columns.Add("G" + i.ToString().Trim(), typeof(decimal));
                    temp.single_detail.Columns.Add("frequency", typeof(int));
                    temp.single_detail.Columns.Add("rate", typeof(decimal));

                    temp.single_detail.Columns.Add("avg", typeof(decimal));

                    temp.single_detail.PrimaryKey = new DataColumn[] { temp.single_detail.Columns["mark"] };

                    var single_avg = from row in dt.AsEnumerable()
                                     group row by row.Field<string>("D" + dr["number"].ToString().Substring(1)) into grp
                                     select new
                                     {
                                         choice = grp.Key,
                                         count = grp.Count(),
                                         avg = grp.Average(row => row.Field<decimal>("totalmark"))
                                     };
                    foreach (var item in single_avg)
                    {
                        DataRow single_row = temp.single_detail.NewRow();
                        single_row["mark"] = choiceTransfer(item.choice.ToString());
                        single_row["frequency"] = item.count;
                        single_row["rate"] = item.count / Convert.ToDecimal(xz_count) * 100;
                        single_row["avg"] = item.avg;



                        for (i = 1; i <= _groupnum; i++)
                        {
                            single_row["G" + i.ToString().Trim()] = 0m;
                        }

                        temp.single_detail.Rows.Add(single_row);

                    }



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
                        groups_row[item.groups.ToString().Trim()] = item.count;
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
                    for (i = 1; i <= _groupnum; i++)
                    {
                        single_total_row["G" + i.ToString().Trim()] = 0m;
                        single_avg_row["G" + i.ToString().Trim()] = 0m;
                    }
                    foreach (var item in vertical)
                    {
                        single_total_row[item.groups.ToString().Trim()] = Convert.ToDecimal(item.count);
                        single_avg_row[item.groups.ToString().Trim()] = item.avg / (decimal)dr["fullmark"];
                    }
                    single_total_row["frequency"] = xz_count;
                    single_total_row["rate"] = 100.0m;
                    single_total_row["avg"] = ZH_avg;

                    single_avg_row["frequency"] = 0;
                    single_avg_row["rate"] = 0m;
                    single_avg_row["avg"] = 0m;

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
                        nochoice_row["avg"] = 0m;
                        foreach (DataRow temp_dr in temp.single_detail.Rows)
                        {
                            nochoice_row["avg"] = (decimal)nochoice_row["avg"] + (decimal)temp_dr["avg"] * (int)temp_dr["frequency"];
                            nochoice_row["frequency"] = (int)nochoice_row["frequency"] + (int)temp_dr["frequency"];
                            for (i = 1; i <= _groupnum; i++)
                                nochoice_row["G" + i.ToString().Trim()] = (decimal)nochoice_row["G" + i.ToString().Trim()] + (decimal)temp_dr["G" + i.ToString().Trim()];

                        }
                        nochoice_row["rate"] = (int)nochoice_row["frequency"] / Convert.ToDecimal(xz_count) * 100m;
                        if ((int)nochoice_row["frequency"] == 0)
                            nochoice_row["avg"] = 0;
                        else
                            nochoice_row["avg"] = (decimal)nochoice_row["avg"] / (int)nochoice_row["frequency"];


                        _single_detail.Rows.InsertAt(nochoice_row, _single_detail.Rows.Count - 2);
                        temp.single_detail = _single_detail;

                        


                    }
                    else
                    {
                        temp.stype = WordData.single_type.multiple;

                        
                    }
                    for (i = 0; i < temp.single_detail.Rows.Count - 2; i++)
                    {
                        for (int j = 1; j <= _groupnum; j++)
                        {
                            if ((int)temp.single_detail.Rows[i]["frequency"] != 0)
                                temp.single_detail.Rows[i]["G" + j.ToString()] = (decimal)temp.single_detail.Rows[i]["G" + j.ToString()] / (int)temp.single_detail.Rows[i]["frequency"] * 100;
                            else
                                temp.single_detail.Rows[i]["G" + j.ToString()] = 0;
                        }
                    }
                    DataRow ans_row = temp.single_detail.Rows.Find(choiceTransfer(_standard_ans.Rows.Find(dr["number"].ToString().Substring(1))["da"].ToString().Trim()));
                    ans_row["mark"] = "*" + ans_row["mark"];


                }
                else
                {
                    temp.stype = WordData.single_type.sub;
                    temp.single_detail.Columns.Add("mark", typeof(string));
                    for (i = 1; i <= _groupnum; i++)
                        temp.single_detail.Columns.Add("G" + i.ToString().Trim(), typeof(decimal));
                    temp.single_detail.Columns.Add("frequency", typeof(decimal));
                    temp.single_detail.Columns.Add("rate", typeof(decimal));
                    temp.single_detail.Columns.Add("avg", typeof(decimal));

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
                        if (!temp.single_detail.Rows.Contains(string.Format("{0:F1}",item.mark) + "～"))
                        {
                            DataRow temp_dr = temp.single_detail.NewRow();
                            temp_dr["mark"] = string.Format("{0:F1}", item.mark) + "～";
                            temp_dr["frequency"] = item.count;
                            temp_dr["rate"] = 0;
                            temp_dr["avg"] = item.avg * item.count;
                            for (i = 1; i <= _groupnum; i++)
                            {
                                temp_dr["G" + i.ToString().Trim()] = 0m;
                            }
                            temp.single_detail.Rows.Add(temp_dr);
                        }
                        else
                        {
                            DataRow oldrow = temp.single_detail.Rows.Find(string.Format("{0:F1}", item.mark) + "～");
                            oldrow["frequency"] = (decimal)oldrow["frequency"] + item.count;
                            oldrow["avg"] = (decimal)oldrow["avg"] + item.avg * item.count;
                        }
                    }
                    foreach (DataRow row in temp.single_detail.Rows)
                    {
                        row["rate"] = ((decimal)row["frequency"] / xz_count) * 100;
                        row["avg"] = (decimal)row["avg"] / (decimal)row["frequency"];
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
                        temp_dr[item.groups.ToString().Trim()] = (decimal)temp_dr[item.groups.ToString().Trim()] + item.count;

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
                    total_dr["frequency"] = xz_count;
                    total_dr["rate"] = 100.0m;
                    total_dr["avg"] = ZH_avg;

                    avg_dr["mark"] = "得分率";
                    avg_dr["frequency"] = 0;
                    avg_dr["rate"] = 0m;
                    avg_dr["avg"] = 0m;

                    for (i = 1; i <= _groupnum; i++)
                    {
                        total_dr["G" + i.ToString().Trim()] = 0m;
                        avg_dr["G" + i.ToString().Trim()] = 0m;
                    }
                    foreach (var item in vertical)
                    {
                        total_dr[item.groups.ToString().Trim()] = Convert.ToDecimal(item.count);
                        avg_dr[item.groups.ToString().Trim()] = item.avg / (decimal)dr["fullmark"];
                    }

                    temp.single_detail.Rows.Add(total_dr);
                    temp.single_detail.Rows.Add(avg_dr);


                    for (i = 0; i < temp.single_detail.Rows.Count - 2; i++)
                    {
                        for (int j = 1; j <= _groupnum; j++)
                        {
                            if ((decimal)temp.single_detail.Rows[i]["frequency"] != 0)
                                temp.single_detail.Rows[i]["G" + j.ToString()] = (decimal)temp.single_detail.Rows[i]["G" + j.ToString()] / (decimal)temp.single_detail.Rows[i]["frequency"] * 100;
                            else
                                temp.single_detail.Rows[i]["G" + j.ToString()] = 0;
                        }
                    }

                }
                xz_single_data.Add(temp);
            }
            xz_single.Add(xz_single_data);
        }
        public void insertRow(DataRow insert_row, DataTable target, int pos)
        {
            DataRow dr = target.NewRow();
            dr.ItemArray = insert_row.ItemArray;
            target.Rows.InsertAt(dr, pos);
        }
        public void single_groups_analysis()
        {
            foreach (DataRow dr in result.groups_analysis.Rows)
            {
                PartitionData.group_data data = new PartitionData.group_data();
                data.group_detail = new DataTable();
                data.group_dist = new DataTable();
                data.group_dist.Columns.Add("mark", typeof(decimal));
                data.group_dist.Columns.Add("rate", typeof(decimal));

                data.group_dist.PrimaryKey = new DataColumn[] { data.group_dist.Columns["mark"]};

                decimal interval = Math.Ceiling(((decimal)dr["fullmark"]) / 20.0m);
                int tuple_num = Convert.ToInt32(Math.Floor(((decimal)dr["fullmark"]) / interval));
                decimal flag = (interval + 1) / 2.0m;
                DataRow start_row = data.group_dist.NewRow();
                start_row["mark"] = 0;
                start_row["rate"] = 0;
                data.group_dist.Rows.Add(start_row);
                int j = 0;
                for (j = 0; j < tuple_num; j++)
                {
                    DataRow inter_row = data.group_dist.NewRow();
                    inter_row["mark"] = flag;
                    inter_row["rate"] = 0;
                    flag += interval;
                    data.group_dist.Rows.Add(inter_row);
                }
                if (((decimal)dr["fullmark"] - tuple_num * interval) != 0)
                {
                    DataRow last_row = data.group_dist.NewRow();
                    last_row["mark"] = tuple_num * interval + 1 + ((decimal)dr["fullmark"] - tuple_num * interval - 1) / 2.0m;
                    last_row["rate"] = 0;
                    data.group_dist.Rows.Add(last_row);
                }


                

                var freq = from row in _groups_data.AsEnumerable()
                           group row by row.Field<decimal>(dr["number"].ToString().Trim()) into grp
                           orderby grp.Key ascending
                           select new
                       {
                           count = grp.Count(),
                           mark = grp.Key,
                           avg = grp.Average(row => row.Field<decimal>(cor_col))
                       };
                data.group_detail.Columns.Add("mark", typeof(string));
                for (int i = 1; i <= _groupnum; i++)
                    data.group_detail.Columns.Add("G" + i.ToString(), typeof(decimal));
                data.group_detail.Columns.Add("frequency", typeof(int));
                data.group_detail.Columns.Add("rate", typeof(decimal));
                data.group_detail.Columns.Add("avg", typeof(decimal));
                data.group_detail.PrimaryKey = new DataColumn[] { data.group_detail.Columns["mark"] };

                int dist_num = 0;
                foreach (var item in freq)
                {
                    if (!data.group_detail.Rows.Contains(Convert.ToInt32(Math.Ceiling(item.mark)).ToString() + "～"))
                    {
                        DataRow newrow = data.group_detail.NewRow();
                        //DataRow dist_row = data.group_dist.NewRow();
                        //dist_row["mark"] = Convert.ToDecimal(Convert.ToInt32(Math.Ceiling(item.mark)));
                        //dist_row["rate"] = Convert.ToDecimal(item.count);
                        newrow["mark"] = Convert.ToInt32(Math.Ceiling(item.mark)).ToString() + "～";
                        newrow["frequency"] = item.count;
                        newrow["rate"] = 0;
                        newrow["avg"] = item.count * item.avg;
                        for (int i = 1; i <= _groupnum; i++)
                            newrow["G" + i.ToString()] = 0.0m;
                        data.group_detail.Rows.Add(newrow);
                        //data.group_dist.Rows.Add(dist_row);
                    }
                    else
                    {
                        DataRow oldrow = data.group_detail.Rows.Find(Convert.ToInt32(Math.Ceiling(item.mark)).ToString() + "～");
                        
                        oldrow["frequency"] = (int)oldrow["frequency"] + item.count;
                        oldrow["avg"] = (decimal)oldrow["avg"] + item.count * item.avg;

                        //DataRow old_distrow = data.group_dist.Rows.Find(Convert.ToDecimal(Convert.ToInt32(Math.Ceiling(item.mark))));
                        //old_distrow["rate"] = (decimal)old_distrow["rate"] + item.count;
                    }
                    dist_num = Convert.ToInt32(Math.Ceiling(item.mark / interval));


                    //if (dist_num == 0)
                    data.group_dist.Rows[dist_num]["rate"] = (decimal)data.group_dist.Rows[dist_num]["rate"] + item.count;
                    //else
                    //    data.group_dist.Rows[dist_num - 1]["rate"] = (decimal)data.group_dist.Rows[dist_num - 1]["rate"] + item.count;
                    
                    
                    
                }
                foreach (DataRow row in data.group_detail.Rows)
                {
                    row["rate"] = ((int)row["frequency"] / Convert.ToDecimal(result.total_num)) * 100;
                    row["avg"] = (decimal)row["avg"] / (int)row["frequency"];
                }
                foreach (DataRow dr2 in data.group_dist.Rows)
                {
                    dr2["rate"] = (decimal)dr2["rate"] / Convert.ToDecimal(result.total_num) * 100;
                }

                var groups = from row in _groups_data.AsEnumerable()
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
                foreach (var item in groups)
                {
                    DataRow target = data.group_detail.Rows.Find(Convert.ToInt32(Math.Ceiling(item.mark)).ToString() + "～");
                    target[item.groups.ToString().Trim()] = (decimal)target[item.groups.ToString().Trim()] + item.count;
                }

                var gdata = from row in _groups_data.AsEnumerable()
                            group row by row.Field<string>("Groups") into grp
                            select new
                            {
                                gtype = grp.Key,
                                count = grp.Count(),
                                avg = grp.Average(row => row.Field<decimal>(dr["number"].ToString()))
                            };
                DataRow total = data.group_detail.NewRow();
                DataRow avg = data.group_detail.NewRow();
                for (int i = 1; i <= _groupnum; i++)
                {
                    total["G" + i.ToString()] = 0.0m;
                    avg["G" + i.ToString()] = 0.0m;
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

                data.group_detail.Rows.Add(total);
                data.group_detail.Rows.Add(avg);

                for (int i = 0; i < data.group_detail.Rows.Count - 2; i++)
                {
                    for (int num = 1; num <= _groupnum; num++)
                    {
                        if ((int)data.group_detail.Rows[i]["frequency"] != 0)
                            data.group_detail.Rows[i]["G" + num.ToString()] = (decimal)data.group_detail.Rows[i]["G" + num.ToString()] / (int)data.group_detail.Rows[i]["frequency"] * 100;
                        else
                            data.group_detail.Rows[i]["G" + num.ToString()] = 0;
                    }
                }
                result.single_group_analysis.Add(data);
            }
        }
        public void frequency_table()
        {
            result.total_dist.Columns.Add("mark", typeof(decimal));
            result.total_dist.Columns.Add("rate", typeof(decimal));
            decimal flag = 0m;
            decimal interval = 1.0m;
            if (result.fullmark > 20.0m)
            {
                interval = Math.Floor(result.fullmark / 20.0m);
                flag = (interval + 1) / 2.0m;

                int j = 0;
                for (j = 0; j < 20; j++)
                {
                    DataRow inter_row = result.total_dist.NewRow();
                    inter_row["mark"] = flag;
                    inter_row["rate"] = 0;
                    flag += interval;
                    result.total_dist.Rows.Add(inter_row);
                }
                if ((result.fullmark - 20.0m * interval) != 0)
                {
                    DataRow last_row = result.total_dist.NewRow();
                    last_row["mark"] = 20.0m * interval + (result.fullmark - 20.0m * interval + 1) / 2.0m;
                    last_row["rate"] = 0;
                    result.total_dist.Rows.Add(last_row);
                }
            }
            else
            {
                int j = 0;
                for (j = 0; j <= result.fullmark; j++)
                {
                    DataRow inter_row = result.total_dist.NewRow();
                    inter_row["mark"] = Convert.ToDecimal(j);
                    inter_row["num"] = 0;
                    result.total_dist.Rows.Add(inter_row);
                }
            }
            var freq = from row in _basic_data.AsEnumerable()
                       group row by row.Field<decimal>(totalmark_str) into grp
                       orderby grp.Key descending
                       select new
                       {
                           count = grp.Count(),
                           totalmark = grp.Key
                       };
            bool first = true;
            int last_freq = 0;
            foreach (var item in freq)
            {
                DataRow dr = result.freq_analysis.NewRow();
                dr["totalmark"] = item.totalmark;
                dr["frequency"] = item.count;
                dr["rate"] = item.count / Convert.ToDecimal(result.total_num) * 100;
                if (first)
                {
                    dr["accumulateFreq"] = dr["frequency"];
                    dr["accumulateRate"] = dr["rate"];
                    last_freq = (int)dr["frequency"];
                    first = false;
                }
                else
                {
                    dr["accumulateFreq"] = last_freq + (int)dr["frequency"];
                    dr["accumulateRate"] = Convert.ToDecimal(dr["accumulateFreq"]) / result.total_num * 100;
                    last_freq = (int)dr["accumulateFreq"];
                }
                result.freq_analysis.Rows.Add(dr);
            }
            DataTable new_freq = result.freq_analysis.Clone();
            new_freq.PrimaryKey = new DataColumn[] { new_freq.Columns["totalmark"] };
            foreach (DataRow dr in result.freq_analysis.Rows)
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
                }

            }
            result.freq_analysis = new_freq;
            int dist_num = 0;
            for (int i = result.freq_analysis.Rows.Count - 1; i >= 0; i--)
            {
                if (interval == 1.0m)
                {
                    dist_num = Convert.ToInt32(Math.Floor((decimal)result.freq_analysis.Rows[i]["totalmark"]));
                    result.total_dist.Rows[dist_num]["rate"] = (decimal)result.total_dist.Rows[dist_num]["rate"] + Convert.ToDecimal(result.freq_analysis.Rows[i]["frequency"]);
                }
                else
                {
                    dist_num = Convert.ToInt32(Math.Ceiling((decimal)result.freq_analysis.Rows[i]["totalmark"] / interval));
                    if (dist_num > 20)
                        result.total_dist.Rows[20]["rate"] = (decimal)result.total_dist.Rows[20]["rate"] + Convert.ToDecimal(result.freq_analysis.Rows[i]["frequency"]);
                    else if (dist_num == 0)
                        result.total_dist.Rows[dist_num]["rate"] = (decimal)result.total_dist.Rows[dist_num]["rate"] + Convert.ToDecimal(result.freq_analysis.Rows[i]["frequency"]);
                    else
                        result.total_dist.Rows[dist_num - 1]["rate"] = (decimal)result.total_dist.Rows[dist_num - 1]["rate"] + Convert.ToDecimal(result.freq_analysis.Rows[i]["frequency"]);
                }
            }
            foreach (DataRow dr in result.total_dist.Rows)
            {
                dr["rate"] = (decimal)dr["rate"] / Convert.ToDecimal(result.total_num) * 100;
            }
        }
        public class stdev
        {
            int _total_num;
            decimal _avg;
            decimal temp;
            public stdev(int total_num, decimal avg)
            {
                _total_num = total_num;
                _avg = avg;
                temp = 0.0m;
            }
            public void add(decimal mark)
            {
                temp += (mark - _avg) * (mark - _avg);
            }

            public decimal get_value()
            {
                return Convert.ToDecimal(Math.Sqrt(Convert.ToDouble(temp / _total_num)));
            }

        }

        public decimal group_fullmark(string name, int row)
        {
            decimal fullmark = 0.0m;
            if (name.Equals("生物") || name.Equals("政治"))
            {
                fullmark = _config.shengwu_zhengzhi;
            }
            else if (name.Equals("物理") || name.Equals("历史"))
            {
                fullmark = _config.wuli_lishi;
            }
            else if (name.Equals("化学") || name.Equals("地理"))
            {
                fullmark = _config.huaxue_dili;
            }
            else
            {
                string spattern = "^\\d+~\\d+$";
                string org = _groups_ans.Rows[row][1].ToString().Trim();
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
                            fullmark += (decimal)dr["fullmark"];
                        }

                    }
                    else
                    {
                        DataRow dr = result.total_analysis.Rows.Find("t" + th.Trim());
                        fullmark += (decimal)dr["fullmark"];
                    }


                }
            }
            return fullmark;
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
        public void group_mark(DataTable dt)
        {
            var mark = from row in dt.AsEnumerable()
                       group row by row.Field<string>("Groups") into grp
                       select new
                       {
                           name = grp.Key,
                           max = grp.Max(row => row.Field<decimal>("totalmark"))
                       };
            foreach (var temp in mark)
            {
                _config.GroupMark.Add(temp.max);
            }
        }
    }
}
