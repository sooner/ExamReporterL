﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace ExamReport
{
    class HKScriptCalculate
    {
        public DataTable total;
        public DataTable _data;
        public DataTable group_order;
        public DataTable execute_batch_zf(DataTable zf, string key, string field)
        {
            DataTable ret = zf.Clone();
            List<string> stu_tuples = zf.AsEnumerable().Where(c => c.Field<string>(field).Equals(key)).Select(c => c.Field<string>("zkzh")).ToList();
            foreach (string stu_id in stu_tuples)
            {
                DataRow single = ret.NewRow();
                single.ItemArray = execute_zf(zf, stu_id).ItemArray;
                ret.Rows.Add(single);
            }
            return ret;
        }
        public DataTable execute_batch(DataTable group, DataTable data, string key, string field)
        {
            DataTable ret = data.Clone();
            List<string> stu_tuples = data.AsEnumerable().Where(c => c.Field<string>(field).Equals(key)).Select(c => c.Field<string>("zkzh")).ToList();
            foreach(string stu_id in stu_tuples)
            {
                DataRow single = ret.NewRow();
                single.ItemArray = execute2(group, data, stu_id).ItemArray;
                ret.Rows.Add(single);
            }
            return ret;
        }
        public DataRow execute_zf(DataTable zf, string stu_id)
        {
            zf.PrimaryKey = new DataColumn[] { zf.Columns["zkzh"] };
            DataRow dr = zf.Rows.Find(stu_id);
            DataRow ret = zf.NewRow();
            ret.ItemArray = dr.ItemArray;
            for (int i = 0; i < Utils.hk_subject.Length; i++)
            {
                string col = Utils.hk_subject[i];
                decimal mark = 0;
                if (dr[col] != DBNull.Value)
                {
                    mark = (decimal)dr[col];
                    int count = zf.AsEnumerable().Count(c => c[col] != DBNull.Value && c.Field<decimal>(col) < mark);
                    int totalcount = zf.AsEnumerable().Count(c => c[col] != DBNull.Value);
                    ret[col] = Convert.ToDecimal(count) / Convert.ToDecimal(totalcount) * 100;
                }
                else
                    ret[col] = 0;

                
            }
            return ret;
        }
        public DataRow execute2(DataTable group, DataTable data, string stu_id)
        {
            DataRow dr = data.Rows.Find(stu_id);
            DataRow ret = data.NewRow();
            ret.ItemArray = dr.ItemArray;
            int totalcount = data.Rows.Count;
            for (int i = 0; i < group.Rows.Count; i++)
            {
                string col = "FZ" + (i+1).ToString();
                decimal mark = (decimal)dr[col];
                int count = data.AsEnumerable().Count(c => c.Field<decimal>(col) < mark);
                ret[col] = Convert.ToDecimal(count) / Convert.ToDecimal(totalcount) * 100;
            }

            return ret;

        }
        public void preprocess(DataTable data, Analysis.HK_hierarchy hk_hierarchy)
        {
            data.Columns.Add("PR_total", System.Type.GetType("System.Decimal"));
            data.Columns.Add("rank", typeof(string));
            for (int i = 0; i < data.Columns.Count; i++)
            {
                if (data.Columns[i].ColumnName.StartsWith("FZ"))
                {
                    data.Columns.Add("PR" + data.Columns[i].ColumnName.Substring(2).Trim(), System.Type.GetType("System.Decimal"));
                }
            }

            foreach (DataRow dr in data.Rows)
            {
                foreach (DataColumn dc in data.Columns)
                {
                    if (dc.ColumnName.StartsWith("PR"))
                        dr[dc] = 0;
                }
                decimal totalmark = (decimal)dr["totalmark"];
                if ( totalmark >= hk_hierarchy.A_low)
                {
                    dr["rank"] = "A";
                }
                else if (totalmark >= hk_hierarchy.B_low)
                {
                    dr["rank"] = "B";
                }
                else if (totalmark >= hk_hierarchy.C_low)
                {
                 
                    dr["rank"] = "C";
                }
                else if (totalmark >= hk_hierarchy.D_low)
                {
                    dr["rank"] = "D";
                }
                else
                {
                    dr["rank"] = "E";
                }
            }

            
            
        }
        public void execute(DataTable ans, DataTable group, DataTable data)
        {

            int totalnum = data.Rows.Count;
            decimal mark = (decimal)data.Rows[0]["totalmark"];
            int count = 1;
            data.Rows[0]["PR_total"] = PercentRank(1, totalnum);
            for (int i = 1; i < data.Rows.Count; i++)
            {
                DataRow dr = data.Rows[i];
                if ((decimal)dr["totalmark"] == mark)
                {
                    dr["PR_total"] = PercentRank(count, totalnum);

                }
                else
                {
                    dr["PR_total"] = PercentRank(i + 1, totalnum);
                    count = i + 1;
                    mark = (decimal)dr["totalmark"];
                }

            }
            DataView dv = null;
            for (int i = 0; i < group.Rows.Count; i++)
            {
                dv = data.DefaultView;
                dv.Sort = "FZ" + (i + 1).ToString() + " desc";
                data = dv.ToTable();
                mark = (decimal)data.Rows[0]["FZ" + (i + 1).ToString()];
                count = 1;
                data.Rows[0]["PR" + (i + 1).ToString()] = PercentRank(1, totalnum);
                for (int j = 1; j < data.Rows.Count; j++)
                {
                    DataRow dr = data.Rows[j];
                    if ((decimal)dr["FZ" + (i + 1).ToString()] == mark)
                        dr["PR" + (i + 1).ToString()] = PercentRank(count, totalnum);
                    else
                    {
                        dr["PR" + (i + 1).ToString()] = PercentRank(j + 1, totalnum);
                        count = j + 1;
                        mark = (decimal)dr["FZ" + (i + 1).ToString()];
                    }
                }
            }

            dv = data.DefaultView;
            dv.Sort = "totalmark desc";

            _data = dv.ToTable();
            _data.PrimaryKey = new DataColumn[] { _data.Columns["studentid"] };
            total = new DataTable();
            total.Columns.Add("type", System.Type.GetType("System.String"));
            total.Columns.Add("num", typeof(int));

            for (int i = 0; i < group.Rows.Count; i++)
                total.Columns.Add("FZ" + (i + 1).ToString(), typeof(decimal));
            total.PrimaryKey = new DataColumn[] { total.Columns["type"] };
            total.Rows.Add(getNewRow(total, "total", group.Rows.Count));
            total.Rows.Add(getNewRow(total, "excellent", group.Rows.Count));
            total.Rows.Add(getNewRow(total, "well", group.Rows.Count));
            total.Rows.Add(getNewRow(total, "pass", group.Rows.Count));
            total.Rows.Add(getNewRow(total, "fail", group.Rows.Count));

            for (int i = 0; i < group.Rows.Count; i++)
                total.Rows[0]["FZ" + (i + 1).ToString()] = (decimal)data.Compute("Avg(FZ" + (i + 1).ToString() + ")", "");
            total.Rows[0]["num"] = data.Rows.Count;
            var num_count = from row in data.AsEnumerable()
                            group row by row.Field<string>("rank") into grp
                            select new
                            {
                                rank = grp.Key,
                                count = grp.Count()
                            };
            foreach (var item in num_count)
                total.Rows.Find(item.rank.ToString().Trim())["num"] = item.count;
            for (int i = 0; i < group.Rows.Count; i++)
            {
                var rank = from row in data.AsEnumerable()
                           group row by row.Field<string>("rank") into grp
                           select new
                           {
                               rank = grp.Key,
                               mark = grp.Average(row => row.Field<decimal>("FZ" + (i + 1).ToString())),
                               count = grp.Count()
                           };
                foreach (var item in rank)
                {
                    if (item.rank.Equals("excellent"))
                        total.Rows.Find("excellent")["FZ" + (i + 1).ToString()] = item.mark;
                    else if (item.rank.Equals("well"))
                        total.Rows.Find("well")["FZ" + (i + 1).ToString()] = item.mark;
                    else if (item.rank.Equals("pass"))
                        total.Rows.Find("pass")["FZ" + (i + 1).ToString()] = item.mark;
                    else if (item.rank.Equals("fail"))
                        total.Rows.Find("fail")["FZ" + (i + 1).ToString()] = item.mark;
                }
            }

        }

        decimal PercentRank(int rank, int count)
        {
            return 100 - (100.0m * rank - 50.0m) / count;
        }

        DataRow getNewRow(DataTable dt, string type, int count)
        {
            DataRow nr = dt.NewRow();
            nr["type"] = type;
            for (int i = 0; i < count; i++)
                nr["FZ" + (i + 1).ToString()] = 0m;
            return nr;
        }


    }
}
