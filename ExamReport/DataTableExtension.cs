using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace ExamReport
{
    public static class DataTableExtension
    {
        public static decimal Skewness(this DataTable dt, string col)
        {
            decimal avg = Avg(dt, col);
            double stdev_pow_3 = Math.Pow(Convert.ToDouble(StDev(dt, col)), 3);
            double e_pow_3 = dt.AsEnumerable().Select(
                c => Math.Pow((Convert.ToDouble(c.Field<decimal>(col)) - Convert.ToDouble(avg)), 3)).Average();
            return Convert.ToDecimal(e_pow_3 / stdev_pow_3);
        }
        public static decimal StDev(this DataTable dt, string col)
        {
            decimal avg = Avg(dt, col);
            double stdev = dt.AsEnumerable().Select(
                c => Math.Pow((Convert.ToDouble(c.Field<decimal>(col)) - Convert.ToDouble(avg)), 2)).Average();
            return Convert.ToDecimal(Math.Sqrt(stdev));
        }
        public static decimal Avg(this DataTable dt, string col)
        {
            return Convert.ToDecimal(dt.Compute("Avg(" + col + ")", ""));
        }
        public static decimal Max(this DataTable dt, string col)
        {
            return Convert.ToDecimal(dt.Compute("Max(" + col + ")", ""));
        }
        public static decimal Min(this DataTable dt, string col)
        {
            return Convert.ToDecimal(dt.Compute("Min(" + col + ")", ""));
        }
        public static DataTable LanguageTrans(this DataTable dt, string exam)
        {
            foreach (DataRow dr in dt.Rows)
            {
                foreach (DataColumn dc in dt.Columns)
                {
                    if (dc.ColumnName.Trim().Equals("sub"))
                    {
                        switch (exam)
                        {
                            case "zk":
                            case "hk":
                            case "ngk":
                                dr[dc] = Utils.hk_lang_trans(dr[dc].ToString());
                                break;
                            
                            default:
                                dr[dc] = Utils.language_trans(dr[dc].ToString());
                                break;
                        }
                        //switch (dr[dc].ToString())
                        //{
                        //    case "sx":
                        //        dr[dc] = "数学";
                        //        break;
                        //    case "yw":
                        //        dr[dc] = "语文";
                        //        break;
                        //    case "yy":
                        //        dr[dc] = "英语";
                        //        break;
                        //    case "wl":
                        //        dr[dc] = "物理";
                        //        break;
                        //    case "hx":
                        //        dr[dc] = "化学";
                        //        break;
                        //    case "sw":
                        //        dr[dc] = "生物";
                        //        break;
                        //    case "zz":
                        //        dr[dc] = "政治";
                        //        break;
                        //    case "ls":
                        //        dr[dc] = "历史";
                        //        break;
                        //    case "dl":
                        //        dr[dc] = "地理";
                        //        break;
                        //    case "lz":
                        //        dr[dc] = "理综";
                        //        break;
                        //    case "wz":
                        //        dr[dc] = "文综";
                        //        break;
                        //    case "sxl":
                        //        dr[dc] = "数学理";
                        //        break;
                        //    case "sxw":
                        //        dr[dc] = "数学文";
                        //        break;
                        //    default:
                        //        break;
                        //}
                    }
                    else if(dc.ColumnName.Trim().Equals("exam"))
                    {
                        switch (dr[dc].ToString())
                        {
                            case "zk":
                                dr[dc] = "中考";
                                break;
                            case "hk":
                                dr[dc] = "会考";
                                break;
                            case "gk":
                            case "gk2020":
                                dr[dc] = "高考";
                                break;
                            default:
                                break;
                        }
                    }
                    else if (dc.ColumnName.Trim().Equals("ans") || dc.ColumnName.Trim().Equals("grp") || dc.ColumnName.Trim().Equals("zh"))
                    {
                        switch (dr[dc].ToString())
                        {
                            case "1":
                                dr[dc] = "已录入";
                                break;
                            case "0":
                                dr[dc] = "未录入";
                                break;
                            default:
                                break;
                        } 
                    }
                    else if (dc.ColumnName.Trim().Equals("gtype"))
                    {
                        switch (dr[dc].ToString())
                        {
                            case "p":
                                dr[dc] = "按人数分";
                                break;
                            case "m":
                                dr[dc] = "按分数分";
                                break;
                            default:
                                break;
                        } 
                    }

                }
            }
            return dt;
        }
        public static void sort(this DataTable dt, string col)
        {
            DataView dv = dt.DefaultView;
            dv.Sort = col;
            dt = dv.ToTable();
        }
        public static int SeperateGroups(this DataTable dt, ZK_database.GroupType gtype, decimal divider, string groupname)
        {
            int _group_num = 0;
            int totalsize = dt.Rows.Count;
            if (gtype.Equals(ZK_database.GroupType.population))
            {
                int remainder = 0;
                int groupnum = Math.DivRem(totalsize, Convert.ToInt32(divider), out remainder);
                _group_num = Convert.ToInt32(divider);
                int remainderCount = 1;
                string groupstring = "";
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    if (i < ((groupnum + 1) * remainder))
                    {
                        if (i % (groupnum + 1) == 0)
                        {
                            groupstring = "G" + remainderCount.ToString();
                            remainderCount++;
                        }

                    }
                    else
                    {
                        if ((i - (groupnum + 1) * remainder) % groupnum == 0)
                        {
                            groupstring = "G" + remainderCount.ToString();
                            remainderCount++;
                        }
                    }
                    dt.Rows[i][groupname] = groupstring;
                }
            }
            else
            {
                decimal baseMark = 0.0m;
                string groupstring = "G1";
                int dividerCount = 1;
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    if ((decimal)dt.Rows[i]["totalmark"] > (baseMark + divider))
                    {
                        dividerCount++;
                        groupstring = "G" + dividerCount.ToString();
                        baseMark = (decimal)dt.Rows[i]["totalmark"];
                    }
                    dt.Rows[i][groupname] = groupstring;
                }
                _group_num = dividerCount;
            }
            return _group_num;
        }
        public static int SeperateGroupsByColumnName(this DataTable dt, ZK_database.GroupType gtype, decimal divider, string ColumnName)
        {
            int _group_num = 0;
            int totalsize = dt.Rows.Count;
            
            
            if (gtype.Equals(ZK_database.GroupType.population))
            {
                int remainder = 0;
                int groupnum = Math.DivRem(totalsize, Convert.ToInt32(divider), out remainder);
                _group_num = Convert.ToInt32(divider);
                int remainderCount = 1;
                string groupstring = "";
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    if (i < ((groupnum + 1) * remainder))
                    {
                        if (i % (groupnum + 1) == 0)
                        {
                            groupstring = "G" + remainderCount.ToString();
                            remainderCount++;
                        }

                    }
                    else
                    {
                        if ((i - (groupnum + 1) * remainder) % groupnum == 0)
                        {
                            groupstring = "G" + remainderCount.ToString();
                            remainderCount++;
                        }
                    }
                    dt.Rows[i]["groups"] = groupstring;
                }
            }
            else
            {
                decimal baseMark = 0.0m;
                string groupstring = "G1";
                int dividerCount = 1;
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    if ((decimal)dt.Rows[i][ColumnName] > (baseMark + divider))
                    {
                        dividerCount++;
                        groupstring = "G" + dividerCount.ToString();
                        baseMark = (decimal)dt.Rows[i]["totalmark"];
                    }
                    dt.Rows[i]["groups"] = groupstring;
                }
                _group_num = dividerCount;
            }
            return _group_num;
        }
        public static DataTable filteredtable(this DataTable dt, string keyword, string[] items)
        {
            DataView dv = dt.DefaultView;
            StringBuilder sb = new StringBuilder();
            sb.Clear();
            sb.Append(keyword + " IN (");
            if (items.Length > 1)
            {
                int i;
                for (i = 0; i < items.Length - 1; i++)
                {
                    sb.Append("'" + items[i] + "',");
                }
                sb.Append("'" + items[i] + "')");
            }
            else
            {
                sb.Append("'" + items[0] + "')");
            }
            dv.RowFilter = sb.ToString();
            if (dt.Columns.Contains("totalmark"))
                dv.Sort = "totalmark";
            else if (dt.Columns.Contains("zf"))
                dv.Sort = "zf";

            if (dv.ToTable().Rows.Count == 0)
            {
                string key;
                switch (keyword)
                {
                    case "QX":
                    case "qxdm":
                        key = "区县代码";
                        break;
                    case "xxdm":
                        key = "学校代码";
                        break;
                    default:
                        key = keyword;
                        break;
                }
                
                throw new ArgumentException("不存在" + key + "为" + string.Join(",", items) + "的数据");
            }
            return dv.ToTable();
        }

        public static DataTable equalfilter(this DataTable dt, string key, string item)
        {
            DataView dv = dt.DefaultView;
            dv.RowFilter = key + " = '" + item + "'";
            if (dt.Columns.Contains("totalmark"))
                dv.Sort = "totalmark";
            else if (dt.Columns.Contains("zf"))
                dv.Sort = "zf";
            return dv.ToTable();
        }

        public static DataTable Likefilter(this DataTable dt, string col, string key)
        {
            DataView dv = dt.DefaultView;
            dv.RowFilter = col + " like " + key;
            if (dt.Columns.Contains("totalmark"))
                dv.Sort = "totalmark";
            else if (dt.Columns.Contains("zf"))
                dv.Sort = "zf";
            return dv.ToTable();
        }

        public static decimal CalCor(this DataTable dt, string Xname, string Yname)
        {
            DataColumn Xdc = dt.Columns[Xname];
            DataColumn Ydc = dt.Columns[Yname];

            decimal Xtotal = 0;
            decimal Ytotal = 0;

            decimal XSquare = 0;
            decimal YSquare = 0;

            decimal XYSqure = 0;
            foreach (DataRow dr in dt.Rows)
            {
                decimal X = (decimal)dr[Xdc];
                decimal Y = (decimal)dr[Ydc];

                Xtotal += X;
                Ytotal += Y;

                XSquare += X * X;
                YSquare += Y * Y;

                XYSqure += X * Y;
            }

            int totalnum = dt.Rows.Count;

            decimal numerator = XYSqure - Xtotal * Ytotal / totalnum;
            decimal Xdenominator = XSquare - Xtotal * Xtotal / totalnum;
            decimal Ydenominator = YSquare - Ytotal * Ytotal / totalnum;
            decimal deno = Convert.ToDecimal(Math.Sqrt(Convert.ToDouble(Xdenominator * Ydenominator)));
            if(deno == 0)
                return 0;
            else
                return numerator / deno;
        }
    }
}
