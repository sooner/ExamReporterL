using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Collections;
using System.Text.RegularExpressions;

namespace ExamReport
{
    public partial class ZK_database
    {
        public enum GroupType { population, totalmark };
        GroupType _gtype;
        decimal _divider;
        public OleDbConnection sqlConnection;
        public DataTable _basic_data;
        public DataTable _standard_ans;
        public DataTable _groups;
        public DataTable _group_data;
        public int _group_num;
        MetaData _mdata;

        List<List<string>> name_list;
        public DataTable newStandard;

        public ZK_database(MetaData mdata, DataTable standard_ans, DataTable groups, GroupType gtype, decimal divider)
        {
            _groups = groups;
            _gtype = gtype;
            _divider = divider;
            _standard_ans = standard_ans;
            _basic_data = new DataTable();
            _group_data = new DataTable();
            _mdata = mdata;

            name_list = new List<List<string>>();
        }
        public string DBF_data_process(string fileadd, MyWizard form)
        {
            Stopwatch st = new Stopwatch();
            st.Start();
            string filePath = @fileadd;
            string file = System.IO.Path.GetFileName(filePath);
            string path = System.IO.Path.GetDirectoryName(filePath);
            string filename = System.IO.Path.GetFileNameWithoutExtension(filePath);
            string filext = System.IO.Path.GetExtension(filePath);

            string conn = @"Provider=vfpoledb;Data Source=" + path + ";Collating Sequence=machine;";
            using (OleDbConnection dbfConnection = new OleDbConnection(conn))
            {
                OleDbDataAdapter adpt = new OleDbDataAdapter("select * from " + file, dbfConnection);
                DataSet mySet = new DataSet();

                adpt.Fill(mySet);
                dbfConnection.Close();
                form.ShowPro(15, 2);
                if (mySet.Tables.Count > 1)
                    return "more than 1 tables";
                DataTable dt = mySet.Tables[0];
                int count = dt.Columns.Count;
                int i;

                newStandard = StandardAnsRecontruction(_standard_ans, name_list);

                DataTable basic_data = new DataTable();
                basic_data.Columns.Add("kh", System.Type.GetType("System.String"));
                basic_data.Columns.Add("xxdm", System.Type.GetType("System.String"));
                basic_data.Columns.Add("totalmark", typeof(decimal));
                for (i = 0; i < newStandard.Rows.Count; i++)
                    basic_data.Columns.Add("T" + ((string)newStandard.Rows[i]["th"]).Trim(), System.Type.GetType("System.Decimal"));
                bool first = true;

                string omrstr = dt.Columns.Contains("Omrstr") ? "Omrstr" : "Info";
                bool has_xz = false;
                if (dt.Columns.Contains("xz"))
                    has_xz = true;
                foreach (DataRow dr in dt.Rows)
                {
                    string an = (string)dr[omrstr];
                    char[] ans = an.Trim().ToCharArray();

                    if (first)
                    {
                        for (i = 0; i < newStandard.Rows.Count; i++)
                        {
                            if(!newStandard.Rows[i]["da"].ToString().Trim().Equals(""))
                                basic_data.Columns.Add("D" + newStandard.Rows[i]["th"].ToString().Trim(), typeof(string));
                        }
                        first = false;
                        basic_data.Columns.Add("Groups",typeof(string));
                        basic_data.Columns.Add("qxdm", typeof(string));
                        if(has_xz)
                            basic_data.Columns.Add("XZ", typeof(string));
                    }

                    DataRow newRow = basic_data.NewRow();
                    newRow["kh"] = ((string)dr["kh"]).Trim();
                    if (dt.Columns.Contains("xxdm"))
                        newRow["xxdm"] = dr["xxdm"].ToString().Trim();
                    else
                        newRow["xxdm"] = "00";
                    newRow["totalmark"] = (decimal)dr["totalmark"];
                    decimal obj_mark = 0;
                    decimal sub_mark = 0;
                    int sub_count = 3, obj_count = 0;
                    for (int n = 0; n < dt.Columns.Count; n++)
                    {
                        if (dt.Columns[n].ColumnName.StartsWith("item"))
                        {
                            sub_count = n;
                            break;
                        }
                    }

                    int total_count = 0;
                    foreach (DataRow ans_dr in newStandard.Rows)
                    {
                        if (ans_dr["da"].ToString().Trim().Equals(""))
                        {
                            if (name_list[total_count] == null)
                            {
                                if ((decimal)dr[sub_count + obj_count] > Convert.ToDecimal(ans_dr["fs"]))
                                    throw new ArgumentException("第" + (string)ans_dr["th"] + "题满分值小于实际分值！");
                                newRow["T" + (string)ans_dr["th"]] = (decimal)dr[sub_count + obj_count];
                                sub_mark += (decimal)dr[sub_count];
                                sub_count++;
                            }
                            else
                            {
                                decimal temp_mark = 0;
                                foreach (string temp_th in name_list[total_count])
                                {
                                    temp_mark += (decimal)newRow["T" + temp_th];
                                }
                                newRow["T" + (string)ans_dr["th"]] = temp_mark;
                            }
                        }
                        else
                        {
                            if (obj_count < ans.Length)
                            {
                                string temp = ((string)ans_dr["da"]).Trim();
                                string th = "T" + ((string)ans_dr["th"]).Trim();
                                if (ans[obj_count].ToString().Equals(temp))
                                {
                                    newRow[th] = Convert.ToDecimal(ans_dr["fs"]);
                                    obj_mark += Convert.ToDecimal(ans_dr["fs"]);
                                }
                                else if (_mdata.PartialRight != 0 && Utils.isContain(temp, ans[obj_count].ToString()))
                                {
                                    if (_mdata.PartialRight > Convert.ToDecimal(ans_dr["fs"]))
                                        throw new ArgumentException("选择题半分分数大于满分分数！");

                                    newRow[th] = _mdata.PartialRight;
                                    obj_mark += _mdata.PartialRight;

                                }
                                else
                                {
                                    newRow[th] = 0.0;
                                }
                                newRow["D" + ((string)ans_dr["th"]).Trim()] = ans[obj_count].ToString();

                                obj_count++;
                            }
                            else
                                throw new ArgumentException("标准答案选择题数量大于数据库中选择题数量");
                        }
                        total_count++;
                    }
                    
                    if (_mdata.sub_iszero && sub_mark == 0)
                        continue;
                    if (_mdata.fullmark_iszero && (decimal)newRow["totalmark"] == 0)
                        continue;
                    newRow["Groups"] = "";
                    if (dt.Columns.Contains("qxdm"))
                        newRow["qxdm"] = dr["qxdm"].ToString().Trim();
                    else
                        newRow["qxdm"] = "00";
                    if (has_xz)
                        newRow["XZ"] = dr["xz"].ToString().Trim();
                    basic_data.Rows.Add(newRow);
                }
                _basic_data = basic_data.Copy();
                DataView dv = basic_data.DefaultView;
                dv.Sort = "totalmark";
                _basic_data = dv.ToTable();
                form.ShowPro(30, 2);
                int totalsize = _basic_data.Rows.Count;
                if (_gtype.Equals(GroupType.population))
                {
                    int remainder = 0;
                    int groupnum = Math.DivRem(totalsize, Convert.ToInt32(_divider), out remainder);
                    _group_num = Convert.ToInt32(_divider);
                    int remainderCount = 1;
                    string groupstring = "";
                    for (i = 0; i < _basic_data.Rows.Count; i++)
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
                        _basic_data.Rows[i]["Groups"] = groupstring;
                    }
                }
                else
                {
                    decimal baseMark = 0.0m;
                    string groupstring = "G1";
                    int dividerCount = 1;
                    for (i = 0; i < _basic_data.Rows.Count; i++)
                    {
                        if ((decimal)_basic_data.Rows[i]["totalmark"] > (baseMark + _divider))
                        {
                            dividerCount++;
                            groupstring = "G" + dividerCount.ToString();
                            baseMark = (decimal)_basic_data.Rows[i]["totalmark"];
                        }
                        _basic_data.Rows[i]["Groups"] = groupstring;
                    }
                    _group_num = dividerCount;
                }
                
                #region divide the table into groups
                //StringBuilder objectdata = new StringBuilder();
                _group_data.Columns.Add("kh", System.Type.GetType("System.String"));
                _group_data.Columns.Add("xxdm", System.Type.GetType("System.String"));
                _group_data.Columns.Add("totalmark", System.Type.GetType("System.Decimal"));
                ArrayList tm = new ArrayList();
                string spattern = "^\\d+~\\d+$";
                for(i=0; i<_groups.Rows.Count; i++)
                {
                    ArrayList tz = new ArrayList();
                    string row_name = _groups.Rows[i][0].ToString().Trim();
                    //_group_data.Columns.Add(row_name, System.Type.GetType("System.Decimal"));
                    _group_data.Columns.Add("FZ" + (i + 1).ToString(), System.Type.GetType("System.Decimal"));
                    string org = _groups.Rows[i][1].ToString().Trim();
                    string[] org_char = org.Split(new char[2]{',','，'});
                    foreach (string th in org_char)
                    {

                        if (System.Text.RegularExpressions.Regex.IsMatch(th, spattern))
                        //if(th.Contains('~'))
                        {
                            string[] num = th.Split('~');
                            int j;
                            int size = Convert.ToInt32(num[0]) < Convert.ToInt32(num[1])? Convert.ToInt32(num[1]): Convert.ToInt32(num[0]);
                            int start = Convert.ToInt32(num[0]) > Convert.ToInt32(num[1]) ? Convert.ToInt32(num[1]) : Convert.ToInt32(num[0]);
                            //此处需判断size和start的边界问题
                            for (j = start; j < size + 1; j++)
                            {
                                tz.Add(j);
                            }

                        }
                        else
                            tz.Add(th);
                    }
                    tm.Add(tz);
                }
                _group_data.Columns.Add("Groups", typeof(string));
                _group_data.Columns.Add("qxdm", typeof(string));
                foreach (DataRow dr in _basic_data.Rows)
                {
                    DataRow newRow = _group_data.NewRow();
                    newRow["kh"] = ((string)dr[0]).Trim();
                    newRow["xxdm"] = ((string)dr[1]).Trim();
                    newRow["Groups"] = ((string)dr["Groups"]).Trim();
                    newRow["qxdm"] = dr["qxdm"].ToString().Trim();
                    newRow["totalmark"] = dr[2];
                    int j;
                    for (j = 0; j < _groups.Rows.Count; j++)
                    {
                        decimal count_ = 0;
                        foreach (object s in (ArrayList)tm[j])
                        {
                            count_ += (decimal)dr["T" + s.ToString()];
                        }
                        newRow[j+3] = count_;
                    }
                    _group_data.Rows.Add(newRow);
                }
                #endregion


                st.Stop();
                
                return st.ElapsedMilliseconds.ToString();
            } 
            
        }
        public DataTable StandardAnsRecontruction(DataTable dt, List<List<string>> name)
        {
            DataTable newtable = dt.Clone();
            Stack<string> sk = new Stack<string>();

            newtable.PrimaryKey = new DataColumn[] { newtable.Columns["th"] };
            foreach (DataRow dr in dt.Rows)
            {
                DataRow newrow = newtable.NewRow();
                string th = dr["th"].ToString().Trim();
                //if (!th.Contains("_"))
                //{
                //    newrow.ItemArray = dr.ItemArray;
                //    newtable.Rows.Add(newrow);
                //    name.Add(null);
                //    continue;
                //}
                if (sk.Count == 0)
                {
                    if (th.Contains("_"))
                        sk.Push(th);
                    newrow.ItemArray = dr.ItemArray;
                    newtable.Rows.Add(newrow);
                    name.Add(null);
                }
                else
                {
                    string prefix = omit_tail(sk.Peek());
                    if (th.StartsWith(prefix))
                    {
                        if (th.Contains("_"))
                            sk.Push(th);
                        newrow.ItemArray = dr.ItemArray;
                        newtable.Rows.Add(newrow);
                        name.Add(null);
                    }
                    else
                    {
                        while (true)
                        {

                            popstack(newtable, sk, name);
                            if (!sk.Peek().Contains("_"))
                            {
                                sk.Pop();
                                if (th.Contains("_"))
                                    sk.Push(th);
                                newrow.ItemArray = dr.ItemArray;
                                newtable.Rows.Add(newrow);
                                name.Add(null);
                                break;
                            }
                            else if (th.StartsWith(omit_tail(sk.Peek())))
                            {
                                if (th.Contains("_"))
                                    sk.Push(th);
                                newrow.ItemArray = dr.ItemArray;
                                newtable.Rows.Add(newrow);
                                name.Add(null);
                                break;
                            }
                            else
                                continue;

                        }
                    }
                }

            }
            while (sk.Count > 0)
            {

                popstack(newtable, sk, name);
                if (!sk.Peek().Contains("_"))
                    sk.Pop();

            }
            return newtable;
        }
        public void popstack(DataTable dt, Stack<string> sk, List<List<string>> name)
        {
            List<string> record = new List<string>();
            string temp_th;
            DataRow dr = dt.NewRow();
            double mark = 0;
            while (true)
            {
                temp_th = sk.Pop();
                record.Add(temp_th);
                mark += Convert.ToDouble(dt.Rows.Find(temp_th)["fs"]);
                if (sk.Count != 0 && sk.Peek().StartsWith(omit_tail(temp_th)))
                    continue;
                else
                    break;
            }
            sk.Push(omit_tail(temp_th));
            if (record.Count > 1)
            {
                dr["th"] = omit_tail(temp_th);
                dr["fs"] = Convert.ToInt32(mark).ToString();
                dr["da"] = "";
                dt.Rows.Add(dr);
                name.Add(record);
            }
        }
        public string omit_tail(string serial)
        {
            Regex num_regex = new Regex(@"(\d+_)+\d+$");
            if (!num_regex.IsMatch(serial))
                throw new ArgumentException("标准答案 " + serial + " 题号格式不正确！");
            MatchCollection match = Regex.Matches(serial, @"\w+(?=_\d+$)");
            if (match.Count > 1)
                throw new ArgumentException("标准答案 " + serial + " 题号格式不正确！");
            return match[0].ToString();
        }
    }

}
