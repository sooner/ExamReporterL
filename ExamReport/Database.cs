using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.OleDb;
using System.Text.RegularExpressions;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace ExamReport
{
    class Database
    {
        public DataTable _groups;
        public int _group_num;
        public ZK_database.GroupType _gtype;
        public decimal _divider;
        public DataTable _standard_ans;
        public DataTable _basic_data;
        public DataTable _group_data;

        public DataTable newStandard;

        public DataTable zh_single_data;
        public DataTable zh_group_data;
        public DataTable ZH_standard_ans;
        string filePath;
        string file;
        string path;
        string filename;
        string filext;
        List<List<string>> name_list;
        MetaData _mdata;
        public DataTable zh_groups;
        public string _sub_name = null;

        OleDbConnection dbfConnection;

        public Database(MetaData mdata, DataTable standard_ans, DataTable groups, ZK_database.GroupType gtype, decimal divider)
        {
            _groups = groups;
            _gtype = gtype;
            _divider = divider;
            _standard_ans = standard_ans;
            _standard_ans.PrimaryKey = new DataColumn[] { _standard_ans.Columns[0] };
            _basic_data = new DataTable();
            _group_data = new DataTable();
            name_list = new List<List<string>>();
            _mdata = mdata;
        }
        public Database(MetaData mdata, DataTable standard_ans, DataTable groups, ZK_database.GroupType gtype, decimal divider, DataTable wenli, string sub_name)
        {
            zh_groups = wenli;
            _sub_name = sub_name;
            _groups = groups;
            _gtype = gtype;
            _divider = divider;
            _standard_ans = standard_ans;
            _standard_ans.PrimaryKey = new DataColumn[] { _standard_ans.Columns[0] };
            _basic_data = new DataTable();
            _group_data = new DataTable();
            name_list = new List<List<string>>();
            _mdata = mdata;

        }
        public Database()
        {
        }

        public int ZH_postprocess()
        {
            string name = _sub_name.Substring(3);
            Regex number = new Regex("^[Tt]\\d+$");
            //zh_groups.PrimaryKey = new DataColumn[] { zh_groups.Columns[0] };
            //DataRow target = zh_groups.Rows.Find(name);
            DataRow target = null;
            foreach (DataRow dr in zh_groups.Rows)
            {
                if (dr[0].ToString().Trim().Equals(name))
                {
                    target = dr;
                    break;
                }
            }
            if (target == null)
                throw new ArgumentException("文理综分类中不存在"+ name +"科分组题目信息");
            string[] tz = target[1].ToString().Trim().Split(new char[2] { ',', '，' });
            List<string> tzs = new List<string>();
            group_process(tz, tzs);
            var standard_th = newStandard.AsEnumerable().Select(c => c.Field<string>("th").Trim());
            List<string> whole_th = new List<string>();
            foreach (string th in tzs)
            {
                var similar = standard_th.Where(c => c.StartsWith(th + "_")).Select(c => c.ToString().Trim());
                if (similar.Count() != 0)
                {
                    foreach (var temp_th in similar)
                        whole_th.Add(temp_th);
                }
            }
            whole_th.AddRange(tzs);
            //List<List<string>> ZH_name_list = new List<List<string>>();
            ZH_standard_ans = newStandard.Clone();
            //ZH_standard_ans = newStandard.Clone();
            int name_list_count = 0;
            foreach (DataRow dr in newStandard.Rows)
            {
                if (whole_th.Contains((string)dr["th"]))
                {
                    ZH_standard_ans.ImportRow(dr);
                    //ZH_name_list.Add(name_list[name_list_count]);
                }
                name_list_count++;
            }
            if (ZH_standard_ans.Rows.Count != whole_th.Count)
                throw new ArgumentException(name + " 题组中存在未知题号！");

            //ZH_standard_ans = StandardAnsRecontruction(temp_ZH_standard_ans, ZH_name_list);
            zh_single_data = new DataTable();
            zh_single_data.Columns.Add("kh", System.Type.GetType("System.String"));
            zh_single_data.Columns.Add("zkzh", System.Type.GetType("System.String"));
            zh_single_data.Columns.Add("xxdm", System.Type.GetType("System.String"));
            zh_single_data.Columns.Add("xb", System.Type.GetType("System.String"));
            zh_single_data.Columns.Add("totalmark", typeof(decimal));
            foreach (DataRow dr in ZH_standard_ans.Rows)
            {
                zh_single_data.Columns.Add("T" + ((string)dr["th"]).Trim(), typeof(decimal));
            }

            int multiple_choice_num = 0;
            foreach (string temp in whole_th)
            {
                if (_basic_data.Columns.Contains("D" + temp))
                {
                    zh_single_data.Columns.Add("D" + temp, typeof(string));
                    multiple_choice_num++;
                }
            }
            zh_single_data.Columns.Add("Groups", typeof(string));
            zh_single_data.Columns.Add("qxdm", typeof(string));
            if (_basic_data.Columns.Contains("X"))
                zh_single_data.Columns.Add("X", typeof(string));
            zh_single_data.Columns.Add("ZH_totalmark", typeof(decimal));
            _group_data.Columns.Add("ZH_totalmark", typeof(decimal));
            int row = 0;
            foreach (DataRow dr in _basic_data.Rows)
            {
                DataRow newrow = zh_single_data.NewRow();
                decimal totalmark = 0;
                foreach (DataColumn dc in zh_single_data.Columns)
                {
                    if (!dc.ColumnName.Equals("ZH_totalmark"))
                    {
                        newrow[dc] = dr[dc.ColumnName];
                        
                    }
                }
                foreach (string th in tzs)
                {
                    totalmark += (decimal)newrow["T" + th];
                }
                
                //for (int i = 0; i < ZH_standard_ans.Rows.Count; i++)
                //{
                //    if (ZH_name_list[i] == null)
                //    {
                //        newrow[i + 4] = dr[zh_single_data.Columns[i + 4].ColumnName];
                //        totalmark += (decimal)newrow[i + 4];
                //    }

                //    else
                //    {
                //        decimal temp_mark = 0;
                //        foreach (string temp_th in ZH_name_list[i])
                //        {
                //            temp_mark += (decimal)dr["T" + temp_th];
                //        }
                //        newrow[i + 4] = temp_mark;
                //    }
                //}
                //for (int i = ZH_standard_ans.Rows.Count + 4; i < zh_single_data.Columns.Count - 1; i++)
                //    newrow[i] = dr[zh_single_data.Columns[i].ColumnName];

                newrow["ZH_totalmark"] = totalmark;
                //if (!_group_data.Rows[row]["kh"].ToString().Trim().Equals(newrow["kh"].ToString().Trim()))
                //    throw new Exception();
                _group_data.Rows[row]["ZH_totalmark"] = totalmark;
                zh_single_data.Rows.Add(newrow);
                row++;
            }
            //var zh_result = _group_data.AsEnumerable().Join(zh_single_data.AsEnumerable().Select(c => new
            //{
            //    kh = c.Field<string>("kh"),
            //    ZH_totalmark = c.Field<decimal>("ZH_totalmark")
            //}), c => c.Field<string>("kh"), p => p.kh, (c, p) => new
            //{
            //    c = c,
            //    p = p
            //});
            //DataTable zh_temp = _group_data.Clone();

            //foreach (var item in zh_result)
            //{

            //}
            List<List<string>> group_th = new List<List<string>>();
            zh_group_data = new DataTable();
            zh_group_data.Columns.Add("kh", typeof(string));
            zh_group_data.Columns.Add("zkzh", System.Type.GetType("System.String"));
            zh_group_data.Columns.Add("xxdm", typeof(string));
            zh_group_data.Columns.Add("xb", typeof(string));
            zh_group_data.Columns.Add("totalmark", typeof(decimal));
            int cor_count = 1;
            foreach (DataRow dr in zh_groups.Rows)
            {

                string group_name = dr[0].ToString().Trim();
                zh_group_data.Columns.Add("FZ" + cor_count.ToString(), typeof(decimal));
                string[] th_string = dr[1].ToString().Trim().Split(new char[2] { ',', '，' });

                List<string> th = new List<string>();
                group_process(th_string, th);
                group_th.Add(th);
                cor_count++;
            }
            zh_group_data.Columns.Add("Groups", typeof(string));
            zh_group_data.Columns.Add("qxdm", typeof(string));

            foreach (DataRow dr in _basic_data.Rows)
            {
                DataRow newrow = zh_group_data.NewRow();
                newrow["kh"] = dr["kh"].ToString();
                newrow["zkzh"] = dr["zkzh"].ToString();
                newrow["xxdm"] = dr["xxdm"].ToString();
                newrow["Groups"] = ((string)dr["Groups"]).Trim();
                newrow["qxdm"] = dr["qxdm"].ToString().Trim();
                newrow["totalmark"] = dr["totalmark"];
                newrow["xb"] = dr["xb"].ToString().Trim();

                for (int i = 0; i < zh_groups.Rows.Count; i++)
                {

                    decimal mark = 0;
                    foreach (string temp in group_th[i])
                    {
                        mark += (decimal)dr["T" + temp];
                    }
                    newrow[i + 5] = mark;

                }
                zh_group_data.Rows.Add(newrow);
            }
            
            if (Utils.saveMidData)
            {
                //create_groups_table(zh_group_data, true);
            }
            //update_standard_ans();
            return multiple_choice_num;
        }
        public void group_process(string[] tz, List<string> tzs)
        {
            string spattern = "^\\d+~\\d+$";
            foreach (string temp in tz)
            {
                if (System.Text.RegularExpressions.Regex.IsMatch(temp, spattern))
                //if(th.Contains('~'))
                {
                    string[] num = temp.Split('~');
                    int j;
                    int size = Convert.ToInt32(num[0]) < Convert.ToInt32(num[1]) ? Convert.ToInt32(num[1]) : Convert.ToInt32(num[0]);
                    int start = Convert.ToInt32(num[0]) > Convert.ToInt32(num[1]) ? Convert.ToInt32(num[1]) : Convert.ToInt32(num[0]);
                    //此处需判断size和start的边界问题
                    for (j = start; j < size + 1; j++)
                    {
                        tzs.Add(j.ToString());
                    }

                }
                else
                    tzs.Add(temp);
            }
        }
        public void zk_zf_process(string fileadd)
        {
            filePath = @fileadd;
            file = System.IO.Path.GetFileName(filePath);
            path = System.IO.Path.GetDirectoryName(filePath);
            filename = System.IO.Path.GetFileNameWithoutExtension(filePath);
            filext = System.IO.Path.GetExtension(filePath);

            string conn = @"Provider=vfpoledb;Data Source=" + path + ";Collating Sequence=machine;";

            dbfConnection = new OleDbConnection(conn);

            OleDbDataAdapter adpt = new OleDbDataAdapter("select * from " + file + " where zf<>0", dbfConnection);
            DataSet mySet = new DataSet();
            adpt.Fill(mySet);
            dbfConnection.Close();
            _basic_data = mySet.Tables[0];

        }
        public void ZF_data_process(string fileadd)
        {
            filePath = @fileadd;
            file = System.IO.Path.GetFileName(filePath);
            path = System.IO.Path.GetDirectoryName(filePath);
            filename = System.IO.Path.GetFileNameWithoutExtension(filePath);
            filext = System.IO.Path.GetExtension(filePath);

            string conn = @"Provider=vfpoledb;Data Source=" + path + ";Collating Sequence=machine;";

            dbfConnection = new OleDbConnection(conn);

            OleDbDataAdapter adpt = new OleDbDataAdapter("select * from " + file + " where zf<>0", dbfConnection);
            DataSet mySet = new DataSet();
            adpt.Fill(mySet);
            dbfConnection.Close();
            _basic_data = mySet.Tables[0];

            _basic_data.Columns.Add("type", typeof(string));
            Regex w_mh = new Regex(@"^1\d+");
            Regex l_mh = new Regex(@"^5\d+");
            foreach (DataRow dr in _basic_data.Rows)
            {
                if (w_mh.IsMatch((string)dr["mh"]))
                    dr["type"] = "w";
                else if (l_mh.IsMatch((string)dr["mh"]))
                    dr["type"] = "l";
                else
                    dr["type"] = "n";
            }

        }

        public string DBF_data_process(string fileadd)
        {
            List<string> sub_tz = new List<string>();
            filePath = @fileadd;
            file = System.IO.Path.GetFileName(filePath);
            path = System.IO.Path.GetDirectoryName(filePath);
            filename = System.IO.Path.GetFileNameWithoutExtension(filePath);
            filext = System.IO.Path.GetExtension(filePath);

            string conn = @"Provider=vfpoledb;Data Source=" + path + ";Collating Sequence=machine;";
            Regex choice = new Regex("^[A-Za-z0@]+$");
            dbfConnection = new OleDbConnection(conn);

            OleDbDataAdapter adpt = new OleDbDataAdapter("select * from " + file, dbfConnection);
            //OleDbDataAdapter adpt = new OleDbDataAdapter("select * from " + file + " where Qk<>1", dbfConnection);
            DataSet mySet = new DataSet();

            try
            {
                adpt.Fill(mySet);
            }
            catch (OleDbException e)
            {
                throw e;
            }
            dbfConnection.Close();
            if (mySet.Tables.Count > 1)
                return "more than 1 tables";
            DataTable dt = mySet.Tables[0];
            int count = dt.Columns.Count;
            int i;
            xz_check(dt, _mdata.xz);
            if (_mdata.sub_iszero)
            {
                DataTable group_t;
                if (_sub_name == null)
                    group_t = _groups;
                else
                    group_t = zh_groups;

                if (group_t.AsEnumerable().Where(c => c.Field<string>("tz").Trim().Equals("主观题")).Count() == 0)
                {
                    _mdata.wizard.CheckData(1, "数据分析题组中没有主观题，默认按无选择题选项的题目为主观题？");
                    foreach (DataRow dr in _standard_ans.Rows)
                        if (dr["da"].ToString().Trim().Equals(""))
                            sub_tz.Add(dr["th"].ToString().Trim());
                }
                else
                    th_parse(group_t.AsEnumerable()
                        .Where(c => c.Field<string>("tz").Trim().Equals("主观题"))
                        .First().Field<string>("th").Trim(), sub_tz);
            }
            newStandard = StandardAnsRecontruction(_standard_ans, name_list);

            _basic_data.Columns.Add("kh", System.Type.GetType("System.String"));
            _basic_data.Columns.Add("zkzh", System.Type.GetType("System.String"));
            _basic_data.Columns.Add("xm", System.Type.GetType("System.String"));
            _basic_data.Columns.Add("sfzjh", System.Type.GetType("System.String"));
            _basic_data.Columns.Add("xjh", System.Type.GetType("System.String"));
            _basic_data.Columns.Add("xxdm", System.Type.GetType("System.String"));
            _basic_data.Columns.Add("totalmark", typeof(decimal));
            _basic_data.Columns.Add("xb", typeof(string));
            _basic_data.Columns.Add("Groups", typeof(string));
            _basic_data.Columns.Add("qxdm", typeof(string));
            //if (has_xz)
            //    _basic_data.Columns.Add("X", typeof(string));
            for (i = 0; i < newStandard.Rows.Count; i++)
                _basic_data.Columns.Add("T" + ((string)newStandard.Rows[i]["th"]).Trim(), System.Type.GetType("System.Decimal"));

            for (i = 0; i < _standard_ans.Rows.Count; i++)
                if (!_standard_ans.Rows[i]["da"].ToString().Trim().Equals(""))
                    _basic_data.Columns.Add("D" + ((string)_standard_ans.Rows[i]["th"]).Trim(), typeof(string));

            foreach (string xz in _mdata.xz)
                _basic_data.Columns.Add("X" + xz, typeof(string));
            Hashtable Multi_ans = new Hashtable(); 
            foreach (DataRow dr in newStandard.Rows)
            {
                if (dr["da"].ToString().Contains('('))
                {
                    Hashtable hs = new Hashtable();
                    string da_str = dr["da"].ToString().Trim();
                    foreach (Match match in Regex.Matches(da_str, @"([\w|@|:|;|<|=|>|?],[\w|\.]+)"))
                    {
                        string[] temp = match.ToString().Split(new char[] {','});
                        if(temp.Length != 2)
                            throw new ArgumentException("标准答案中第"+ dr["th"].ToString() + "题答案格式不对，括号内应为两个值，由逗号隔开");
                        if (Utils.choiceTransfer(temp[0]) == null)
                            throw new ArgumentException("标准答案中第" + dr["th"].ToString() + "题第一个值应为答案，该答案不存在");
                        decimal mark = 0;
                        try
                        {
                            mark = Convert.ToDecimal(temp[1]);
                            if(mark > Convert.ToDecimal(dr["fs"]))
                                throw new ArgumentException("标准答案中第" + dr["th"].ToString() + "题第二个值应为得分，不能大于满分");
                        }
                        catch (FormatException e)
                        {
                            throw new ArgumentException("标准答案中第" + dr["th"].ToString() + "题第二个值应为得分，该得分无效");
                        }
                        hs.Add(temp[0], mark);
                        if (mark == Convert.ToDecimal(dr["fs"]))
                            dr["da"] = temp[0];
                    }
                    Multi_ans.Add(dr["th"], hs);

                }
            }

            foreach(DataRow dr in _standard_ans.Rows)
            {
                if(!dt.Columns.Contains("T" + dr["th"].ToString().Trim()))
                    throw new ArgumentException("数据库中不存在题号为" + "T" + dr["th"].ToString().Trim() + "数据");
                
            }
            int rowline = 0;
            foreach (DataRow dr in dt.Rows)
            {
                DataRow newRow = _basic_data.NewRow();
                newRow["kh"] = dr["kh"].ToString().Trim();
                newRow["zkzh"] = dr["zkzh"].ToString().Trim();
                newRow["xxdm"] = dr["xxdm"].ToString().Trim();
                newRow["xm"] = dt.Columns.Contains("xm")? dr["xm"].ToString().Trim(): "";
                newRow["sfzjh"] = dt.Columns.Contains("sfzjh")? dr["sfzjh"].ToString().Trim() : "";
                newRow["xjh"] = dt.Columns.Contains("xjh")? dr["xjh"].ToString().Trim() : "";
                newRow["totalmark"] = 0;
                newRow["xb"] = dr["xb"].ToString().Trim();
                newRow["Groups"] = "";
                newRow["qxdm"] = dr["qxdm"].ToString().Trim();
                decimal obj_mark = 0;
                
                int total_count = 0;
                
                foreach (DataRow ans_dr in newStandard.Rows)
                {
                    if (ans_dr["da"].ToString().Trim().Equals(""))
                    {
                        if (name_list[total_count] == null)
                        {
                            if ((decimal)dr["T" + (string)ans_dr["th"]] > Convert.ToDecimal(ans_dr["fs"]))
                                throw new ArgumentException("第" + (string)ans_dr["th"] + "题满分值小于实际分值！");
                            newRow["T" + (string)ans_dr["th"]] = (decimal)dr["T" + (string)ans_dr["th"]];
                            //sub_mark += (decimal)dr["T" + (string)ans_dr["th"]];
                            newRow["totalmark"] = (decimal)newRow["totalmark"] + (decimal)dr["T" + (string)ans_dr["th"]];
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

                        string th = "T" + ((string)ans_dr["th"]).Trim();
                        string temp_ans = dr[th].ToString().Trim();
                        //if (!choice.IsMatch(temp_ans))
                        //    throw new ArgumentException("考号" + dr["kh"] + "的第" + ((string)ans_dr["th"]).Trim() + "题的答案为：" + temp_ans + " 不属于字母组合！");
                        if (Multi_ans.Contains(ans_dr["th"]))
                        {
                            Hashtable hs_temp = (Hashtable)Multi_ans[ans_dr["th"]];
                            decimal val;
                            if (hs_temp.Contains(temp_ans))
                                val = (decimal)hs_temp[temp_ans];
                            else
                                val = 0;
                            newRow[th] = val;
                            obj_mark += val;
                            newRow["totalmark"] = (decimal)newRow["totalmark"] + val;

                        }
                        else
                        {
                            string temp = ((string)ans_dr["da"]).Trim();

                            if (temp_ans.Equals(temp))
                            {
                                decimal val = Convert.ToDecimal(ans_dr["fs"]);
                                newRow[th] = val;
                                obj_mark += val;
                                newRow["totalmark"] = (decimal)newRow["totalmark"] + val;

                            }
                            else if (_mdata.PartialRight != 0 && Utils.isContain(temp, temp_ans))
                            {
                                if (_mdata.PartialRight > Convert.ToDecimal(ans_dr["fs"]))
                                    throw new ArgumentException("选择题半分分数大于满分分数！");

                                decimal val = _mdata.PartialRight;
                                newRow[th] = val;
                                obj_mark += val;
                                newRow["totalmark"] = (decimal)newRow["totalmark"] + val;


                            }
                            else
                                newRow[th] = 0.0;
                        }
                        newRow["D" + ((string)ans_dr["th"]).Trim()] = temp_ans;

                    }
                    total_count++;
                }
                    
                foreach(string xz in _mdata.xz)
                    newRow["X" + xz] = dr["X" + xz];
                
                if (_mdata.fullmark_iszero && (decimal)newRow["totalmark"] == 0)
                    continue;

                if (_mdata.sub_iszero)
                {
                    decimal sub_mark = 0;
                    foreach (string sub_th in sub_tz)
                    {

                        sub_mark += (decimal)newRow["T" + sub_th];
                    }
                    if (sub_mark == 0)
                        continue;
                }
                //if (_mdata.sub_iszero && sub_mark == 0)
                //    continue;
               
                //if (has_xz)
                //    newRow["X"] = dr["X"].ToString().Trim();
                _basic_data.Rows.Add(newRow);

                if (rowline % 5000 == 0 && GC.GetTotalMemory(true) > 60000000)
                    ClearMemory();
                
                rowline++;
            }

           
            _basic_data.DefaultView.Sort = "totalmark asc";
            _basic_data = _basic_data.DefaultView.ToTable();
            int totalsize = _basic_data.Rows.Count;
            if (_gtype.Equals(ZK_database.GroupType.population))
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
            create_groups();
            //if (Utils.saveMidData)
            //{
            //    Utils.create_groups_table(_basic_data, Utils.year + "高考" + Utils.subject + "基础数据");
            //    Utils.create_groups_table(_group_data, Utils.year + "高考" + Utils.subject + "题组数据");

            //}
            return "";
        }
        [DllImport("kernel32.dll", EntryPoint = "SetProcessWorkingSetSize")]
        public static extern int SetProcessWorkingSetSize(IntPtr process, int minSize, int maxSize);
        /// <summary>
        /// 释放内存
        /// </summary>
        public static void ClearMemory()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            if (Environment.OSVersion.Platform == PlatformID.Win32NT)
            {
                SetProcessWorkingSetSize(System.Diagnostics.Process.GetCurrentProcess().Handle, -1, -1);
            }
        }
        
        public char[] TransferCharArray(string[] ans_group)
        {
            char[] newgroup = new char[ans_group.Length];
            for(int i = 0; i < ans_group.Length; i++)
            {
                newgroup[i] = Convert.ToChar(ans_group[i]);
            }
            return newgroup;
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
        
        
        public void create_groups()
        {
            #region divide the table into groups
            //StringBuilder objectdata = new StringBuilder();
            _group_data.Columns.Add("kh", System.Type.GetType("System.String"));
            _group_data.Columns.Add("zkzh", System.Type.GetType("System.String"));
            _group_data.Columns.Add("xxdm", System.Type.GetType("System.String"));
            _group_data.Columns.Add("totalmark", System.Type.GetType("System.Decimal"));
            _group_data.Columns.Add("xb", typeof(string));
            List<List<string>> tm = new List<List<string>>();
            
            for (int i = 0; i < _groups.Rows.Count; i++)
            {
                List<string> tz = new List<string>();
                //string row_name = _groups.Rows[i][0].ToString().Trim();
                _group_data.Columns.Add("FZ"+(i+1).ToString(), System.Type.GetType("System.Decimal"));
                string org = _groups.Rows[i][1].ToString().Trim();
                th_parse(org, tz);
                
                tm.Add(tz);
            }
            _group_data.Columns.Add("Groups", typeof(string));
            _group_data.Columns.Add("qxdm", typeof(string));
            foreach (DataRow dr in _basic_data.Rows)
            {
                DataRow newRow = _group_data.NewRow();
                newRow["kh"] = ((string)dr["kh"]).Trim();
                newRow["zkzh"] = ((string)dr["zkzh"]).Trim();
                newRow["xxdm"] = ((string)dr["xxdm"]).Trim();
                newRow["Groups"] = ((string)dr["Groups"]).Trim();
                newRow["qxdm"] = dr["qxdm"].ToString().Trim();
                newRow["xb"] = dr["xb"].ToString().Trim();
                newRow["totalmark"] = dr["totalmark"];
                int j;
                for (j = 0; j < _groups.Rows.Count; j++)
                {
                    decimal count_ = 0;
                    foreach (string s in tm[j])
                    {
                        count_ += (decimal)dr["T" + s];
                    }
                    newRow[j + 5] = count_;
                }
                _group_data.Rows.Add(newRow);
            }
            
            #endregion
        }


        public void xz_check(DataTable db, List<string> xz)
        {
            foreach (string temp in xz)
            {
                if (!db.Columns.Contains("X" + temp))
                    throw new ArgumentException("标答所记录选做题" + temp + "在数据库中不存在！");
            }
        }

        public void th_parse(string org, List<string> tz)
        {
            string spattern = "^\\d+~\\d+$";
            string[] org_char = org.Split(new char[2] { ',', '，' });
            foreach (string th in org_char)
            {

                if (!th.Equals(""))
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
                            tz.Add(j.ToString().Trim());
                        }

                    }
                    else
                        tz.Add(th);
                }
            }
        }
    }
}
