﻿using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using System.Data;
using MySql.Data;
using MySql.Data.MySqlClient;

namespace ExamReport
{
    public class MetaData
    {
        public string _year;
        public string _exam;
        public string _sub;

        public string ywyy_choice;
        public string zh_choice;

        public string log_name;

        public decimal _fullmark;
        public decimal _sub_fullmark;
        public ZK_database.GroupType _grouptype;
        public int _group_num = 0;

        public bool fullmark_iszero = true;
        public bool sub_iszero = true;
        public decimal PartialRight = 0;

        public List<ArrayList> CJ_list;
        public List<ArrayList> QXSF_list;
        public List<ArrayList> SF_list;

        public DataTable ans;
        public DataTable grp;

        public DataTable basic;
        public DataTable group;

        public DataTable zh_basic;
        public DataTable zh_group;

        public DataTable zh_ans;
        public DataTable zh_grp;


        public bool isSampled;

        public Dictionary<string, List<string>> zh_groups_group;

        public List<string> xz;
        public Dictionary<string, List<string>> groups_group;

        public MyWizard wizard;
        public MetaData(string year, string exam, string sub, bool isLoading)
        {
            if (!isLoading)
            {
                MySqlDataReader reader = MySqlHelper.ExecuteReader(
                    MySqlHelper.Conn,
                    CommandType.Text,
                    "select * from exam_meta_data where year='"
                    + year + "' and exam='"
                    + exam + "' and sub='"
                    + sub + "'", null);

                if (reader.Read())
                    _exam = exam;
                else
                {
                    //reader = MySqlHelper.ExecuteReader(
                    //    MySqlHelper.Conn, 
                    //    CommandType.Text, 
                    //    "select * from exam_meta_data where year='"
                    //    + year + "' and exam='"
                    //    + "ngk" + "' and sub='"
                    //    + sub + "'", null);
                    //if (!reader.Read())
                    //    throw new Exception("数据库异常，该数据不存在");
                    _exam = "ngk";
                }
            }
            else
                _exam = exam;
            _year = year;
            if (_exam.Equals("gk"))
                _sub = Utils.language_trans(sub);
            else if (_exam.Equals("hk") || _exam.Equals("ngk"))
                _sub = Utils.hk_lang_trans(sub);
            else if(_exam.Equals("zk"))
                _sub = Utils.hk_lang_trans(sub);
            else
                _sub = Utils.language_trans(sub);
        }
        public bool check()
        {
            MySqlDataReader reader = MySqlHelper.ExecuteReader(MySqlHelper.Conn, CommandType.Text, "select * from exam_meta_data where year='"
                + _year + "' and exam='"
                + _exam + "' and sub='"
                + Utils.language_trans(_sub) + "'", null);
            if (reader.Read())
                return false;
            return true;
        }
        public bool insert_data()
        {
            //检查是否已存在该数据
            int val = MySqlHelper.ExecuteNonQuery(MySqlHelper.Conn, CommandType.Text, "insert into exam_meta_data (year,exam,sub,ans,grp,fullmark,zh,gtype,gnum) values ('"
                + _year + "', '"
            + _exam + "','"
            + Utils.language_trans(_sub) + "','1','1',"
            + Convert.ToInt32(_fullmark).ToString() + ","
            + Convert.ToInt32(_sub_fullmark).ToString() + ",'" 
            + gtype_to_string(_grouptype) + "'," 
            + Convert.ToInt32(_group_num).ToString() + ")", null);
            if (val <= 0)
                throw new Exception("未知错误，数据库写入错误");

            return true;
        }
        public void rollback()
        {
            MySqlDataReader reader = MySqlHelper.ExecuteReader(MySqlHelper.Conn, CommandType.Text, "select * from exam_meta_data where year='"
                + _year + "' and exam='"
                + _exam + "' and sub='"
                + Utils.language_trans(_sub) + "'", null);
            if (reader.Read())
            {
                MySqlHelper.ExecuteNonQuery(MySqlHelper.Conn, CommandType.Text, "delete from exam_meta_data where year='"
                + _year + "' and exam='"
                + _exam + "' and sub='"
                + Utils.language_trans(_sub) + "'", null);
            }

            MySqlHelper.ExecuteNonQuery(MySqlHelper.Conn, CommandType.Text, "drop table if exists " + 
                Utils.get_ans_tablename(_year, _exam, Utils.language_trans(_sub)), null);
            MySqlHelper.ExecuteNonQuery(MySqlHelper.Conn, CommandType.Text, "drop table if exists " +
                Utils.get_fz_tablename(_year, _exam, Utils.language_trans(_sub)), null);
            MySqlHelper.ExecuteNonQuery(MySqlHelper.Conn, CommandType.Text, "drop table if exists " +
                Utils.get_basic_tablename(_year, _exam, Utils.language_trans(_sub)), null);
            MySqlHelper.ExecuteNonQuery(MySqlHelper.Conn, CommandType.Text, "drop table if exists " +
                Utils.get_group_tablename(_year, _exam, Utils.language_trans(_sub)), null);

            MySqlHelper.ExecuteNonQuery(MySqlHelper.Conn, CommandType.Text, "drop table if exists " + "zh_" +
                Utils.get_ans_tablename(_year, _exam, Utils.language_trans(_sub)), null);
            MySqlHelper.ExecuteNonQuery(MySqlHelper.Conn, CommandType.Text, "drop table if exists " + "zh_" +
                Utils.get_fz_tablename(_year, _exam, Utils.language_trans(_sub)), null);
            MySqlHelper.ExecuteNonQuery(MySqlHelper.Conn, CommandType.Text, "drop table if exists " + "zh_" +
                Utils.get_basic_tablename(_year, _exam, Utils.language_trans(_sub)), null);
            MySqlHelper.ExecuteNonQuery(MySqlHelper.Conn, CommandType.Text, "drop table if exists " + "zh_" +
                Utils.get_group_tablename(_year, _exam, Utils.language_trans(_sub)), null);

            MySqlHelper.ExecuteNonQuery(MySqlHelper.Conn, CommandType.Text, "drop table if exists " +
                Utils.get_zt_tablename(_year, _exam, Utils.language_trans(_sub)), null);


        }
        public bool check(string sub)
        {
            switch (sub)
            {
                case "wl":
                case "hx":
                case "sw":
                case "zz":
                case "dl":
                case "ls":
                    return true;
                default:
                    return false;
            }
        }

        public void get_meta_data()
        {
            MySqlDataReader reader = MySqlHelper.ExecuteReader(MySqlHelper.Conn, CommandType.Text, "select * from exam_meta_data where year='"
                + _year + "' and exam='"
                + _exam + "' and sub='"
                + Utils.language_trans(_sub) + "'", null);
            if (!reader.Read())
            {
                throw new Exception("数据库异常，不存在该数据");
            }
            _fullmark = Convert.ToDecimal(reader["fullmark"]);
            _grouptype = string_to_gtype(reader["gtype"].ToString().Trim());
            _group_num = Convert.ToInt32(reader["gnum"]);

            if(_exam.Equals("gk") && check(Utils.language_trans(_sub)))
                _sub_fullmark = Convert.ToDecimal(reader["zh"]);
                
        }

        public List<string> get_column_name()
        {
            MySqlDataReader reader = MySqlHelper.ExecuteReader(MySqlHelper.Conn, CommandType.Text, "select * from exam_meta_data where year='"
                + _year + "' and exam='"
                + _exam + "' and sub='"
                + Utils.language_trans(_sub) + "'", null);
            if (!reader.Read())
                throw new Exception("数据库异常，不存在该数据");
            string table_name = Utils.get_basic_tablename(_year, _exam, Utils.language_trans(_sub));
            if (_sub.Contains("理综") || _sub.Contains("文综"))
#pragma warning disable CS1717 // 对同一变量进行赋值；是否希望对其他变量赋值?
                table_name = table_name;//留个地方
#pragma warning restore CS1717 // 对同一变量进行赋值；是否希望对其他变量赋值?
            else if (_sub.Equals("总分") || _sub.Contains("行政版"))
                table_name = Utils.get_zt_tablename(_year, _exam, Utils.language_trans(_sub));
            reader = MySqlHelper.ExecuteReader(MySqlHelper.Conn, CommandType.Text, "select COLUMN_NAME from information_schema.COLUMNS where table_name = '" + table_name + "'", null);
            List<string> name = new List<string>();
            while (reader.Read())
            {
                name.Add(reader["COLUMN_NAME"].ToString());
            }
            return name;

        }

        public string get_column_type(string column_name)
        {
            MySqlDataReader reader = MySqlHelper.ExecuteReader(MySqlHelper.Conn, CommandType.Text, "select * from exam_meta_data where year='"
                + _year + "' and exam='"
                + _exam + "' and sub='"
                + Utils.language_trans(_sub) + "'", null);
            if (!reader.Read())
                throw new Exception("数据库异常，不存在该数据");
            string table_name = Utils.get_basic_tablename(_year, _exam, Utils.language_trans(_sub));
            if (_sub.Contains("理综") || _sub.Contains("文综"))
#pragma warning disable CS1717 // 对同一变量进行赋值；是否希望对其他变量赋值?
                table_name = table_name;//留个地方
#pragma warning restore CS1717 // 对同一变量进行赋值；是否希望对其他变量赋值?
            else if (_sub.Equals("总分") || _sub.Contains("行政版"))
                table_name = Utils.get_zt_tablename(_year, _exam, Utils.language_trans(_sub));
            reader = MySqlHelper.ExecuteReader(MySqlHelper.Conn, CommandType.Text, "describe " + table_name + " " + column_name, null);

            if(!reader.Read())
                throw new Exception("数据库异常，不存在该列");
            return reader["type"].ToString();
        }
        private string gtype_to_string(ZK_database.GroupType type)
        {
            switch (type)
            {
                case ZK_database.GroupType.population:
                    return "p";
                case ZK_database.GroupType.totalmark:
                    return "m";
                default:
                    return "n";
            }
        }

        private ZK_database.GroupType string_to_gtype(string type)
        {
            switch (type)
            {
                case "p":
                    return ZK_database.GroupType.population;
                case "m":
                    return ZK_database.GroupType.totalmark;
                default:
                    return ZK_database.GroupType.population;
            }
        }
        public string get_zh()
        {
            switch (_sub)
            {
                case "理综-化学":
                case "化学":
                case "理综-物理":
                case "物理":
                case "理综-生物":
                case "生物":
                    return "lz";
                case "文综-政治":
                case "政治":
                case "文综-地理":
                case "地理":
                case "文综-历史":
                case "历史":
                    return "wz";
                default:
                    return "unkown";
            }
        }
        public string get_cn_sub()
        {
            return _sub;
        }
        public string get_sub()
        {
            string sub;
            if (_exam.Equals("gk"))
                sub = Utils.language_trans(_sub);
            else if (_exam.Equals("hk") || _exam.Equals("ngk"))
                sub = Utils.hk_lang_trans(_sub);
            else if (_exam.Equals("zk"))
                sub = Utils.hk_lang_trans(_sub);
            else
                sub = Utils.language_trans(_sub);

            return sub;

        }
        public void get_SF_data(string addr)
        {
            SF_list = get_excel_data(addr);
        }
        public void get_CJ_data(string addr)
        {
            CJ_list = get_excel_data(addr);
        }
        public void get_QXSF_data(string addr)
        {
            QXSF_list = get_excel_data(addr);
        }
        
        public List<ArrayList> get_excel_data(string addr)
        {
            excel_process pro = new excel_process(addr);
            return pro.getData();
        }

        public void calculateFullmark()
        {
            int fmark = 0;
            foreach (DataRow row in ans.Rows)
            {
                if (((string)row["sample"]).Equals("1"))
                    fmark += Convert.ToInt32(row["fs"]);
            }
            _fullmark = fmark;
        }
        public void get_ans()
        {
            xz = new List<string>();
            ans = DBHelper.get_ans(Utils.get_ans_tablename(_year, _exam, Utils.language_trans(_sub)), ref xz);
        }
        public void get_sample_ans()
        {
            xz = new List<string>();
            ans = DBHelper.get_ans(Utils.get_sample_ans_tablename(_year, _exam, Utils.language_trans(_sub)), ref xz);
            calculateFullmark();
        }
        public void get_fz()
        {
            groups_group = new Dictionary<string, List<string>>();
            grp = DBHelper.get_fz(Utils.get_fz_tablename(_year, _exam, Utils.language_trans(_sub)), ref groups_group);
        }
        public void get_sample_fz()
        {
            groups_group = new Dictionary<string, List<string>>();
            grp = DBHelper.get_fz(Utils.get_sample_fz_tablename(_year, _exam, Utils.language_trans(_sub)), ref groups_group);
        }
        public void get_zf_data()
        {
            basic = get_mysql_table(Utils.get_zt_tablename(_year, _exam, Utils.language_trans(_sub)));
        }
        public void get_basic_data()
        {
            basic = get_mysql_table(Utils.get_basic_tablename(_year, _exam, Utils.language_trans(_sub)));
        }
        public void get_basic_sample_data()
        {
            basic = get_mysql_table(Utils.get_basic_tablename(_year, "ngk", Utils.language_trans(_sub))+"_sample");
        }

        public void get_group_data()
        {
            group = get_mysql_table(Utils.get_group_tablename(_year, _exam, Utils.language_trans(_sub)));
        }

        public void get_group_sample_data()
        {
            group = get_mysql_table(Utils.get_group_tablename(_year, "ngk", Utils.language_trans(_sub))+"_sample");
        }

        public void get_zh_basic_data()
        {
            zh_basic = get_mysql_table("zh_" + Utils.get_basic_tablename(_year, _exam, Utils.language_trans(_sub)));
        }
        public void get_zh_group_data()
        {
            zh_group = get_mysql_table("zh_" + Utils.get_group_tablename(_year, _exam, Utils.language_trans(_sub)));
        }
        public void get_zh_ans()
        {
            zh_ans = DBHelper.get_only_ans("zh_" + Utils.get_ans_tablename(_year, _exam, Utils.language_trans(_sub)));
        }
        public void get_zh_fz()
        {
            zh_groups_group = new Dictionary<string, List<string>>();
            zh_grp = DBHelper.get_fz("zh_" + Utils.get_fz_tablename(_year, _exam, Utils.language_trans(_sub)), ref zh_groups_group);
        }
        public DataTable get_mysql_table(string name)
        {
            return MySqlHelper.GetDataSet(MySqlHelper.Conn, CommandType.Text, "select * from " + name, null).Tables[0];
        }
        public string get_basic_table_name()
        {
            return Utils.get_basic_tablename(_year, _exam, Utils.language_trans(_sub));
        }
        public string get_group_table_name()
        {
            return Utils.get_group_tablename(_year, _exam, Utils.language_trans(_sub));
        }
        public string get_dbf_table_name()
        {
            return _year + "_" + _exam + "_" + Utils.language_trans(_sub) + ".dbf";
        }
        public string get_dbf_sample_table_name()
        {
            return _year + "_" + _exam + "_" + Utils.language_trans(_sub) + "_sample" + ".dbf";
        }
    }
}
