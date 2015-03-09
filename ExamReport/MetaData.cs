using System;
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

        public string log_name;

        public decimal _fullmark;
        public ZK_database.GroupType _grouptype;
        public int _group_num;

        public List<ArrayList> CJ_list;
        public List<ArrayList> QX_list;

        public DataTable ans;
        public DataTable grp;

        public DataTable basic;
        public DataTable group;

        public List<string> xz;
        public Dictionary<string, List<string>> groups_group;
        public MetaData(string year, string exam, string sub)
        {
            _year = year;
            _exam = exam;
            _sub = sub;
        }

        public bool insert_data()
        {
            //检查是否已存在该数据
            MySqlDataReader reader = MySqlHelper.ExecuteReader(MySqlHelper.Conn, CommandType.Text, "select * from exam_meta_data where year='"
                + _year + "' and exam='"
                + _exam + "' and sub='"
                + _sub + "'", null);
            if (reader.Read())
                throw new DuplicateNameException();


            int val = MySqlHelper.ExecuteNonQuery(MySqlHelper.Conn, CommandType.Text, "insert into exam_meta_data (year,exam,sub,ans,grp,fullmark,zh,gtype,gnum) values ('"
                + _year + "', '"
            + _exam + "','"
            + _sub + "','1','1',"
            + Convert.ToInt32(_fullmark).ToString() + ",'"
            + check(_sub) + "','" 
            + gtype_to_string(_grouptype) + "'," 
            + Convert.ToInt32(_group_num).ToString() + ")", null);
            if (val <= 0)
                throw new Exception("未知错误，数据库写入错误");

            return true;
        }

        private string check(string sub)
        {
            switch (sub)
            {
                case "wl":
                case "hx":
                case "sw":
                case "zz":
                case "dl":
                case "ls":
                    return "1";
                default:
                    return "0";
            }
        }

        public void get_meta_data()
        {
            MySqlDataReader reader = MySqlHelper.ExecuteReader(MySqlHelper.Conn, CommandType.Text, "select * from exam_meta_data where year='"
                + _year + "' and exam='"
                + _exam + "' and sub='"
                + _sub + "'", null);
            if (!reader.Read())
                throw new Exception("数据库异常，不存在该数据");
            _fullmark = Convert.ToDecimal(reader["fullmark"]);
            _grouptype = string_to_gtype(reader["gtype"].ToString().Trim());
            _group_num = Convert.ToInt32(reader["gnum"]);
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
                    return "p";
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

        public void get_CJ_data(string addr)
        {
            CJ_list = get_excel_data(addr);
        }
        public void get_QX_data(string addr)
        {
            QX_list = get_excel_data(addr);
        }
        public List<ArrayList> get_excel_data(string addr)
        {
            excel_process pro = new excel_process(addr);
            return pro.getData();
        }
        public void get_ans()
        {
            xz = new List<string>();
            ans = DBHelper.get_ans(Utils.get_ans_tablename(_year, _exam, _sub), ref xz);
        }

        public void get_fz()
        {
            groups_group = new Dictionary<string, List<string>>();
            grp = DBHelper.get_fz(Utils.get_fz_tablename(_year, _exam, _sub), ref groups_group);
        }
        public void get_basic_data()
        {
            basic = get_mysql_table(Utils.get_basic_tablename(_year, _exam, _sub));
        }

        public void get_group_data()
        {
            group = get_mysql_table(Utils.get_group_tablename(_year, _exam, _sub));
        }
        public DataTable get_mysql_table(string name)
        {
            return MySqlHelper.GetDataSet(MySqlHelper.Conn, CommandType.Text, "select * from " + name, null).Tables[0];
        }
    }
}
