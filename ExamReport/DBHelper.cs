using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Configuration;
using MySql.Data;
using MySql.Data.MySqlClient;
using System.Data.OleDb;

namespace ExamReport
{
    public static class DBHelper
    {
        public static void create_fz_table(string name, DataTable dt, Dictionary<string, List<string>> groups_group)
        {
            name += "_fz";
            DataTable copy = dt.Copy();
            copy.Columns.Add("id", typeof(string));
            copy.Columns.Add("dz", typeof(string));
            int count = 0;
            foreach (string key in groups_group.Keys)
            {
                List<string> groups = groups_group[key];
                foreach (string group in groups)
                {
                    copy.Rows[count]["dz"] = key;
                    copy.Rows[count]["id"] = "FZ" + (count + 1).ToString();
                    count++;
                }
            }
            create_mysql_table(copy, name, "100", "decimal(4,1)");
        }

        public static void create_ans_table(string name, DataTable dt, List<string> xz)
        {
            name += "_ans";
            DataTable copy = dt.Copy();
            copy.Columns.Add("xz", typeof(string));
            foreach (DataRow dr in copy.Rows)
            {
                if (xz.Contains(dr["th"].ToString().Trim()))
                    dr["xz"] = 1;
                else
                    dr["xz"] = 0;
            }
            create_mysql_table(copy, name);
        }

        public static DataTable get_fz(string name, ref Dictionary<string, List<string>> groups_group)
        {
            DataTable data = MySqlHelper.GetDataSet(MySqlHelper.Conn, CommandType.Text, "select * from "
                + name, null).Tables[0];
            groups_group = data.AsEnumerable().GroupBy(c => c.Field<string>("dz")).Select(c => new
            {
                key = c.Key,
                value = c.Select(p => p.Field<string>("tz")).ToList()
            }).ToDictionary(c => c.key, c => c.value);

            data.Columns.Remove(data.Columns["dz"]);
            return data;
        }

        public static DataTable get_ans(string name, ref List<string> xz)
        {
            DataTable data = get_only_ans(name);

            xz = data.AsEnumerable().Where(c => c.Field<string>("xz").Equals("1")).Select(c => c.Field<string>("th")).ToList();
            
            data.Columns.Remove(data.Columns["xz"]);
            return data;

        }
        public static DataTable get_only_ans(string name)
        {
            return MySqlHelper.GetDataSet(MySqlHelper.Conn, CommandType.Text, "select * from "
                + name, null).Tables[0];
        }
        public static bool delete_row(string year, string exam, string sub)
        {
            string _exam = Utils.language_trans(exam);
            string _sub = Utils.language_trans(sub);

            //检查是否存在这一条数据
            MySqlDataReader reader = MySqlHelper.ExecuteReader(MySqlHelper.Conn, CommandType.Text, "select * from exam_meta_data where year='"
                + year + "' and exam='"
                + _exam + "' and sub='"
                + _sub + "'", null);
            if(!reader.Read())
                throw new ArgumentException("数据库不一致，该条数据不存在");
            //int val = MySqlHelper.ExecuteNonQuery(MySqlHelper.Conn, CommandType.Text, "select * from exam_meta_data where year='"
            //    + year + "' and exam='"
            //    + _exam + "' and sub='"
            //    + _sub + "'", null);
            //if (val <= 0)
            //    throw new ArgumentException("数据库不一致，该条数据不存在");
            MySqlConnection conn = new MySqlConnection(MySqlHelper.Conn);
            conn.Open();

            MySqlTransaction trans = conn.BeginTransaction();
            

            MySqlHelper.ExecuteNonQuery(trans, CommandType.Text, "delete from exam_meta_data where year='"
                + year + "' and exam='"
                + _exam + "' and sub='"
                + _sub + "'", null);
            if (_sub.Equals("zf") || _sub.EndsWith("xz"))
            {
                MySqlHelper.ExecuteNonQuery(trans, CommandType.Text, "drop table " + Utils.get_zt_tablename(year, _exam, _sub), null);
            }
            else
            {
                MySqlHelper.ExecuteNonQuery(trans, CommandType.Text, "drop table " + Utils.get_basic_tablename(year, _exam, _sub), null);
                MySqlHelper.ExecuteNonQuery(trans, CommandType.Text, "drop table " + Utils.get_group_tablename(year, _exam, _sub), null);
                MySqlHelper.ExecuteNonQuery(trans, CommandType.Text, "drop table " + Utils.get_ans_tablename(year, _exam, _sub), null);
                MySqlHelper.ExecuteNonQuery(trans, CommandType.Text, "drop table " + Utils.get_fz_tablename(year, _exam, _sub), null);

                if (_exam.Equals("gk") && (Utils.language_trans(_sub).Contains("理综") || Utils.language_trans(_sub).Contains("文综")))
                {
                    MySqlHelper.ExecuteNonQuery(trans, CommandType.Text, "drop table " + "zh_" + Utils.get_basic_tablename(year, _exam, _sub), null);
                    MySqlHelper.ExecuteNonQuery(trans, CommandType.Text, "drop table " + "zh_" + Utils.get_group_tablename(year, _exam, _sub), null);
                    MySqlHelper.ExecuteNonQuery(trans, CommandType.Text, "drop table " + "zh_" + Utils.get_ans_tablename(year, _exam, _sub), null);
                    MySqlHelper.ExecuteNonQuery(trans, CommandType.Text, "drop table " + "zh_" + Utils.get_fz_tablename(year, _exam, _sub), null);

                }
            }
            trans.Commit();
            conn.Close();

            //if(val == 0)
            //    throw new ArgumentException("数据库不一致，该条数据不存在");
            return true;
        }
        public static void create_mysql_table(DataTable groups_data, string filename, string charsize, string datastype)
        {
            MySqlHelper.ExecuteNonQuery(MySqlHelper.Conn, CommandType.Text, "drop table if exists " + filename, null);
            StringBuilder objectdata = new StringBuilder();
            objectdata.Clear();
            int i = 0;

            objectdata.Append("CREATE TABLE " + filename + " (\n");
            int count = 0;
            foreach (DataColumn dc in groups_data.Columns)
            {
                objectdata.Append("\t" + dc.ColumnName + " ");
                if (dc.DataType.ToString().Equals("System.String"))
                    objectdata.Append("text");
                else if (dc.DataType.ToString().Equals("System.Decimal"))
                    objectdata.Append(datastype);
                else if (dc.DataType.ToString().Equals("System.Int32"))
                    objectdata.Append("int");
                else
                    i++;
                count++;
                if (count != groups_data.Columns.Count)
                    objectdata.Append(",\n");
                else
                    objectdata.Append(")");
            }

            MySqlHelper.ExecuteNonQuery(MySqlHelper.Conn, CommandType.Text, objectdata.ToString(), null);
            MySqlTransaction tx = null;
            using (MySqlConnection conn = new MySqlConnection(MySqlHelper.Conn))
            {
                conn.Open();
                tx = conn.BeginTransaction();
                
                int row_count = 0;
                foreach (DataRow dr in groups_data.Rows)
                {
                    objectdata.Clear();
                    objectdata.Append("INSERT INTO " + filename + " VALUES (");

                    for (i = 0; i < groups_data.Columns.Count; i++)
                    {
                        if (groups_data.Columns[i].DataType.ToString().Equals("System.String"))
                            objectdata.Append("'" + dr[i].ToString().Trim() + "'");
                        else if (groups_data.Columns[i].DataType.ToString().Equals("System.Decimal"))
                            objectdata.Append((decimal)dr[i]);
                        else if (groups_data.Columns[i].DataType.ToString().Equals("System.Int32"))
                            objectdata.Append((int)dr[i]);

                        if (i != groups_data.Columns.Count - 1)
                            objectdata.Append(",");
                        else
                            objectdata.Append(");");

                    }
                    MySqlHelper.ExecuteNonQuery(tx, CommandType.Text, objectdata.ToString(), null);

                    if (row_count % 500 == 0)
                    {
                        tx.Commit();
                        tx = conn.BeginTransaction();
                    }

                    row_count++;
                }

                tx.Commit();
                conn.Close();
               
            }
        }
        public static void create_mysql_table(DataTable groups_data, string filename)
        {
            

            string charsize = ConfigurationManager.AppSettings["charsize"].ToString().Trim();
            create_mysql_table(groups_data, filename, charsize, "decimal(5,2)");

            

        }

        public static void create_mysql_table_datastyle(DataTable groups_data, string filename)
        {
            string charsize = ConfigurationManager.AppSettings["charsize"].ToString().Trim();
            create_mysql_table(groups_data, filename, charsize, "decimal(5,2)");
        }
    }
}
