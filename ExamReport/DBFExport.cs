using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.OleDb;
using System.Text.RegularExpressions;

namespace ExamReport
{
    public class DBFExport
    {
        DataTable summary;
        DataTable cj_data;
        DataTable sf_data;
        DataTable _basic;
        public DBFExport(DataTable basic, DataTable groups)
        {
            _basic = basic;
            summary = basic.AsEnumerable().Join(
                groups.AsEnumerable(),
                a => a.Field<string>("zkzh"),
                b => b.Field<string>("zkzh"),
                (a, b) => new { a, b }).ToDataTable();
        }

        public void join_cj(List<ArrayList> cj_list)
        {
            DataTable cj = new DataTable();
            cj.Columns.Add("qxdm", typeof(string));
            cj.Columns.Add("countyid", typeof(string));

            for (int i = 1; i < cj_list[0].Count; i++)
            {
                DataRow dr = cj.NewRow();
                dr["qxdm"] = cj_list[0][i].ToString().Trim();
                dr["countyid"] = "c";
                cj.Rows.Add(dr);
            }

            for (int i = 1; i < cj_list[1].Count; i++)
            {
                DataRow dr = cj.NewRow();
                dr["qxdm"] = cj_list[1][i].ToString().Trim();
                dr["countyid"] = "j";
                cj.Rows.Add(dr);
            }

            var temp = from basic in _basic.AsEnumerable()
                       join c in cj.AsEnumerable()
                       on basic.Field<string>("qxdm") equals c.Field<string>("qxdm")
                       select new 
                       {
                           zkzh = basic.Field<string>("zkzh"),
                           countyid = c.Field<string>("countyid")
                       };

            cj_data = temp.ToDataTable();
            
        }

        public void join_sf(List<ArrayList> sf_list)
        {
            DataTable sf = new DataTable();
            sf.Columns.Add("xxdm", typeof(string));
            sf.Columns.Add("llsx", typeof(string));

            for (int i = 1; i < sf_list[0].Count; i++)
            {
                DataRow dr = sf.NewRow();
                dr["xxdm"] = sf_list[0][i].ToString().Trim();
                dr["llsx"] = "1";
                sf.Rows.Add(dr);
            }

            for (int i = 1; i < sf_list[1].Count; i++)
            {
                DataRow dr = sf.NewRow();
                dr["xxdm"] = sf_list[1][i].ToString().Trim();
                dr["llsx"] = "2";
                sf.Rows.Add(dr);
            }
            var temp = from basic in _basic.AsEnumerable()
                       join s in sf.AsEnumerable()
                       on basic.Field<string>("xxdm") equals s.Field<string>("xxdm") into result
                       from r in result.DefaultIfEmpty()
                       select new 
                       {
                           zkzh = basic.Field<string>("zkzh"),
                           llsx = r == null? "0":r.Field<string>("llsx")
                       };

            sf_data = temp.ToDataTable();

        }

        public void do_export(string addr, string name)
        {

            using (var dbconnection = new OleDbConnection(@"Provider=vfpoledb;Data Source=" + addr + ";Collating Sequence=machine;"))
            {
                StringBuilder objectdata = new StringBuilder();
                objectdata.Clear();
                Regex th = new Regex("^[Tt]\\d");
                Regex fz = new Regex("^[Ff][Zz]\\d");
                Regex dd = new Regex("^[Dd]\\d");

                objectdata.Append("CREATE TABLE " + name + " ( zkzh C(20), xxdm C(10), xb C(1), qxdm C(10), score N(4,1), groups C(4), countyid C(1), llsx C(1), qxsf C(1)");
                foreach (DataColumn dc in ((DataRow)summary.Rows[0]["a"]).Table.Columns)
                {
                    if (th.IsMatch(dc.ColumnName))
                        objectdata.Append(", " + dc.ColumnName + " N(4,1)");
                    if (dd.IsMatch(dc.ColumnName))
                        objectdata.Append(", " + dc.ColumnName + " C(10)");    
                }
                foreach (DataColumn dc in ((DataRow)summary.Rows[0]["b"]).Table.Columns)
                {
                    if (fz.IsMatch(dc.ColumnName))
                        objectdata.Append(", " + dc.ColumnName + " N(4,1)");
                }
                objectdata.Append(")");

                dbconnection.Open();
                OleDbCommand cmd = dbconnection.CreateCommand();
                cmd.CommandText = objectdata.ToString();
                cmd.ExecuteNonQuery();


                OleDbTransaction tx = dbconnection.BeginTransaction();

                int row_count = 0;
                foreach (DataRow ddr in summary.Rows)
                {
                    DataRow dr = (DataRow)ddr["a"];
                    objectdata.Clear();
                    objectdata.Append("INSERT INTO " + name + " VALUES (");
                    objectdata.Append("'" + dr["zkzh"].ToString().Trim() + "',");
                    objectdata.Append("'" + dr["xxdm"].ToString().Trim() + "',");
                    objectdata.Append("'" + dr["xb"].ToString().Trim() + "',");
                    objectdata.Append("'" + dr["qxdm"].ToString().Trim() + "',");
                    objectdata.Append(dr["totalmark"].ToString().Trim() + ",");
                    objectdata.Append("'" + dr["groups"].ToString().Trim() + "',");
                    objectdata.Append("'" + cj_data.Rows[row_count]["countyid"].ToString().Trim() + "',");
                    objectdata.Append("'" + sf_data.Rows[row_count]["llsx"].ToString().Trim() + "',");
                    objectdata.Append("'0'");
                    foreach(DataColumn dc in dr.Table.Columns)
                    {
                        if (th.IsMatch(dc.ColumnName))
                            objectdata.Append("," + dr[dc.ColumnName]);
                        if (dd.IsMatch(dc.ColumnName))
                            objectdata.Append(",'" + dr[dc.ColumnName] + "'");
                    }
                    DataRow dbr = (DataRow)ddr["b"];
                    foreach (DataColumn dc in dbr.Table.Columns)
                    {
                        if(fz.IsMatch(dc.ColumnName))
                            objectdata.Append("," + dbr[dc.ColumnName]);
                    }
                    objectdata.Append(")");

                    cmd.CommandText = objectdata.ToString();
                    cmd.Transaction = tx;
                    cmd.ExecuteNonQuery();

                    if (row_count % 500 == 0)
                    {
                        tx.Commit();
                        tx = dbconnection.BeginTransaction();
                    }

                    row_count++;
                }

                tx.Commit();
            }
        }
        public void do_zh_export(string addr, string name, string sub)
        {

            using (var dbconnection = new OleDbConnection(@"Provider=vfpoledb;Data Source=" + addr + ";Collating Sequence=machine;"))
            {
                StringBuilder objectdata = new StringBuilder();
                objectdata.Clear();
                Regex th = new Regex("^[Tt]\\d");
                Regex fz = new Regex("^[Ff][Zz]\\d");
                Regex dd = new Regex("^[Dd]\\d");

                objectdata.Append("CREATE TABLE " + name + " ( zkzh C(20), xxdm C(10), xb C(1), qxdm C(10), score N(4,1), " + sub + "_score N(4,1)," + " groups C(4)," + sub + "_groups C(4)," + " countyid C(1), llsx C(1), qxsf C(1)");
                foreach (DataColumn dc in ((DataRow)summary.Rows[0]["a"]).Table.Columns)
                {
                    if (th.IsMatch(dc.ColumnName))
                        objectdata.Append(", " + dc.ColumnName + " N(4,1)");
                    if (dd.IsMatch(dc.ColumnName))
                        objectdata.Append(", " + dc.ColumnName + " C(10)");
                }
                foreach (DataColumn dc in ((DataRow)summary.Rows[0]["b"]).Table.Columns)
                {
                    if (fz.IsMatch(dc.ColumnName))
                        objectdata.Append(", " + dc.ColumnName + " N(4,1)");
                }
                objectdata.Append(")");

                dbconnection.Open();
                OleDbCommand cmd = dbconnection.CreateCommand();
                cmd.CommandText = objectdata.ToString();
                cmd.ExecuteNonQuery();


                OleDbTransaction tx = dbconnection.BeginTransaction();

                int row_count = 0;
                foreach (DataRow ddr in summary.Rows)
                {
                    DataRow dr = (DataRow)ddr["a"];
                    objectdata.Clear();
                    objectdata.Append("INSERT INTO " + name + " VALUES (");
                    objectdata.Append("'" + dr["zkzh"].ToString().Trim() + "',");
                    objectdata.Append("'" + dr["xxdm"].ToString().Trim() + "',");
                    objectdata.Append("'" + dr["xb"].ToString().Trim() + "',");
                    objectdata.Append("'" + dr["qxdm"].ToString().Trim() + "',");
                    objectdata.Append(dr["totalmark"].ToString().Trim() + ",");
                    objectdata.Append(dr["ZH_totalmark"].ToString().Trim() + ",");
                    objectdata.Append("'" + dr["groups"].ToString().Trim() + "',");
                    objectdata.Append("'" + dr[sub + "_groups"].ToString().Trim() + "',");
                    objectdata.Append("'" + cj_data.Rows[row_count]["countyid"].ToString().Trim() + "',");
                    objectdata.Append("'" + sf_data.Rows[row_count]["llsx"].ToString().Trim() + "',");
                    objectdata.Append("'0'");
                    foreach (DataColumn dc in dr.Table.Columns)
                    {
                        if (th.IsMatch(dc.ColumnName))
                            objectdata.Append("," + dr[dc.ColumnName]);
                        if (dd.IsMatch(dc.ColumnName))
                            objectdata.Append(",'" + dr[dc.ColumnName] + "'");
                    }
                    DataRow dbr = (DataRow)ddr["b"];
                    foreach (DataColumn dc in dbr.Table.Columns)
                    {
                        if (fz.IsMatch(dc.ColumnName))
                            objectdata.Append("," + dbr[dc.ColumnName]);
                    }
                    objectdata.Append(")");

                    cmd.CommandText = objectdata.ToString();
                    cmd.Transaction = tx;
                    cmd.ExecuteNonQuery();

                    if (row_count % 500 == 0)
                    {
                        tx.Commit();
                        tx = dbconnection.BeginTransaction();
                    }

                    row_count++;
                }

                tx.Commit();
            }
        }
    }
}
