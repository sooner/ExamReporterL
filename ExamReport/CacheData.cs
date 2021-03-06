﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using MySql.Data;
using MySql.Data.MySqlClient;

namespace ExamReport
{
    public class CacheData
    {
        public void save_partitiondata(string year, string exam, string sub, PartitionData pdata)
        {

            string tablename = "total_statistic";
            string basic = year + "_" + exam + "_" + sub;

            create_init_table(tablename);

            int val = MySqlHelper.ExecuteNonQuery(MySqlHelper.Conn, CommandType.Text, "insert into " + tablename + " (year,exam,sub,total_num,fullmark,max,min,avg,stDev,Dfactor,difficulty,alfa,standardErr,mean,mode,skewness,kertosis) values ('"
            + year + "','"
            + exam + "','"
            + sub + "',"
            + pdata.total_num + ","
            + pdata.fullmark + ","
            + pdata.max + ","
            + pdata.min + ","
            + pdata.avg + ","
            + pdata.stDev + ","
            + pdata.Dfactor + ","
            + pdata.difficulty + ","
            + "0,0,0,0,0,0)", null);

            if (val <= 0)
                throw new Exception("未知错误，数据库写入错误");

            DBHelper.create_mysql_table_datastyle(pdata.total_analysis, basic + "_total_analysis");
            DBHelper.create_mysql_table_datastyle(pdata.groups_analysis, basic + "_group_analysis");
            //DBHelper.create_mysql_table_datastyle(pdata.totalmark_dist, basic + "_totalmark_dist");
        }

        public void load_partitiondata(string year, string exam, string sub, PartitionData pdata)
        {

            string tablename = "total_statistic";
            string basic = year + "_" + exam + "_" + sub;

            MySqlDataReader reader = MySqlHelper.ExecuteReader(MySqlHelper.Conn, CommandType.Text, "select * from " + tablename + " where year='"
                + year + "' and exam='"
                + exam + "' and sub='"
                + sub + "'", null);

            if (!reader.Read())
                throw new Exception("缺少" + basic + "的数据");

            pdata.total_num = Convert.ToInt32(reader["total_num"]);
            pdata.fullmark = Convert.ToDecimal(reader["fullmark"]);
            pdata.max = Convert.ToDecimal(reader["max"]);
            pdata.min = Convert.ToDecimal(reader["min"]);
            pdata.avg = Convert.ToDecimal(reader["avg"]);
            pdata.stDev = Convert.ToDecimal(reader["stDev"]);
            pdata.Dfactor = Convert.ToDecimal(reader["Dfactor"]);
            pdata.difficulty = Convert.ToDecimal(reader["difficulty"]);
        
        }
        public void save_totaldata(string year, string exam, string sub, WordData total)
        {

            string tablename = "total_statistic";
            string basic = year + "_" + exam + "_" + sub;

            create_init_table(tablename);

            int val = MySqlHelper.ExecuteNonQuery(MySqlHelper.Conn, CommandType.Text, "insert into " + tablename + " (year,exam,sub,total_num,fullmark,max,min,avg,stDev,Dfactor,difficulty,alfa,standardErr,mean,mode,skewness,kertosis) values ('"
            + year + "','"
            + exam + "','"
            + sub + "',"
            + total.total_num + ","
            + total.fullmark + ","
            + total.max + ","
            + total.min + ","
            + total.avg + ","
            + total.stDev + ","
            + total.Dfactor + ","
            + total.difficulty + ","
            + total.alfa + ","
            + total.standardErr + ","
            + total.mean + ","
            + total.mode + ","
            + total.skewness + ","
            + total.kertosis + ")", null);

            if (val <= 0)
                throw new Exception("未知错误，数据库写入错误");

            DBHelper.create_mysql_table_datastyle(total.total_analysis, basic + "_total_analysis");
            DBHelper.create_mysql_table_datastyle(total.group_analysis, basic + "_group_analysis");
            DBHelper.create_mysql_table_datastyle(total.totalmark_dist, basic + "_totalmark_dist");


        }

        public void load_totaldata(string year, string exam, string sub, WordData total)
        {

            string tablename = "total_statistic";
            string basic = year + "_" + exam + "_" + sub;

            MySqlDataReader reader = MySqlHelper.ExecuteReader(MySqlHelper.Conn, CommandType.Text, "select * from "+ tablename + " where year='"
                + year + "' and exam='"
                + exam + "' and sub='"
                + sub + "'", null);

            if(!reader.Read())
                throw new Exception("缺少"+basic+"的数据");

            total.total_num = Convert.ToInt32(reader["total_num"]);
            total.fullmark = Convert.ToDecimal(reader["fullmark"]);
            total.max = Convert.ToDecimal(reader["max"]);
            total.min = Convert.ToDecimal(reader["min"]);
            total.avg = Convert.ToDecimal(reader["avg"]);
            total.stDev = Convert.ToDecimal(reader["stDev"]);
            total.Dfactor = Convert.ToDecimal(reader["Dfactor"]);
            total.difficulty = Convert.ToDecimal(reader["difficulty"]);
            total.alfa = Convert.ToDecimal(reader["alfa"]);
            total.standardErr = Convert.ToDecimal(reader["standardErr"]);
            total.mean = Convert.ToDecimal(reader["mean"]);
            total.mode = Convert.ToDecimal(reader["mode"]);
            total.skewness = Convert.ToDecimal(reader["skewness"]);
            total.kertosis = Convert.ToDecimal(reader["kertosis"]);

            total.total_analysis = MySqlHelper.GetDataSet(MySqlHelper.Conn, CommandType.Text, "select * from " + basic + "_total_analysis", null).Tables[0];
            total.group_analysis = MySqlHelper.GetDataSet(MySqlHelper.Conn, CommandType.Text, "select * from " + basic + "_group_analysis", null).Tables[0];
            total.totalmark_dist = MySqlHelper.GetDataSet(MySqlHelper.Conn, CommandType.Text, "select * from " + basic + "_totalmark_dist", null).Tables[0];
            if(sub.Equals("lz"))
                total._groups_ans = MySqlHelper.GetDataSet(MySqlHelper.Conn, CommandType.Text, "select * from " + "zh_" + year + "_gk_wl_fz", null).Tables[0];
            else if(sub.Equals("wz"))
                total._groups_ans = MySqlHelper.GetDataSet(MySqlHelper.Conn, CommandType.Text, "select * from " + "zh_" + year + "_gk_dl_fz", null).Tables[0];
            else
                total._groups_ans = MySqlHelper.GetDataSet(MySqlHelper.Conn, CommandType.Text, "select * from " + basic + "_fz", null).Tables[0];

        }

        public void save_zf_data(string year, string exam, string sub, ZF_worddata total)
        {

            string tablename = "total_statistic";
            string basic = year + "_" + exam + "_" + sub;

            create_init_table(tablename);

            int val = MySqlHelper.ExecuteNonQuery(MySqlHelper.Conn, CommandType.Text, "insert into " + tablename + " (year,exam,sub,total_num,fullmark,max,min,avg,stDev,Dfactor,difficulty,alfa,standardErr,mean,mode,skewness,kertosis) values ('"
            + year + "','"
            + exam + "','"
            + sub + "',"
            + total.total_num + ","
            + total.fullmark + ","
            + total.max + ","
            + total.min + ","
            + total.avg + ","
            + total.stDev + ","
            + total.Dfactor + ","
            + total.difficulty + ","
            + "0,"
            + "0,"
            + "0,"
            + "0,"
            + "0,"
            + "0)", null);

            if (val <= 0)
                throw new Exception("未知错误，数据库写入错误");

            DBHelper.create_mysql_table_datastyle(total.dist, basic + "_dist");

        }

        public void load_zf_data(string year, string exam, string sub, ZF_worddata total)
        {

            string tablename = "total_statistic";
            string basic = year + "_" + exam + "_" + sub;

            MySqlDataReader reader = MySqlHelper.ExecuteReader(MySqlHelper.Conn, CommandType.Text, "select * from " + tablename + " where year='"
                + year + "' and exam='"
                + exam + "' and sub='"
                + sub + "'", null);

            if (!reader.Read())
                throw new Exception("缺少" + basic + "的数据");

            total.total_num = Convert.ToInt32(reader["total_num"]);
            total.fullmark = Convert.ToDecimal(reader["fullmark"]);
            total.max = Convert.ToDecimal(reader["max"]);
            total.min = Convert.ToDecimal(reader["min"]);
            total.avg = Convert.ToDecimal(reader["avg"]);
            total.stDev = Convert.ToDecimal(reader["stDev"]);
            total.Dfactor = Convert.ToDecimal(reader["Dfactor"]);
            total.difficulty = Convert.ToDecimal(reader["difficulty"]);

            total.dist = MySqlHelper.GetDataSet(MySqlHelper.Conn, CommandType.Text, "select * from " + basic + "_dist", null).Tables[0];
        }

        public void create_init_table(string tablename)
        {
            //MySqlHelper.ExecuteNonQuery(MySqlHelper.Conn, CommandType.Text, "drop table if exists " + tablename, null);
            MySqlHelper.ExecuteNonQuery(MySqlHelper.Conn, CommandType.Text, "create table if not exists " + tablename
               + " (year varchar(255), exam varchar(255), sub varchar(255), total_num int, fullmark decimal(4,1), max decimal(4,1), min decimal(4,1), avg decimal(5,2), stDev decimal(5,2), Dfactor decimal(5,2), difficulty decimal(5,2),"
           + "alfa decimal(5,2), standardErr decimal(5,2), mean decimal(5,2), mode decimal(5,2), skewness decimal(5,2), kertosis decimal(5,2))", null);


        }




    }
}
