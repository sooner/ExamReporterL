﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace ExamReport
{
    public class LoadDatabase
    {
        public MyWizard wizard;
        excel_process ans;
        excel_process groups;
        excel_process wenli;

        public string year;
        public string exam;
        public string database_str;
        public string ans_str;
        public string group_str;
        public string wenli_str;
        
        public string sub;

        public decimal fullmark;
        public decimal sub_fullmark;
        public ZK_database.GroupType grouptype;
        public decimal divider;

        public DataTable basic_data;
        public DataTable group_data;
        public DataTable zh_group_data;
        public DataTable zh_single_data;
        

        public void start_process()
        {
            if (!sub.Equals("总分"))
            {
                ans = new excel_process(ans_str);
                ans.run(true);
                wizard.ShowPro(5, 1);
                groups = new excel_process(group_str);
                groups.run(false);
                if (sub.Contains("理综") || sub.Contains("文综"))
                {
                    wenli = new excel_process(wenli_str);
                    wenli.run(false);
                }

                wizard.ShowPro(10, 1);
            }

            MetaData md = new MetaData(year,
                Utils.language_trans(exam),
                Utils.language_trans(sub));
            if (sub.Contains("理综") || sub.Contains("文综"))
            {
                md._sub_fullmark = sub_fullmark;
            }
            md._fullmark = fullmark;
            md._grouptype = grouptype;
            md._group_num = Convert.ToInt32(divider);

            try
            {
                md.insert_data();
            }
            catch (DuplicateNameException ex)
            {
                wizard.ErrorM("该数据已存储，请先删除后再添加");
            }
            catch (Exception ex)
            {
                wizard.ErrorM(ex.Message);

            }

            switch (exam)
            {
                case "中考":
                    zk_database_process();
                    break;
                case "会考":
                    hk_database_process();
                    break;
                case "高考":
                    gk_database_process();
                    break;
                default:
                    break;

            }
            
        }

        public void zk_database_process()
        {
            

            ZK_database db = new ZK_database(ans.dt, groups.dt, grouptype, divider);
            db.DBF_data_process(database_str, wizard);

            if (db._basic_data.Columns.Contains("XZ"))
            {
                XZ_group_separate(db._basic_data);
            }
            basic_data = db._basic_data;
            group_data = db._group_data;

            
            DBHelper.create_mysql_table(basic_data, Utils.get_basic_tablename(year, Utils.language_trans(exam), Utils.language_trans(sub)));
            DBHelper.create_mysql_table(group_data, Utils.get_group_tablename(year, Utils.language_trans(exam), Utils.language_trans(sub)));
            DBHelper.create_ans_table(Utils.get_tablename(year, Utils.language_trans(exam), Utils.language_trans(sub)), db.newStandard, ans.xz_th);
            DBHelper.create_fz_table(Utils.get_tablename(year, Utils.language_trans(exam), Utils.language_trans(sub)), groups.dt, groups.groups_group);
            wizard.ShowPro(100, 3);

        }

        public void hk_database_process()
        {
            HK_database hk = new HK_database(ans.dt, groups.dt, grouptype, divider);
            hk.DBF_data_process(database_str);
            ans.dt = hk.newStandard;
        }

        public void gk_database_process()
        {

            if (sub.Equals("总分"))
            {
                GK_database db = new GK_database();
                db.ZF_data_process(database_str);
            }
            else if (sub.Contains("理综") ||
                    sub.Contains("文综"))
            {
                int ch_num = 0;
                GK_database db = new GK_database(ans.dt, groups.dt, grouptype, divider);
                db.DBF_data_process(database_str);

                ch_num = db.ZH_postprocess(wenli.dt, sub.Substring(3));

                basic_data = db._basic_data;
                group_data = db._group_data;
                zh_single_data = db.zh_single_data;
                zh_group_data = db.zh_group_data;

                DBHelper.create_mysql_table(basic_data, "zh_" + Utils.get_basic_tablename(year, Utils.language_trans(exam), Utils.language_trans(sub)));
                DBHelper.create_mysql_table(zh_group_data, "zh_" + Utils.get_group_tablename(year, Utils.language_trans(exam), Utils.language_trans(sub)));
                DBHelper.create_ans_table("zh_" + Utils.get_tablename(year, Utils.language_trans(exam), Utils.language_trans(sub)), db.newStandard, ans.xz_th);
                DBHelper.create_fz_table("zh_" + Utils.get_tablename(year, Utils.language_trans(exam), Utils.language_trans(sub)), wenli.dt, wenli.groups_group);

                DBHelper.create_mysql_table(zh_single_data, Utils.get_basic_tablename(year, Utils.language_trans(exam), Utils.language_trans(sub)));
                DBHelper.create_mysql_table(group_data, Utils.get_group_tablename(year, Utils.language_trans(exam), Utils.language_trans(sub)));
                DBHelper.create_ans_table(Utils.get_tablename(year, Utils.language_trans(exam), Utils.language_trans(sub)), db.ZH_standard_ans, ans.xz_th);
                DBHelper.create_fz_table(Utils.get_tablename(year, Utils.language_trans(exam), Utils.language_trans(sub)), groups.dt, groups.groups_group);
                wizard.ShowPro(100, 3);
            }
            else
            {
                GK_database db = new GK_database(ans.dt, groups.dt, grouptype, divider);
                db.DBF_data_process(database_str);

                if (db._basic_data.Columns.Contains("XZ"))
                {
                    XZ_group_separate(db._basic_data);
                }

                basic_data = db._basic_data;
                group_data = db._group_data;

                DBHelper.create_mysql_table(basic_data, Utils.get_basic_tablename(year, Utils.language_trans(exam), Utils.language_trans(sub)));
                DBHelper.create_mysql_table(group_data, Utils.get_group_tablename(year, Utils.language_trans(exam), Utils.language_trans(sub)));
                DBHelper.create_ans_table(Utils.get_tablename(year, Utils.language_trans(exam), Utils.language_trans(sub)), db.newStandard, ans.xz_th);
                DBHelper.create_fz_table(Utils.get_tablename(year, Utils.language_trans(exam), Utils.language_trans(sub)), groups.dt, groups.groups_group);
                wizard.ShowPro(100, 3);
            }
        }


        void XZ_group_separate(DataTable temp_dt)
        {
            if (!temp_dt.Columns.Contains("xz_groups"))
                temp_dt.Columns.Add("xz_groups", typeof(string));
            var xz_tuple = from row in temp_dt.AsEnumerable()
                           group row by row.Field<string>("XZ") into grp
                           select new
                           {
                               name = grp.Key
                           };
            foreach (var item in xz_tuple)
            {
                DataView dv = temp_dt.equalfilter("XZ", item.name).DefaultView;
                DataTable inter_table = dv.ToTable();
                inter_table.SeperateGroups(grouptype, divider, "xz_groups");
                var temp = from row in temp_dt.AsEnumerable()
                           join row2 in inter_table.AsEnumerable() on row.Field<string>("studentid") equals row2.Field<string>("studentid")
                           where row.Field<string>("XZ") == item.name
                           select new
                           {
                               row1 = row,
                               groups = row2.Field<string>("xz_groups")
                           };
                foreach (var inner_item in temp)
                {
                    inner_item.row1.SetField<string>("xz_groups", inner_item.groups);
                }
            }
        }
    }
}
