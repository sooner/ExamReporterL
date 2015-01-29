using System;
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

        public string exam;
        public string database_str;
        public string ans_str;
        public string group_str;

        public string sub;

        public decimal fullmark;
        public ZK_database.GroupType grouptype;
        public decimal divider;

        public void start_process()
        {
            if (!sub.Equals("总分"))
            {
                ans = new excel_process(ans_str);
                ans.run(true);
                wizard.ShowPro(5, 1);
                groups = new excel_process(group_str);
                groups.run(false);
                wizard.ShowPro(10, 2);
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
