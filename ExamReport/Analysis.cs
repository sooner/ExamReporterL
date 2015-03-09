using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using Telerik.WinControls.UI;
using System.Data;

namespace ExamReport
{
    class Analysis
    {
        private string exam_type;

        public RadGridView _gridview;
        public string qx_addr;
        public string cj_addr;
        public string qx_code;
        public mainform _form;
        public string _exam;
        public Analysis(mainform form)
        {
            _form = form;
        }
        public void zk_zt_start()
        {
            exam_type = "zk_zt";
            _exam = "zk";
            start();
        }
        public void zk_zt_process(MetaData mdata)
        {
            WordData result = new WordData(mdata.groups_group);
            _form.ShowPro("zk_zt", 1, mdata.log_name + "数据分析中...");
            Total_statistic stat = new Total_statistic(result, mdata.basic, mdata._fullmark, mdata.ans, mdata.group, mdata.grp, mdata._group_num);
            stat.statistic_process(false);
            if (mdata.basic.Columns.Contains("XZ"))
                stat.xz_postprocess(mdata.xz);
            _form.ShowPro("zk_zt", 1, mdata.log_name + "文档生成中...");
            WordCreator creator = new WordCreator(result);
            creator.creating_word();
        }
        public void start()
        {
            _form.ShowPro(exam_type, 0, "开始处理...");
            foreach (GridViewRowInfo row in _gridview.Rows)
            {
                if (row.Cells["checkbox"].Value != null)
                {

                    string year = row.Cells["year"].Value.ToString().Trim();
                    string exam = _exam;
                    string sub = Utils.language_trans(row.Cells["sub"].Value.ToString().Trim());

                    string log = year + "年" + Utils.language_trans(exam) + row.Cells["sub"].Value.ToString().Trim();
                    _form.ShowPro(exam_type, 1, log + "数据读取...");
                    MetaData mdata = new MetaData(year, exam, sub);
                    //try
                    //{
                    mdata.get_meta_data();
                    //}
                    //catch (Exception e)
                    //{
                    //    _form.ErrorM(e.Message);
                    //}
                    mdata.log_name = log;
                    mdata.get_basic_data();
                    mdata.get_group_data();
                    mdata.get_ans();
                    mdata.get_fz();

                    switch (exam_type)
                    {
                        case "zk_zt":
                            zk_zt_process(mdata);
                            break;
                        case "zk_qx":
                            mdata.get_CJ_data(cj_addr);
                            mdata.get_QX_data(qx_addr);
                            zk_qx_process(mdata);
                            break;
                        default:
                            break;
                    }




                }
            }
            _form.ShowPro(exam_type, 2, "完成！");
        }
        public void zk_qx_start()
        {
            _exam = "zk";
            exam_type = "zk_qx";
            start();
        }



        public void zk_qx_process(MetaData mdata)
        {
            ArrayList sdata = new ArrayList();
            ArrayList totaldata = new ArrayList();
            _form.ShowPro(exam_type, 1, mdata.log_name + "数据分析中...");
            Partition_statistic total = new Partition_statistic("市整体", mdata.basic, mdata._fullmark, mdata.ans, mdata.group, mdata.grp, mdata._group_num);
            total.statistic_process(false);
            if (mdata.basic.Columns.Contains("XZ"))
                total.xz_postprocess(mdata.xz);
            totaldata.Add(total.result);

            for (int mark = 0; mark < mdata.CJ_list.Count; mark++)
            {
                string[] CQ_code = new string[mdata.CJ_list[mark].Count - 1];

                for (int i = 1; i < mdata.CJ_list[mark].Count; i++)
                {
                    CQ_code[i - 1] = mdata.CJ_list[mark][i].ToString().Trim();
                }
                DataTable CQ_data = mdata.basic.filteredtable("QX", CQ_code);
                DataTable CQ_groups_data = mdata.group.filteredtable("QX", CQ_code);

                Partition_statistic CQ = new Partition_statistic(mdata.CJ_list[mark][0].ToString().Trim(), CQ_data, mdata._fullmark, mdata.ans, CQ_groups_data, mdata.group, mdata._group_num);
                CQ.statistic_process(false);
                if (mdata.basic.Columns.Contains("XZ"))
                    CQ.xz_postprocess(mdata.xz);
                totaldata.Add(CQ.result);
            }

            DataTable QX_total_data = mdata.basic.filteredtable("QX", QXTransfer(qx_code));
            DataTable QX_groups_data = mdata.group.filteredtable("QX", QXTransfer(qx_code));

            Partition_statistic QX_total = new Partition_statistic("区整体", QX_total_data, mdata._fullmark, mdata.ans, QX_groups_data, mdata.group, mdata._group_num);
            QX_total.statistic_process(false);
            if (mdata.basic.Columns.Contains("XZ"))
                QX_total.xz_postprocess(mdata.xz);
            totaldata.Add(QX_total.result);

            CalculateClassTotal(QX_total_data, QX_groups_data, totaldata, sdata, mdata);
            _form.ShowPro(exam_type, 1, mdata.log_name + "文档生成中...");
            Partition_wordcreator create = new Partition_wordcreator(totaldata, sdata, mdata.group, mdata.groups_group);
            create.creating_word();
        }

        string[] QXTransfer(string QX)
        {

            if (QX.Contains(','))
            {
                string[] district = QX.Split(',');
                return district;
            }
            else
            {
                string[] other = { QX };
                return other;
            }
        }

        void CalculateClassTotal(DataTable total, DataTable groups_data, ArrayList totaldata, ArrayList sdata, MetaData mdata)
        {
            int totalnum = 0;
            for (int i = 0; i < mdata.QX_list.Count; i++)
                totalnum += (mdata.QX_list[i].Count - 1);
            string[] SF_code = new string[totalnum];
            totalnum = 0;
            for (int i = 0; i < mdata.QX_list.Count; i++)
            {
                for (int j = 1; j < mdata.QX_list[i].Count; j++)
                {
                    SF_code[totalnum] = mdata.QX_list[i][j].ToString().Trim();
                    totalnum++;
                }
            }
            DataTable ClassTotal_data = total.filteredtable("schoolcode", SF_code);
            DataTable ClassGroupTotal_data = groups_data.filteredtable("schoolcode", SF_code);

            int groupnum = ClassTotal_data.SeperateGroups(mdata._grouptype, mdata._group_num, "groups");
            ClassGroupTotal_data.SeperateGroups(mdata._grouptype, mdata._group_num, "groups");
            if (ClassTotal_data.Columns.Contains("XZ"))
                XZ_group_separate(ClassTotal_data, mdata);
            Partition_statistic ClassTotal = new Partition_statistic("分类整体", ClassTotal_data, mdata._fullmark, mdata.ans, ClassGroupTotal_data, mdata.group, mdata._group_num);
            ClassTotal.statistic_process(false);
            if (ClassTotal_data.Columns.Contains("XZ"))
                ClassTotal.xz_postprocess(mdata.xz);
            totaldata.Add(ClassTotal.result);

            for (int i = 0; i < mdata.QX_list.Count; i++)
            {
                ArrayList temp = mdata.QX_list[i];
                string[] xx_code = new string[temp.Count - 1];
                for (int j = 1; j < temp.Count; j++)
                    xx_code[j - 1] = temp[j].ToString().Trim();
                DataTable xx_data = ClassTotal_data.filteredtable("schoolcode", xx_code);
                DataTable xx_group_data = ClassGroupTotal_data.filteredtable("schoolcode", xx_code);

                Partition_statistic XXTotal = new Partition_statistic(temp[0].ToString().Trim(), xx_data, mdata._fullmark, mdata.ans, xx_group_data, mdata.group, mdata._group_num);
                XXTotal.statistic_process(false);
                if (ClassTotal_data.Columns.Contains("XZ"))
                    XXTotal.xz_postprocess(mdata.xz);
                totaldata.Add(XXTotal.result);
                sdata.Add(XXTotal.result);

            }
            sdata.Add(ClassTotal.result);
        }

        void XZ_group_separate(DataTable temp_dt, MetaData mdata)
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
                inter_table.SeperateGroups(mdata._grouptype, mdata._group_num, "xz_groups");
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
