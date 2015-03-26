﻿using System;
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
        public string sf_addr;
        public string qx_code;
        public mainform _form;
        public string _exam;
        public string save_address;
        public string CurrentDirectory;
        public bool isVisible;
        public Analysis(mainform form)
        {
            _form = form;
        }
        public void gk_zt_start()
        {
            exam_type = "gk_zt";
            _exam = "gk";
            start();
        }
        public Configuration initConfig(string sub, string report, string exam)
        {
            Configuration config = new Configuration();
            config.subject = sub;
            config.report_style = report;
            config.exam = exam;
            config.isVisible = isVisible;
            config.save_address = save_address;
            config.CurrentDirectory = CurrentDirectory;
            return config;
        }
        public void gk_qx_start()
        {
            exam_type = "gk_qx";
            _exam = "gk";
            start();
        }
        public void gk_qx_process(MetaData mdata)
        {
            ArrayList QX = new ArrayList();
            ArrayList total = new ArrayList();
            PartitionQXDataProcess(mdata, total, QX, mdata.basic, mdata.group, mdata._group_num);
            form.ShowPro(70, 4);
            Partition_wordcreator create = new Partition_wordcreator(total, QX, mdata.grp, mdata.groups_group);
            create.creating_word();

            if (subject.Equals("语文") || subject.Equals("英语"))
            {
                form.ShowPro(80, 6);
                Utils.WSLG = true;
                ArrayList WSLG = new ArrayList();
                DataTable QX_data = db._basic_data.filteredtable("QX", QXTransfer(Quxian_list));
                DataTable QX_group = db._group_data.filteredtable("QX", QXTransfer(Quxian_list));

                WSLGCal(QX_data, QX_group, WSLG);

                Partition_wordcreator create2 = new Partition_wordcreator(WSLG, groups.dt, groups.groups_group);
                create2.creating_word();
                Utils.WSLG = false;

            }

        }
        void PartitionQXDataProcess(MetaData mdata, ArrayList result, ArrayList sresult, DataTable data, DataTable group, int groupnum)
        {
            Partition_statistic total = new Partition_statistic("市整体", data, mdata._fullmark, mdata.ans, group, mdata.grp, groupnum);
            total.statistic_process(false);
            if (data.Columns.Contains("XZ"))
                total.xz_postprocess(mdata.xz);
            result.Add(total.result);

            for (int i = 0; i < mdata.SF_list.Count; i++)
            {
                ArrayList sf = mdata.SF_list[i];
                string[] xx_code = new string[sf.Count - 1];
                for (int j = 1; j < sf.Count; j++)
                    xx_code[j - 1] = sf[j].ToString().Trim();
                DataTable temp = data.filteredtable("schoolcode", xx_code);
                DataTable temp_group = group.filteredtable("schoolcode", xx_code);
                Partition_statistic stat = new Partition_statistic(sf[0].ToString(), temp, mdata._fullmark, mdata.ans, temp_group, mdata.grp, groupnum);
                stat.statistic_process(false);
                if (data.Columns.Contains("XZ"))
                    stat.xz_postprocess(mdata.xz);
                result.Add(stat.result);
            }

            for (int i = 0; i < mdata.CJ_list.Count; i++)
            {
                ArrayList cj = mdata.CJ_list[i];
                string[] xx_code = new string[cj.Count - 1];
                for (int j = 1; j < cj.Count; j++)
                    xx_code[j - 1] = cj[j].ToString().Trim();
                DataTable temp = data.filteredtable("QX", xx_code);
                DataTable temp_group = group.filteredtable("QX", xx_code);
                Partition_statistic stat = new Partition_statistic(cj[0].ToString(), temp, mdata._fullmark, mdata.ans, temp_group, mdata.grp, groupnum);
                stat.statistic_process(false);
                if (data.Columns.Contains("XZ"))
                    stat.xz_postprocess(mdata.xz);
                result.Add(stat.result);
            }

            DataTable QX = data.filteredtable("QX", QXTransfer(qx_code));
            DataTable QX_group = group.filteredtable("QX", QXTransfer(qx_code));
            Partition_statistic qx_stat = new Partition_statistic("区整体", QX, mdata._fullmark, mdata.ans, QX_group, mdata.grp, groupnum);
            qx_stat.statistic_process(false);
            if (data.Columns.Contains("XZ"))
                qx_stat.xz_postprocess(mdata.xz);
            result.Add(qx_stat.result);
            if (!Utils.OnlyQZT)
            {
                PartitionDataProcess(mdata, result, mdata.QXSF_list, "schoolcode", QX, QX_group, groupnum, true);
                PartitionDataProcess(mdata, sresult, mdata.QXSF_list, "schoolcode", QX, QX_group, groupnum, false);
            }
            else
            {
                sresult.Add(qx_stat.result);
            }

        }
        public void gk_sf_start()
        {
            exam_type = "gk_sf";
            _exam = "gk";
            start();
        }
        public void gk_sf_process(MetaData mdata)
        {
            Configuration config = initConfig(mdata._sub, "两类示范校", "高考");
            ArrayList list = new ArrayList();
            _form.ShowPro("gk_sf", 1, mdata.log_name + "数据分析中...");
            PartitionDataProcess(mdata, list, mdata.CJ_list, "schoolcode", mdata.basic, mdata.group, mdata._group_num, false);
            _form.ShowPro("gk_sf", 1, mdata.log_name + "文档生成中...");
            Partition_wordcreator create = new Partition_wordcreator(list, mdata.grp, mdata.groups_group);
            create.creating_word();
        }
        public void gk_cj_start()
        {
            exam_type = "gk_cj";
            _exam = "gk";
            start();
        }
        public void gk_cj_process(MetaData mdata)
        {
            Configuration config = initConfig(mdata._sub, "城郊", "高考");

            ArrayList list = new ArrayList();
            _form.ShowPro("gk_cj", 1, mdata.log_name + "数据分析中...");
            PartitionDataProcess(mdata, list, mdata.CJ_list, "QX", mdata.basic, mdata.group, mdata._group_num, false);
            _form.ShowPro("gk_cj", 1, mdata.log_name + "文档生成中...");
            Partition_wordcreator create = new Partition_wordcreator(list, mdata.grp, mdata.groups_group);
            create.creating_word();
        }
        void PartitionDataProcess(MetaData mdata, ArrayList result, List<ArrayList> list, String filter, DataTable data, DataTable group, int groupnum, bool isQXSF)
        {
            int totalnum = 0;
            for (int i = 0; i < list.Count; i++)
                totalnum += (list[i].Count - 1);
            string[] SF_code = new string[totalnum];
            totalnum = 0;
            for (int i = 0; i < list.Count; i++)
            {
                for (int j = 1; j < list[i].Count; j++)
                {
                    SF_code[totalnum] = list[i][j].ToString().Trim();
                    totalnum++;
                }
            }

            DataTable dt = data.filteredtable(filter, SF_code);
            DataTable dt_group = group.filteredtable(filter, SF_code);
            dt.SeperateGroups(mdata._grouptype, mdata._group_num, "groups");
            dt_group.SeperateGroups(mdata._grouptype, mdata._group_num, "groups");
            if (dt.Columns.Contains("XZ"))
                XZ_group_separate(dt, mdata);
            Partition_statistic total = new Partition_statistic("分类整体", dt, mdata._fullmark, mdata.ans, dt_group, mdata.grp, groupnum);
            total.statistic_process(false);
            if (dt.Columns.Contains("XZ"))
                total.xz_postprocess(mdata.xz);
            if (isQXSF)
                result.Add(total.result);
            for (int i = 0; i < list.Count; i++)
            {
                ArrayList temp = list[i];
                string[] xx_code = new string[temp.Count - 1];
                for (int j = 1; j < temp.Count; j++)
                    xx_code[j - 1] = temp[j].ToString().Trim();
                DataTable temp_dt = dt.filteredtable(filter, xx_code);
                DataTable temp_group = dt_group.filteredtable(filter, xx_code);
                Partition_statistic stat = new Partition_statistic(temp[0].ToString(), temp_dt, mdata._fullmark, mdata.ans, temp_group, mdata.grp, groupnum);
                stat.statistic_process(false);
                if (dt.Columns.Contains("XZ"))
                    stat.xz_postprocess(mdata.xz);
                result.Add(stat.result);
            }
            if (!isQXSF)
                result.Add(total.result);
        }
        public void gk_zt_process(MetaData mdata)
        {
            Configuration config = initConfig(mdata._sub, "总体", "高考");
            
            WordData data = new WordData(mdata.groups_group);
            _form.ShowPro("gk_zt", 1, mdata.log_name + "数据分析中...");
            
            Total_statistic stat = new Total_statistic(data, mdata.basic, mdata._fullmark, mdata.ans, mdata.group, mdata.grp, mdata._group_num);
            stat._config = config;
            stat.statistic_process(false);
            if (mdata.basic.Columns.Contains("XZ"))
                stat.xz_postprocess(mdata.xz);
            _form.ShowPro("gk_zt", 1, mdata.log_name + "文档生成中...");
            WordCreator create = new WordCreator(data, config);
            create.creating_word();
        }
        public void zk_zt_start()
        {
            exam_type = "zk_zt";
            _exam = "zk";
            start();
        }
        public void zk_zt_process(MetaData mdata)
        {
            Configuration config = initConfig(mdata._sub, "总体", "中考");
            WordData result = new WordData(mdata.groups_group);
            _form.ShowPro("zk_zt", 1, mdata.log_name + "数据分析中...");
            Total_statistic stat = new Total_statistic(result, mdata.basic, mdata._fullmark, mdata.ans, mdata.group, mdata.grp, mdata._group_num);
            stat._config = config;
            stat.statistic_process(false);
            if (mdata.basic.Columns.Contains("XZ"))
                stat.xz_postprocess(mdata.xz);
            _form.ShowPro("zk_zt", 1, mdata.log_name + "文档生成中...");
            WordCreator creator = new WordCreator(result, config);
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
                    string chi_sub = row.Cells["sub"].Value.ToString().Trim();
                    string sub = Utils.language_trans(chi_sub);

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

                    if (Utils.is_gk_zh(exam, chi_sub))
                    {

                    }
                    else
                    {
                        switch (exam_type)
                        {
                            case "zk_zt":
                                zk_zt_process(mdata);
                                break;
                            case "zk_qx":
                                mdata.get_CJ_data(cj_addr);
                                mdata.get_QXSF_data(qx_addr);
                                zk_qx_process(mdata);
                                break;
                            case "gk_zt":
                                gk_zt_process(mdata);
                                break;
                            case "gk_cj":
                                mdata.get_CJ_data(cj_addr);
                                gk_cj_process(mdata);
                                break;
                            case "gk_sf":
                                mdata.get_CJ_data(sf_addr);
                                gk_sf_process(mdata);
                                break;
                            case "gk_qx":
                                mdata.get_QXSF_data(qx_addr);
                                mdata.get_CJ_data(cj_addr);
                                mdata.get_SF_data(sf_addr);
                                gk_qx_process(mdata);
                                break;
                            default:
                                break;
                        }

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
            for (int i = 0; i < mdata.QXSF_list.Count; i++)
                totalnum += (mdata.QXSF_list[i].Count - 1);
            string[] SF_code = new string[totalnum];
            totalnum = 0;
            for (int i = 0; i < mdata.QXSF_list.Count; i++)
            {
                for (int j = 1; j < mdata.QXSF_list[i].Count; j++)
                {
                    SF_code[totalnum] = mdata.QXSF_list[i][j].ToString().Trim();
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

            for (int i = 0; i < mdata.QXSF_list.Count; i++)
            {
                ArrayList temp = mdata.QXSF_list[i];
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
