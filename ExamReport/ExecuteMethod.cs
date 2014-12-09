﻿using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace ExamReport
{
    public class ExecuteMethod
    {
        public Form1 form;
        excel_process ans;
        excel_process groups;
        excel_process wenli;
        List<ArrayList> QXSF_list;
        List<ArrayList> SF_list;
        List<ArrayList> CJ_list;

        /// <summary>
        /// 科目：语文、数学、英语等
        /// </summary>
        private string subject;

        public string Subject
        {
            get { return subject; }
            set { subject = value; }
        }
        /// <summary>
        /// 考试类型：中考、会考、高考
        /// </summary>
        private string style;

        public string Style
        {
            get { return style; }
            set { style = value; }
        }

        /// <summary>
        /// 报告类型
        /// </summary>
        string report_style;

        public string Report_style
        {
            get { return report_style; }
            set { report_style = value; }
        }
        /// <summary>
        /// 标准答案地址
        /// </summary>
        string ans_address;

        public string Ans_address
        {
            get { return ans_address; }
            set { ans_address = value; }
        }
        /// <summary>
        /// 分组信息地址
        /// </summary>
        string groups_address;

        public string Groups_address
        {
            get { return groups_address; }
            set { groups_address = value; }
        }

        /// <summary>
        /// 数据库文件地址
        /// </summary>
        string database_address;

        public string Database_address
        {
            get { return database_address; }
            set { database_address = value; }
        }

        /// <summary>
        /// 文理综分类
        /// </summary>
        string wenli_catagory;
        public string Wenli_catagory
        {
            get { return wenli_catagory; }
            set { wenli_catagory = value; }
        }
        /// <summary>
        /// 区县学校分类
        /// </summary>
        string quxian_catagory;
        public string Quxian_catagory
        {
            get { return quxian_catagory; }
            set { quxian_catagory = value; }
        }
        /// <summary>
        /// 示范校分类
        /// </summary>
        string shifan_catagory;
        public string Shifan_catagory
        {
            get { return shifan_catagory; }
            set { shifan_catagory = value; }
        }
        string cj_catagory;

        public string Cj_catagory
        {
            get { return cj_catagory; }
            set { cj_catagory = value; }
        }

        public string Quxian_list;
        public Dictionary<string, string> School_code;

        public ZK_database.GroupType grouptype;
        public decimal divider;

        public HK_hierarchy hk_hierarchy;
        public class HK_hierarchy
        {
            public decimal excellent_low;
            public decimal excellent_high;
            public decimal well_low;
            public decimal well_high;
            public decimal pass_low;
            public decimal pass_high;
            public decimal fail_low;
            public decimal fail_high;

        }

        public decimal fullmark;

        

        public void pre_process()
        {
            //try
            //{
                Utils.exam = style;
                Utils.subject = subject;

                if (!subject.Equals("总分"))
                {
                    ans = new excel_process(ans_address);
                    ans.run(true);
                    form.ShowPro(5, 1);
                    groups = new excel_process(groups_address);
                    groups.run(false);
                    form.ShowPro(10, 2);
                }
                if (style.Equals("会考"))
                {
                    HK_process();
                    return;
                }
                Utils.report_style = report_style;
                if (report_style.Equals("区县") || report_style.Equals("学校"))
                {
                    excel_process QXSF = new excel_process(quxian_catagory);
                    QXSF_list = QXSF.getData();

                    excel_process CJ = new excel_process(cj_catagory);
                    CJ_list = CJ.getData();

                    if (style.Equals("高考"))
                    {
                        excel_process SF = new excel_process(shifan_catagory);
                        SF_list = SF.getData();
                    }
                }
                if (report_style.Equals("两类示范校"))
                {
                    excel_process SF = new excel_process(shifan_catagory);
                    SF_list = SF.getData();
                }
                if (report_style.Equals("城郊"))
                {
                    excel_process CJ = new excel_process(cj_catagory);
                    CJ_list = CJ.getData();
                }
                if (subject.Contains("理综") ||
                    subject.Contains("文综"))
                {
                    wenli = new excel_process(wenli_catagory);
                    wenli.run(false);
                }
                if (style.Equals("中考"))
                    ZK_process();
                if (style.Equals("高考"))
                    GK_process();
            //}
            //catch (System.Threading.ThreadAbortException e)
            //{
            //}
            //catch (Exception e)
            //{
            //    form.ErrorM(e.Message.ToString());
            //}
            
        }
        public void ZK_process()
        {
            if(report_style.Equals("总体"))
            {
                ZK_database db = new ZK_database(ans.dt, groups.dt, grouptype, divider);
                db.DBF_data_process(database_address, form);
                WordData result = new WordData(groups.groups_group);
                if (db._basic_data.Columns.Contains("XZ"))
                {
                    XZ_group_separate(db._basic_data);
                }
                form.ShowPro(40, 3);
                Total_statistic stat = new Total_statistic(result, db._basic_data, fullmark, ans.dt, db._group_data, groups.dt, db._group_num);
                stat.statistic_process(false);
                if (db._basic_data.Columns.Contains("XZ"))
                    stat.xz_postprocess(ans.xz_th);
                form.ShowPro(70, 4);
                WordCreator creator = new WordCreator(result);
                creator.creating_word();
            }
            else if (report_style.Equals("区县"))
            {
                ArrayList sdata = new ArrayList();
                ArrayList totaldata = new ArrayList();

                ZK_database db = new ZK_database(ans.dt, groups.dt, grouptype, divider);
                
                db.DBF_data_process(database_address, form);
                form.ShowPro(40, 3);
                if (db._basic_data.Columns.Contains("XZ"))
                {
                    XZ_group_separate(db._basic_data);
                }
                Partition_statistic total = new Partition_statistic("市整体", db._basic_data, fullmark, ans.dt, db._group_data, groups.dt, db._group_num);
                total.statistic_process(false);
                if (db._basic_data.Columns.Contains("XZ"))
                    total.xz_postprocess(ans.xz_th);
                totaldata.Add(total.result);

                for (int mark = 0; mark < CJ_list.Count; mark++)
                {
                    string[] CQ_code = new string[CJ_list[mark].Count - 1];

                    for (int i = 1; i < CJ_list[mark].Count; i++)
                    {
                        CQ_code[i - 1] = CJ_list[mark][i].ToString().Trim();
                    }
                    DataTable CQ_data = db._basic_data.filteredtable("QX", CQ_code);
                    DataTable CQ_groups_data = db._group_data.filteredtable("QX", CQ_code);

                    Partition_statistic CQ = new Partition_statistic(CJ_list[mark][0].ToString().Trim(), CQ_data, fullmark, ans.dt, CQ_groups_data, groups.dt, db._group_num);
                    CQ.statistic_process(false);
                    if (db._basic_data.Columns.Contains("XZ"))
                        CQ.xz_postprocess(ans.xz_th);
                    totaldata.Add(CQ.result);
                }

                DataTable QX_total_data = db._basic_data.filteredtable("QX", QXTransfer(Quxian_list));
                DataTable QX_groups_data = db._group_data.filteredtable("QX", QXTransfer(Quxian_list));

                Partition_statistic QX_total = new Partition_statistic("区整体", QX_total_data, fullmark, ans.dt, QX_groups_data, groups.dt, db._group_num);
                QX_total.statistic_process(false);
                if (db._basic_data.Columns.Contains("XZ"))
                    QX_total.xz_postprocess(ans.xz_th);
                totaldata.Add(QX_total.result);

                CalculateClassTotal(QX_total_data, QX_groups_data, totaldata, sdata);
                form.ShowPro(70, 4);
                Partition_wordcreator create = new Partition_wordcreator(totaldata, sdata, groups.dt, groups.groups_group);
                create.creating_word();
            }
        }
        public void HK_process()
        {
            HK_database hk = new HK_database(ans.dt, groups.dt, grouptype, divider);
            hk.DBF_data_process(database_address);
            ans.dt = hk.newStandard;
            form.ShowPro(40, 3);
            HK_worddata result = new HK_worddata(groups.groups_group);
            Total_statistic stat = new Total_statistic(result, hk._basic_data, fullmark, ans.dt, hk._group_data, groups.dt, hk._group_num);
            stat.statistic_process(false);
            stat.HK_postprocess(hk_hierarchy);
            form.ShowPro(70, 4);
            WordCreator create = new WordCreator(result);
            create.creating_HK_word();


        }
        public void GK_process()
        {

            if (subject.Equals("总分"))
            {
                if (report_style.Equals("两类示范校"))
                {
                    List<ZF_statistic> result = new List<ZF_statistic>();
                    GK_database db = new GK_database();
                    db.ZF_data_process(database_address);
                    form.ShowPro(40, 3);
                    for (int i = 0; i < SF_list.Count; i++)
                    {
                        string[] SF_code = new string[SF_list[i].Count - 1];
                        for (int j = 1; j < SF_list[i].Count; j++)
                            SF_code[j - 1] = SF_list[i][j].ToString().Trim();
                        DataTable temp = db._basic_data.filteredtable("schoolcode", SF_code);
                        ZF_statistic stat = new ZF_statistic(temp, fullmark, SF_list[i][0].ToString().Trim());
                        stat.partition_process();
                        result.Add(stat);
                    }
                    form.ShowPro(70, 4);
                    ZF_wordcreator create = new ZF_wordcreator();
                    create.partition_wordcreate(result, "两类示范校");
                    
                    
                }
                if (report_style.Equals("城郊"))
                {
                    List<ZF_statistic> result = new List<ZF_statistic>();
                    GK_database db = new GK_database();
                    db.ZF_data_process(database_address);
                    form.ShowPro(40, 3);
                    for (int i = 0; i < CJ_list.Count; i++)
                    {
                        string[] cj_code = new string[CJ_list[i].Count - 1];
                        for (int j = 1; j < CJ_list[i].Count; j++)
                            cj_code[j - 1] = CJ_list[i][j].ToString().Trim();
                        DataTable temp = db._basic_data.filteredtable("qxdm", cj_code);
                        ZF_statistic stat = new ZF_statistic(temp, fullmark, CJ_list[i][0].ToString().Trim());
                        stat.partition_process();
                        result.Add(stat);
                    }
                    form.ShowPro(70, 4);
                    ZF_wordcreator create = new ZF_wordcreator();
                    create.partition_wordcreate(result, "城郊");
                }
                if (report_style.Equals("区县"))
                {
                    List<ZF_statistic> result = new List<ZF_statistic>();
                    GK_database db = new GK_database();
                    db.ZF_data_process(database_address);
                    form.ShowPro(40, 3);
                    ZF_statistic total = new ZF_statistic(db._basic_data, fullmark, "市整体");
                    total.partition_process();
                    result.Add(total);
                    for (int i = 0; i < SF_list.Count; i++)
                    {
                        string[] SF_code = new string[SF_list[i].Count - 1];
                        for (int j = 1; j < SF_list[i].Count; j++)
                            SF_code[j - 1] = SF_list[i][j].ToString().Trim();
                        DataTable temp = db._basic_data.filteredtable("schoolcode", SF_code);
                        ZF_statistic stat = new ZF_statistic(temp, fullmark, SF_list[i][0].ToString().Trim());
                        stat.partition_process();
                        result.Add(stat);
                    }
                    for (int i = 0; i < CJ_list.Count; i++)
                    {
                        string[] cj_code = new string[CJ_list[i].Count - 1];
                        for (int j = 1; j < CJ_list[i].Count; j++)
                            cj_code[j - 1] = CJ_list[i][j].ToString().Trim();
                        DataTable temp = db._basic_data.filteredtable("qxdm", cj_code);
                        ZF_statistic stat = new ZF_statistic(temp, fullmark, CJ_list[i][0].ToString().Trim());
                        stat.partition_process();
                        result.Add(stat);
                    }
                    DataTable bq_data = db._basic_data.filteredtable("qxdm", QXTransfer(Quxian_list));
                    ZF_statistic bq = new ZF_statistic(bq_data, fullmark, "本区");
                    bq.partition_process();
                    result.Add(bq);
                    CalculateGKZF(bq_data, result);
                    form.ShowPro(70, 4);
                    ZF_wordcreator create = new ZF_wordcreator();
                    create.partition_wordcreate(result, "区县");

                }
                if (report_style.Equals("总体"))
                {
                    GK_database db = new GK_database();
                    db.ZF_data_process(database_address);
                    form.ShowPro(40, 3);
                    ZF_statistic stat = new ZF_statistic(db._basic_data, fullmark, "总体");
                    stat.partition_process();
                    form.ShowPro(70, 4);
                    ZF_wordcreator create = new ZF_wordcreator();
                    create.total_create(stat);
                }

            }
            else if (subject.Contains("理综") ||
                    subject.Contains("文综"))
            {
                string sub = subject.Substring(3);
                int ch_num = 0;
                GK_database db = new GK_database(ans.dt, groups.dt, grouptype, divider);
                db.DBF_data_process(database_address);
                ans.dt = db.newStandard;
                ch_num = db.ZH_postprocess(wenli.dt, subject.Substring(3));
                DataTable wenli_standard = db.ZH_standard_ans;
                form.ShowPro(40, 3);
                if (ch_num == -1)
                    return;
                decimal sub_fullmark = 0;
                if (sub.Equals("生物") || sub.Equals("政治"))
                    sub_fullmark = Utils.shengwu_zhengzhi;
                else if (sub.Equals("物理") || sub.Equals("历史"))
                    sub_fullmark = Utils.wuli_lishi;
                else if (sub.Equals("化学") || sub.Equals("地理"))
                    sub_fullmark = Utils.huaxue_dili;
                if (report_style.Equals("总体"))
                {

                    WordData total = new WordData(wenli.groups_group);
                    Total_statistic total_stat = new Total_statistic(total, db._basic_data, fullmark, ans.dt, db.zh_group_data, wenli.dt, db._group_num);
                    total_stat.statistic_process(true);

                    WordData single = new WordData(groups.groups_group);

                    Total_statistic single_stat = new Total_statistic(single, db.zh_single_data, sub_fullmark, wenli_standard, db._group_data, groups.dt, db._group_num);
                    single_stat.statistic_process(false);
                    form.ShowPro(70, 4);
                    WordCreator create = new WordCreator(single, total);
                    create.creating_word();
                }
                else if (report_style.Equals("两类示范校"))
                {
                    ArrayList sdata = new ArrayList();
                    ArrayList ZH_data = new ArrayList();

                    string[] total_code = CalculateTotal(SF_list);
                    DataTable total = db._basic_data.filteredtable("schoolcode", total_code);
                    DataTable total_group = db.zh_group_data.filteredtable("schoolcode", total_code);

                    int groupnum = total.SeperateGroups(grouptype, divider, "groups");
                    total_group.SeperateGroups(grouptype, divider, "groups");

                    DataTable single_total = db.zh_single_data.filteredtable("schoolcode", total_code);
                    DataTable single_total_group = db._group_data.filteredtable("schoolcode", total_code);

                    single_total.SeperateGroups(grouptype, divider, "groups");
                    single_total_group.SeperateGroups(grouptype, divider, "groups");
                    for (int i = 0; i < SF_list.Count; i++)
                    {
                        string[] SF_code = new string[SF_list[i].Count - 1];
                        for (int j = 1; j < SF_list[i].Count; j++)
                            SF_code[j - 1] = SF_list[i][j].ToString().Trim();
                        DataTable temp = total.filteredtable("schoolcode", SF_code);
                        DataTable temp_group = total_group.filteredtable("schoolcode", SF_code);
                        Partition_statistic stat = new Partition_statistic(SF_list[i][0].ToString().Trim(), temp, fullmark, ans.dt, temp_group, wenli.dt, groupnum);
                        stat.statistic_process(true);
                        ZH_data.Add(stat.result);

                        DataTable single = single_total.filteredtable("schoolcode", SF_code);
                        DataTable single_group = single_total_group.filteredtable("schoolcode", SF_code);
                        Partition_statistic single_stat = new Partition_statistic(SF_list[i][0].ToString().Trim(), single, sub_fullmark, wenli_standard, single_group, groups.dt, groupnum);
                        single_stat.statistic_process(false);
                        sdata.Add(single_stat.result);
                    }

                    Partition_statistic total_stat = new Partition_statistic("分类整体", total, fullmark, ans.dt, total_group, wenli.dt, groupnum);
                    total_stat.statistic_process(true);
                    ZH_data.Add(total_stat.result);
                    Partition_statistic single_total_stat = new Partition_statistic("分类整体", single_total, sub_fullmark, wenli_standard, single_total_group, groups.dt, groupnum);
                    single_total_stat.statistic_process(false);
                    sdata.Add(single_total_stat.result);
                    form.ShowPro(70, 4);
                    Partition_wordcreator create = new Partition_wordcreator(sdata, groups.dt, groups.groups_group);
                    create.creating_ZH_word(ZH_data, wenli.dt, wenli.groups_group);
                }
                else if (report_style.Equals("城郊"))
                {
                    ArrayList sdata = new ArrayList();
                    ArrayList ZH_data = new ArrayList();
                    string[] total_code = CalculateTotal(CJ_list);

                    DataTable total = db._basic_data.filteredtable("QX", total_code);
                    DataTable total_group = db.zh_group_data.filteredtable("QX", total_code);

                    int groupnum = total.SeperateGroups(grouptype, divider, "groups");
                    total_group.SeperateGroups(grouptype, divider, "groups");

                    DataTable single_total = db.zh_single_data.filteredtable("QX", total_code);
                    DataTable single_total_group = db._group_data.filteredtable("QX", total_code);

                    single_total.SeperateGroups(grouptype, divider, "groups");
                    single_total_group.SeperateGroups(grouptype, divider, "groups");
                    for (int i = 0; i < CJ_list.Count; i++)
                    {
                        string[] SF_code = new string[CJ_list[i].Count - 1];
                        for (int j = 1; j < CJ_list[i].Count; j++)
                            SF_code[j - 1] = CJ_list[i][j].ToString().Trim();
                        DataTable temp = total.filteredtable("QX", SF_code);
                        DataTable temp_group = total_group.filteredtable("QX", SF_code);
                        Partition_statistic stat = new Partition_statistic(CJ_list[i][0].ToString().Trim(), temp, fullmark, ans.dt, temp_group, wenli.dt, groupnum);
                        stat.statistic_process(true);
                        ZH_data.Add(stat.result);

                        DataTable single = single_total.filteredtable("QX", SF_code);
                        DataTable single_group = single_total_group.filteredtable("QX", SF_code);
                        Partition_statistic single_stat = new Partition_statistic(CJ_list[i][0].ToString().Trim(), single, sub_fullmark, wenli_standard, single_group, groups.dt, groupnum);
                        single_stat.statistic_process(false);
                        sdata.Add(single_stat.result);
                    }

                    Partition_statistic total_stat = new Partition_statistic("分类整体", total, fullmark, ans.dt, total_group, wenli.dt, groupnum);
                    total_stat.statistic_process(true);
                    ZH_data.Add(total_stat.result);
                    Partition_statistic single_total_stat = new Partition_statistic("分类整体", single_total, sub_fullmark, wenli_standard, single_total_group, groups.dt, groupnum);
                    single_total_stat.statistic_process(false);
                    sdata.Add(single_total_stat.result);
                    form.ShowPro(70, 4);
                    Partition_wordcreator create = new Partition_wordcreator(sdata, groups.dt, groups.groups_group);
                    create.creating_ZH_word(ZH_data, wenli.dt, wenli.groups_group);

                }
                else if (report_style.Equals("区县"))
                {
                    ArrayList total = new ArrayList();
                    ArrayList QX = new ArrayList();
                    ArrayList ZH_total = new ArrayList();
                    ArrayList ZH_QX = new ArrayList();

                    CalculatePartition(ZH_total, "市整体", db._basic_data, db.zh_group_data, fullmark, wenli.dt, db._group_num, true, ans.dt);
                    //decimal ZH_fullmark = (decimal)((PartitionData)ZH_total[0]).groups_analysis.Rows.Find(sub)["fullmark"];
                    CalculatePartition(total, "市整体", db.zh_single_data, db._group_data, sub_fullmark, groups.dt, db._group_num, false, wenli_standard);
                    for (int i = 0; i < SF_list.Count; i++)
                    {
                        string[] SF_code = new string[SF_list[i].Count - 1];
                        for (int j = 1; j < SF_list[i].Count; j++)
                            SF_code[j - 1] = SF_list[i][j].ToString().Trim();
                        DataTable temp = db._basic_data.filteredtable("schoolcode", SF_code);
                        DataTable temp_group = db.zh_group_data.filteredtable("schoolcode", SF_code);

                        DataTable single = db.zh_single_data.filteredtable("schoolcode", SF_code);
                        DataTable single_table = db._group_data.filteredtable("schoolcode", SF_code);
                        CalculatePartition(ZH_total, SF_list[i][0].ToString(), temp, temp_group, fullmark, wenli.dt, db._group_num, true, ans.dt);
                        CalculatePartition(total, SF_list[i][0].ToString(), single, single_table, sub_fullmark, groups.dt, db._group_num, false, wenli_standard);
                    }
                    for (int i = 0; i < CJ_list.Count; i++)
                    {
                        string[] SF_code = new string[CJ_list[i].Count - 1];
                        for (int j = 1; j < CJ_list[i].Count; j++)
                            SF_code[j - 1] = CJ_list[i][j].ToString().Trim();
                        DataTable temp = db._basic_data.filteredtable("QX", SF_code);
                        DataTable temp_group = db.zh_group_data.filteredtable("QX", SF_code);

                        DataTable single = db.zh_single_data.filteredtable("QX", SF_code);
                        DataTable single_table = db._group_data.filteredtable("QX", SF_code);
                        CalculatePartition(ZH_total, CJ_list[i][0].ToString(), temp, temp_group, fullmark, wenli.dt, db._group_num, true, ans.dt);
                        CalculatePartition(total, CJ_list[i][0].ToString(), single, single_table, sub_fullmark, groups.dt, db._group_num, false, wenli_standard);
                    }
                    DataTable QX_ZH_data = db._basic_data.filteredtable("QX", QXTransfer(Quxian_list));
                    DataTable QX_ZH_group = db.zh_group_data.filteredtable("QX", QXTransfer(Quxian_list));

                    DataTable QX_data = db.zh_single_data.filteredtable("QX", QXTransfer(Quxian_list));
                    DataTable QX_group = db._group_data.filteredtable("QX", QXTransfer(Quxian_list));

                    CalculatePartition(ZH_total, "区整体", QX_ZH_data, QX_ZH_group, fullmark, wenli.dt, db._group_num, true, ans.dt);
                    CalculatePartition(total, "区整体", QX_data, QX_group, sub_fullmark, groups.dt, db._group_num, false, wenli_standard);

                    string[] qxsf_code = CalculateTotal(QXSF_list);
                    DataTable qxsf_zh_data = QX_ZH_data.filteredtable("schoolcode", qxsf_code);
                    DataTable qxsf_zh_group = QX_ZH_group.filteredtable("schoolcode", qxsf_code);
                    DataTable qxsf_data = QX_data.filteredtable("schoolcode", qxsf_code);
                    DataTable qxsf_group = QX_group.filteredtable("schoolcode", qxsf_code);

                    qxsf_zh_data.SeperateGroups(grouptype, divider, "groups");
                    qxsf_zh_group.SeperateGroups(grouptype, divider, "groups");
                    qxsf_data.SeperateGroups(grouptype, divider, "groups");
                    qxsf_group.SeperateGroups(grouptype, divider, "groups");

                    CalculatePartition(ZH_total, "分类整体", qxsf_zh_data, qxsf_zh_group, fullmark, wenli.dt, db._group_num, true, ans.dt);
                    CalculatePartition(total, "分类整体", qxsf_data, qxsf_group, sub_fullmark, groups.dt, db._group_num, false, wenli_standard);
                    for (int i = 0; i < QXSF_list.Count; i++)
                    {
                        string[] SF_code = new string[QXSF_list[i].Count - 1];
                        for (int j = 1; j < QXSF_list[i].Count; j++)
                            SF_code[j - 1] = QXSF_list[i][j].ToString().Trim();
                        DataTable temp = qxsf_zh_data.filteredtable("schoolcode", SF_code);
                        DataTable temp_group = qxsf_zh_group.filteredtable("schoolcode", SF_code);

                        DataTable single = qxsf_data.filteredtable("schoolcode", SF_code);
                        DataTable single_table = qxsf_group.filteredtable("schoolcode", SF_code);
                        CalculatePartition(ZH_total, QXSF_list[i][0].ToString(), temp, temp_group, fullmark, wenli.dt, db._group_num, true, ans.dt);
                        CalculatePartition(total, QXSF_list[i][0].ToString(), single, single_table, sub_fullmark, groups.dt, db._group_num, false, wenli_standard);
                        CalculatePartition(ZH_QX, QXSF_list[i][0].ToString(), temp, temp_group, fullmark, wenli.dt, db._group_num, true, ans.dt);
                        CalculatePartition(QX, QXSF_list[i][0].ToString(), single, single_table, sub_fullmark, groups.dt, db._group_num, false, wenli_standard);
                    }
                    CalculatePartition(ZH_QX, "分类整体", qxsf_zh_data, qxsf_zh_group, fullmark, wenli.dt, db._group_num, true, ans.dt);
                    CalculatePartition(QX, "分类整体", qxsf_data, qxsf_group, sub_fullmark, groups.dt, db._group_num, false, wenli_standard);
                    form.ShowPro(70, 4);
                    Partition_wordcreator create = new Partition_wordcreator(total, QX, groups.dt, groups.groups_group);
                    create.creating_ZH_QX_word(ZH_total, ZH_QX, wenli.dt, wenli.groups_group);
                }
                else if (report_style.Equals("学校"))
                {
                    List<WSLG_partitiondata> total = new List<WSLG_partitiondata>();
                    List<WSLG_partitiondata> single = new List<WSLG_partitiondata>();

                    Utils.WSLG = true;
                    PartitionXXDataProcess(total, db._basic_data, db.zh_group_data, db._group_num, ans.dt, wenli.dt);
                    PartitionXXDataProcess(single, db.zh_single_data, db._group_data, db._group_num, wenli_standard, groups.dt);
                    foreach (KeyValuePair<string, string> kv in School_code)
                    {

                        PartitionXX(total, db._basic_data, db.zh_group_data, db._group_num, kv.Key, ans.dt, wenli.dt);
                        PartitionXX(single, db.zh_single_data, db._group_data, db._group_num, kv.Key, wenli_standard, groups.dt);
                        WordData temp_total = TotalSchoolCal(db._basic_data, db.zh_group_data, db._group_num, kv.Key, ans.dt, wenli.dt, true);
                        WordData temp_single = TotalSchoolCal(db.zh_single_data, db._group_data, db._group_num, kv.Key, wenli_standard, groups.dt, false);

                        SchoolWordCreator swc = new SchoolWordCreator(temp_single, single, wenli.dt, kv.Value, wenli.groups_group);
                        swc.creating_ZH_word(temp_total, total, groups.dt, groups.groups_group);
                    }

                }
            }
            else
            {

                GK_database db = new GK_database(ans.dt, groups.dt, grouptype, divider);
                db.DBF_data_process(database_address);

                if (db._basic_data.Columns.Contains("XZ"))
                {
                    XZ_group_separate(db._basic_data);
                }
                ans.dt = db.newStandard;
                form.ShowPro(40, 3);
                if (report_style.Equals("总体"))
                {
                    WordData data = new WordData(groups.groups_group);
                    Total_statistic stat = new Total_statistic(data, db._basic_data, fullmark, ans.dt, db._group_data, groups.dt, db._group_num);
                    stat.statistic_process(false);
                    if (db._basic_data.Columns.Contains("XZ"))
                        stat.xz_postprocess(ans.xz_th);
                    form.ShowPro(70, 4);
                    WordCreator create = new WordCreator(data);
                    create.creating_word();
                    if (subject.Equals("语文") || subject.Equals("英语"))
                    {
                        form.ShowPro(80, 6);
                        Utils.WSLG = true;
                        ArrayList WSLG = new ArrayList();


                        DataTable W_data = db._basic_data.Likefilter("studentid", "'1*'");
                        DataTable W_group = db._group_data.Likefilter("studentid", "'1*'");

                        Partition_statistic w_stat = new Partition_statistic("文科", W_data, fullmark, ans.dt, W_group, groups.dt, db._group_num);
                        w_stat.statistic_process(false);
                        if (db._basic_data.Columns.Contains("XZ"))
                            w_stat.xz_postprocess(ans.xz_th);
                        WSLG.Add(w_stat.result);

                        DataTable l_data = db._basic_data.Likefilter("studentid", "'5*'");
                        DataTable l_group = db._group_data.Likefilter("studentid", "'5*'");

                        Partition_statistic l_stat = new Partition_statistic("理科", l_data, fullmark, ans.dt, l_group, groups.dt, db._group_num);
                        l_stat.statistic_process(false);
                        if (db._basic_data.Columns.Contains("XZ"))
                            l_stat.xz_postprocess(ans.xz_th);
                        WSLG.Add(l_stat.result);

                        Partition_statistic total_stat = new Partition_statistic("分类整体", db._basic_data, fullmark, ans.dt, db._group_data, groups.dt, db._group_num);
                        total_stat.statistic_process(false);
                        if (db._basic_data.Columns.Contains("XZ"))
                            total_stat.xz_postprocess(ans.xz_th);
                        WSLG.Add(total_stat.result);

                        Partition_wordcreator create2 = new Partition_wordcreator(WSLG, groups.dt, groups.groups_group);
                        create2.creating_word();
                        Utils.WSLG = false;
                    }
                }
                else if (report_style.Equals("两类示范校"))
                {
                    ArrayList list = new ArrayList();
                    PartitionDataProcess(list, SF_list, "schoolcode", db._basic_data, db._group_data, db._group_num, false);
                    form.ShowPro(70, 4);
                    Partition_wordcreator create = new Partition_wordcreator(list, groups.dt, groups.groups_group);
                    create.creating_word();
                }
                else if (report_style.Equals("城郊"))
                {
                    ArrayList list = new ArrayList();
                    PartitionDataProcess(list, CJ_list, "QX", db._basic_data, db._group_data, db._group_num, false);
                    form.ShowPro(70, 4);
                    Partition_wordcreator create = new Partition_wordcreator(list, groups.dt, groups.groups_group);
                    create.creating_word();
                }
                else if (report_style.Equals("区县"))
                {
                    ArrayList QX = new ArrayList();
                    ArrayList total = new ArrayList();
                    PartitionQXDataProcess(total, QX, db._basic_data, db._group_data, db._group_num);
                    form.ShowPro(70, 4);
                    Partition_wordcreator create = new Partition_wordcreator(total, QX, groups.dt, groups.groups_group);
                    create.creating_word();

                    if (subject.Equals("语文") || subject.Equals("英语"))
                    {
                        form.ShowPro(80, 6);
                        Utils.WSLG = true;
                        ArrayList WSLG = new ArrayList();
                        DataTable QX_data = db._basic_data.filteredtable("schoolcode", PartitionSFcode(QXSF_list));
                        DataTable QX_group = db._group_data.filteredtable("schoolcode", PartitionSFcode(QXSF_list));

                        int group = QX_data.SeperateGroups(grouptype, divider, "groups");
                        QX_group.SeperateGroups(grouptype, divider, "groups");
                        if (db._basic_data.Columns.Contains("XZ"))
                            XZ_group_separate(QX_data);
                        DataTable W_data = QX_data.Likefilter("studentid", "'1*'");
                        DataTable W_group = QX_group.Likefilter("studentid", "'1*'");

                        Partition_statistic w_stat = new Partition_statistic("文科", W_data, fullmark, ans.dt, W_group, groups.dt, group);
                        w_stat.statistic_process(false);
                        if (db._basic_data.Columns.Contains("XZ"))
                            w_stat.xz_postprocess(ans.xz_th);
                        WSLG.Add(w_stat.result);

                        DataTable l_data = QX_data.Likefilter("studentid", "'5*'");
                        DataTable l_group = QX_group.Likefilter("studentid", "'5*'");

                        Partition_statistic l_stat = new Partition_statistic("理科", l_data, fullmark, ans.dt, l_group, groups.dt, group);
                        l_stat.statistic_process(false);
                        if (db._basic_data.Columns.Contains("XZ"))
                            l_stat.xz_postprocess(ans.xz_th);
                        WSLG.Add(l_stat.result);

                        Partition_statistic total_stat = new Partition_statistic("分类整体", QX_data, fullmark, ans.dt, QX_group, groups.dt, group);
                        total_stat.statistic_process(false);
                        if (db._basic_data.Columns.Contains("XZ"))
                            total_stat.xz_postprocess(ans.xz_th);
                        WSLG.Add(total_stat.result);

                        Partition_wordcreator create2 = new Partition_wordcreator(WSLG, groups.dt, groups.groups_group);
                        create2.creating_word();
                        Utils.WSLG = false;
                    }

                }
                else if (report_style.Equals("学校"))
                {
                    List<WSLG_partitiondata> total = new List<WSLG_partitiondata>();
                    Utils.WSLG = true;
                    PartitionXXDataProcess(total, db._basic_data, db._group_data, db._group_num, ans.dt, groups.dt);

                    foreach (KeyValuePair<string, string> kv in School_code)
                    {
                        
                        PartitionXX(total, db._basic_data, db._group_data, db._group_num, kv.Key, ans.dt, groups.dt);
                        WordData temp = TotalSchoolCal(db._basic_data, db._group_data, db._group_num, kv.Key, ans.dt, groups.dt, false);
                        SchoolWordCreator swc = new SchoolWordCreator(temp, total, groups.dt, kv.Value, groups.groups_group);
                        swc.creating_word();
                    }

                }
            }
        }
        WordData TotalSchoolCal(DataTable data, DataTable group, int groupnum, string school, DataTable my_ans, DataTable my_group, bool isZonghe)
        {
            DataTable XX = data.filteredtable("schoolcode", new string[] { school });
            DataTable XX_group = group.filteredtable("schoolcode", new string[] { school });

            XX.SeperateGroups(grouptype, Convert.ToDecimal(groupnum), "groups");
            XX_group.SeperateGroups(grouptype, Convert.ToDecimal(groupnum), "groups");

            WordData result = new WordData(groups.groups_group);
            Total_statistic stat = new Total_statistic(result, XX, fullmark, my_ans, XX_group, my_group, groupnum);
            stat.statistic_process(isZonghe);
            if (data.Columns.Contains("XZ"))
                stat.xz_postprocess(ans.xz_th);

            return result;
        }
        void PartitionXX(List<WSLG_partitiondata> result, DataTable data, DataTable group, int groupnum, string school, DataTable my_ans, DataTable my_group)
        {
            DataTable XX = data.filteredtable("schoolcode", new string[] { school });
            DataTable XX_group = group.filteredtable("schoolcode", new string[] { school });
            Partition_statistic XX_stat = new Partition_statistic("本学校", XX, fullmark, my_ans, XX_group, my_group, groupnum);
            XX_stat.statistic_process(false);
            if (data.Columns.Contains("XZ"))
                XX_stat.xz_postprocess(ans.xz_th);
            result.Insert(0, (WSLG_partitiondata)XX_stat.result);
        }
        void PartitionXXDataProcess(List<WSLG_partitiondata> result, DataTable data, DataTable group, int groupnum, DataTable my_ans, DataTable my_group)
        {
            
            DataTable QX = data.filteredtable("QX", QXTransfer(Quxian_list));
            DataTable QX_group = group.filteredtable("QX", QXTransfer(Quxian_list));
            Partition_statistic qx_stat = new Partition_statistic("区整体", QX, fullmark, my_ans, QX_group, my_group, groupnum);
            qx_stat.statistic_process(false);
            if (data.Columns.Contains("XZ"))
                qx_stat.xz_postprocess(ans.xz_th);
            result.Add((WSLG_partitiondata)qx_stat.result);


            Partition_statistic total = new Partition_statistic("市整体", data, fullmark, my_ans, group, my_group, groupnum);
            total.statistic_process(false);
            if (data.Columns.Contains("XZ"))
                total.xz_postprocess(ans.xz_th);
            result.Add((WSLG_partitiondata)total.result);

            for (int i = 0; i < CJ_list.Count; i++)
            {
                ArrayList cj = CJ_list[i];
                string[] xx_code = new string[cj.Count - 1];
                for (int j = 1; j < cj.Count; j++)
                    xx_code[j - 1] = cj[j].ToString().Trim();
                DataTable temp = data.filteredtable("QX", xx_code);
                DataTable temp_group = group.filteredtable("QX", xx_code);
                Partition_statistic stat = new Partition_statistic(cj[0].ToString(), temp, fullmark, my_ans, temp_group, my_group, groupnum);
                stat.statistic_process(false);
                if (data.Columns.Contains("XZ"))
                    stat.xz_postprocess(ans.xz_th);
                result.Add((WSLG_partitiondata)stat.result);
            }

            for (int i = 0; i < SF_list.Count; i++)
            {
                ArrayList sf = SF_list[i];
                string[] xx_code = new string[sf.Count - 1];
                for (int j = 1; j < sf.Count; j++)
                    xx_code[j - 1] = sf[j].ToString().Trim();
                DataTable temp = data.filteredtable("schoolcode", xx_code);
                DataTable temp_group = group.filteredtable("schoolcode", xx_code);
                Partition_statistic stat = new Partition_statistic(sf[0].ToString(), temp, fullmark, my_ans, temp_group, my_group, groupnum);
                stat.statistic_process(false);
                if (data.Columns.Contains("XZ"))
                    stat.xz_postprocess(ans.xz_th);
                result.Add((WSLG_partitiondata)stat.result);
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
        string[] PartitionSFcode(List<ArrayList> list)
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
            return SF_code;
        }
        void PartitionQXDataProcess(ArrayList result, ArrayList sresult, DataTable data, DataTable group, int groupnum)
        {
            Partition_statistic total = new Partition_statistic("市整体", data, fullmark, ans.dt, group, groups.dt, groupnum);
            total.statistic_process(false);
            if (data.Columns.Contains("XZ"))
                total.xz_postprocess(ans.xz_th);
            result.Add(total.result);

            for (int i = 0; i < SF_list.Count; i++)
            {
                ArrayList sf = SF_list[i];
                string[] xx_code = new string[sf.Count - 1];
                for (int j = 1; j < sf.Count; j++)
                    xx_code[j - 1] = sf[j].ToString().Trim();
                DataTable temp = data.filteredtable("schoolcode", xx_code);
                DataTable temp_group = group.filteredtable("schoolcode", xx_code);
                Partition_statistic stat = new Partition_statistic(sf[0].ToString(), temp, fullmark, ans.dt, temp_group, groups.dt, groupnum);
                stat.statistic_process(false);
                if(data.Columns.Contains("XZ"))
                    stat.xz_postprocess(ans.xz_th);
                result.Add(stat.result);
            }

            for (int i = 0; i < CJ_list.Count; i++)
            {
                ArrayList cj = CJ_list[i];
                string[] xx_code = new string[cj.Count - 1];
                for (int j = 1; j < cj.Count; j++)
                    xx_code[j - 1] = cj[j].ToString().Trim();
                DataTable temp = data.filteredtable("QX", xx_code);
                DataTable temp_group = group.filteredtable("QX", xx_code);
                Partition_statistic stat = new Partition_statistic(cj[0].ToString(), temp, fullmark, ans.dt, temp_group, groups.dt, groupnum);
                stat.statistic_process(false);
                if (data.Columns.Contains("XZ"))
                    stat.xz_postprocess(ans.xz_th);
                result.Add(stat.result);
            }
            
            DataTable QX = data.filteredtable("QX", QXTransfer(Quxian_list));
            DataTable QX_group = group.filteredtable("QX", QXTransfer(Quxian_list));
            Partition_statistic qx_stat = new Partition_statistic("区整体", QX, fullmark, ans.dt, QX_group, groups.dt, groupnum);
            qx_stat.statistic_process(false);
            if (data.Columns.Contains("XZ"))
                qx_stat.xz_postprocess(ans.xz_th);
            result.Add(qx_stat.result);
            PartitionDataProcess(result, QXSF_list, "schoolcode", QX, QX_group, groupnum, true);
            PartitionDataProcess(sresult, QXSF_list, "schoolcode", QX, QX_group, groupnum, false);

        }
        void PartitionDataProcess(ArrayList result, List<ArrayList> list, String filter, DataTable data, DataTable group, int groupnum, bool isQXSF)
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
            dt.SeperateGroups(grouptype, divider, "groups");
            dt_group.SeperateGroups(grouptype, divider, "groups");
            if (dt.Columns.Contains("XZ"))
                XZ_group_separate(dt);
            Partition_statistic total = new Partition_statistic("分类整体", dt, fullmark, ans.dt, dt_group, groups.dt, groupnum);
            total.statistic_process(false);
            if (dt.Columns.Contains("XZ"))
                total.xz_postprocess(ans.xz_th);
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
                Partition_statistic stat = new Partition_statistic(temp[0].ToString(), temp_dt, fullmark, ans.dt, temp_group, groups.dt, groupnum);
                stat.statistic_process(false);
                if (dt.Columns.Contains("XZ"))
                    stat.xz_postprocess(ans.xz_th);
                result.Add(stat.result);
            }
            if(!isQXSF)
                result.Add(total.result);
        }
        void CalculatePartition(ArrayList list, String title, DataTable total, DataTable group, decimal fullmark, DataTable group_ans, int groupnum, bool isZonghe, DataTable ans)
        {
            Partition_statistic stat = new Partition_statistic(title, total, fullmark, ans, group, group_ans, groupnum);
            stat.statistic_process(isZonghe);
            list.Add(stat.result);
        }
        string[] CalculateTotal(List<ArrayList> data)
        {
            int totalnum = 0;
            for (int i = 0; i < data.Count; i++)
                totalnum += (data[i].Count - 1);
            string[] SF_code = new string[totalnum];
            totalnum = 0;
            for (int i = 0; i < data.Count; i++)
            {
                for (int j = 1; j < data[i].Count; j++)
                {
                    SF_code[totalnum] = data[i][j].ToString().Trim();
                    totalnum++;
                }
            }
            return SF_code;

        }
        void CalculateGKZF(DataTable total, List<ZF_statistic> result)
        {
            int totalnum = 0;
            for (int i = 0; i < QXSF_list.Count; i++)
                totalnum += (QXSF_list[i].Count - 1);
            string[] SF_code = new string[totalnum];
            totalnum = 0;
            for (int i = 0; i < QXSF_list.Count; i++)
            {
                for (int j = 1; j < QXSF_list[i].Count; j++)
                {
                    SF_code[totalnum] = QXSF_list[i][j].ToString().Trim();
                    totalnum++;
                }
            }

            DataTable flztdata = total.filteredtable("schoolcode", SF_code);
            ZF_statistic flzt = new ZF_statistic(flztdata, fullmark, "分类整体");
            flzt.partition_process();
            result.Add(flzt);

            for (int i = 0; i < QXSF_list.Count; i++)
            {
                ArrayList temp = QXSF_list[i];
                string[] xx_code = new string[temp.Count - 1];
                for (int j = 1; j < temp.Count; j++)
                    xx_code[j - 1] = temp[j].ToString().Trim();
                DataTable data = flztdata.filteredtable("schoolcode", xx_code);
                ZF_statistic stat = new ZF_statistic(data, fullmark, temp[0].ToString().Trim());
                stat.partition_process();
                result.Add(stat);

            }

        }
        void CalculateClassTotal(DataTable total, DataTable groups_data, ArrayList totaldata, ArrayList sdata)
        {
            int totalnum = 0;
            for (int i = 0; i < QXSF_list.Count; i++)
                totalnum += (QXSF_list[i].Count - 1);
            string[] SF_code = new string[totalnum];
            totalnum = 0;
            for (int i = 0; i < QXSF_list.Count; i++)
            {
                for (int j = 1; j < QXSF_list[i].Count; j++)
                {
                    SF_code[totalnum] = QXSF_list[i][j].ToString().Trim();
                    totalnum++;
                }
            }
            DataTable ClassTotal_data = total.filteredtable("schoolcode", SF_code);
            DataTable ClassGroupTotal_data = groups_data.filteredtable("schoolcode", SF_code);

            int groupnum = ClassTotal_data.SeperateGroups(grouptype, divider, "groups");
            ClassGroupTotal_data.SeperateGroups(grouptype, divider, "groups");
            if (ClassTotal_data.Columns.Contains("XZ"))
                XZ_group_separate(ClassTotal_data);
            Partition_statistic ClassTotal = new Partition_statistic("分类整体", ClassTotal_data, fullmark, ans.dt, ClassGroupTotal_data, groups.dt, groupnum);
            ClassTotal.statistic_process(false);
            if (ClassTotal_data.Columns.Contains("XZ"))
                ClassTotal.xz_postprocess(ans.xz_th);
            totaldata.Add(ClassTotal.result);

            for (int i = 0; i < QXSF_list.Count; i++)
            {
                ArrayList temp = QXSF_list[i];
                string[] xx_code = new string[temp.Count - 1];
                for (int j = 1; j < temp.Count; j++)
                    xx_code[j - 1] = temp[j].ToString().Trim();
                DataTable xx_data = ClassTotal_data.filteredtable("schoolcode", xx_code);
                DataTable xx_group_data = ClassGroupTotal_data.filteredtable("schoolcode", xx_code);

                Partition_statistic XXTotal = new Partition_statistic(temp[0].ToString().Trim(), xx_data, fullmark, ans.dt, xx_group_data, groups.dt, groupnum);
                XXTotal.statistic_process(false);
                if (ClassTotal_data.Columns.Contains("XZ"))
                    XXTotal.xz_postprocess(ans.xz_th);
                totaldata.Add(XXTotal.result);
                sdata.Add(XXTotal.result);

            }
            sdata.Add(ClassTotal.result);
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
            
            //switch (QX) 
            //{
            //    case "东城区":
            //        string[] dongcheng = { "01", "03" };
            //        return dongcheng;
            //    case "西城区":
            //        string[] xicheng = { "02" };
            //        return xicheng;
            //    case "朝阳区":
            //        string[] chaoyang = { "05" };
            //        return chaoyang;
            //    case "丰台区":
            //        string[] fengtai = { "06" };
            //        return fengtai;
            //    case "石景山区":
            //        string[] shijingshan = { "07" };
            //        return shijingshan;
            //    case "海淀区":
            //        string[] haidian = { "08" };
            //        return haidian;
            //    case "门头沟区":
            //        string[] mentougou = { "09" };
            //        return mentougou;
            //    case "燕山区":
            //        string[] yanshan = { "10" };
            //        return yanshan;
            //    case "房山区":
            //        string[] fangshan = { "11" };
            //        return fangshan;
            //    case "通州区":
            //        string[] tongzhou = { "12" };
            //        return tongzhou;
            //    case "顺义区":
            //        string[] shunyi = { "13" };
            //        return shunyi;
            //    case "昌平区":
            //        string[] changping = { "14" };
            //        return changping;
            //    case "大兴区":
            //        string[] daxing = { "15" };
            //        return daxing;
            //    case "怀柔区":
            //        string[] huairou = { "16" };
            //        return huairou;
            //    case "平谷区":
            //        string[] pinggu = { "17" };
            //        return pinggu;
            //    case "密云县":
            //        string[] miyun = { "28" };
            //        return miyun;
            //    case "延庆县":
            //        string[] yanqing = { "29" };
            //        return yanqing;
            //    default:
            //        string[] def = { "" };
            //        return def;
            //}
        }

    }
}
