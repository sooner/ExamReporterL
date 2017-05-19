using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using Telerik.WinControls.UI;
using System.Data;
using System.IO;

namespace ExamReport
{
    public delegate void gk_process(MetaData mdata);
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
        public string QX = "";
        public decimal _fullmark;
        public decimal _sub_fullmark;
        public bool is_sub_cor = false;

        public string curryear;
        public string currmonth;

        public ZK_database.GroupType _grouptype;
        public int _group_num = 0;

        public DataTable custom_data;


        public Dictionary<string, string> school;
        public Dictionary<string, string> school_qx;
        public string school_name;
        public string school_code;


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

        public Utils.UnionType hk_script_type;
        public string stu_id;
        public string qx;
        public string qx_name;
        public string xx;
        public string xx_name;
        public string date;

        public int wen_first_level;
        public int wen_second_level;
        public int wen_third_level;

        public int li_first_level;
        public int li_second_level;
        public int li_third_level;

        public int zk_first_level;
        public int zk_second_level;

        public string compare_year1;
        public string compare_year2;

        public bool compare_cj_is_wk;
        public string cj_comp_year;

        public Analysis(mainform form)
        {
            _form = form;
        }
        public void start()
        {
            _form.ShowPro(exam_type, 0, "开始处理...");
            foreach (GridViewRowInfo row in _gridview.Rows)
            {
                if (row.Cells["checkbox"].Value != null && (bool)row.Cells["checkbox"].Value == true)
                {
                    string year = row.Cells["year"].Value.ToString().Trim();
                    string exam = _exam;
                    string chi_sub = row.Cells["sub"].Value.ToString().Trim();
                    string sub = Utils.language_trans(chi_sub);

                    if (exam_type.Equals("gk_xz") && !sub.EndsWith("xz"))
                        continue;
                    string log = year + "年" + Utils.language_trans(exam) + row.Cells["sub"].Value.ToString().Trim();
                    _form.ShowPro(exam_type, 1, log + "数据读取...");
                    MetaData mdata = new MetaData(year, exam, sub);
                    //try
                    //{
                        mdata.get_meta_data();
                        _fullmark = mdata._fullmark;

                        _grouptype = mdata._grouptype;
                        _group_num = mdata._group_num;

                        if (_exam.Equals("gk") && (sub.Equals("yy") || sub.Equals("yw")))
                            mdata.ywyy_choice = row.Cells["SpecChoice"].Value.ToString().Trim();

                        mdata.log_name = log;
                        if (sub.Equals("zf") || sub.EndsWith("xz"))
                        {
                            mdata.get_zf_data();
                        }
                        else
                        {
                            mdata.get_basic_data();
                            mdata.get_group_data();
                            mdata.get_ans();
                            mdata.get_fz();
                        }
                        if (Utils.is_gk_zh(exam, chi_sub))
                        {
                            mdata.get_zh_basic_data();
                            mdata.get_zh_group_data();
                            mdata.get_zh_ans();
                            mdata.get_zh_fz();

                            mdata.basic = SortTable(mdata.basic, "ZH_totalmark");
                            mdata.group = SortTable(mdata.group, "ZH_totalmark");
                            mdata.basic.SeperateGroupsByColumnName(mdata._grouptype, mdata._group_num, "ZH_totalmark");
                            mdata.group.SeperateGroupsByColumnName(mdata._grouptype, mdata._group_num, "ZH_totalmark");

                            if (row.Cells["SpecChoice"].Value.ToString().Trim().Equals("科目总分相关"))
                            {
                                is_sub_cor = true;
                                _sub_fullmark = mdata._sub_fullmark;

                                
                                //List<decimal> res = new List<decimal>();
                                //List<string> flag = new List<string>();
                                //for (int i = 0; i < mdata.basic.Rows.Count; i++)
                                //{
                                //    res.Add((decimal)mdata.basic.Rows[i]["ZH_totalmark"]);
                                //    flag.Add((string)mdata.basic.Rows[i]["groups"]);
                                //}
                                //decimal num = res[40];
                            }
                            else
                            {
                                is_sub_cor = false;
                                _sub_fullmark = mdata._fullmark;

                                //mdata.basic.SeperateGroupsByColumnName(mdata._grouptype, mdata._group_num, "totalmark");
                                //mdata.group.SeperateGroupsByColumnName(mdata._grouptype, mdata._group_num, "totalmark");
                            }
                            switch (exam_type)
                            {
                                case "gk_zt":
                                    gk_zh_zt_process(mdata);
                                    break;
                                case "gk_cj":
                                    mdata.get_CJ_data(cj_addr);
                                    gk_zh_cj_process(mdata);
                                    break;
                                case "gk_sf":
                                    mdata.get_SF_data(sf_addr);
                                    gk_zh_sf_process(mdata);
                                    break;
                                case "gk_qx":
                                    mdata.get_QXSF_data(qx_addr);
                                    mdata.get_CJ_data(cj_addr);
                                    mdata.get_SF_data(sf_addr);
                                    gk_zh_qx_process(mdata);
                                    break;
                                case "gk_xx":
                                    mdata.get_CJ_data(cj_addr);
                                    mdata.get_SF_data(sf_addr);
                                    gk_zh_xx_process(mdata);
                                    break;
                                case "gk_cus":
                                    gk_zh_cus_process(mdata);
                                    break;
                                default:
                                    break;
                            }
                        }
                        else if (sub.Equals("zf"))
                        {
                            switch (exam_type)
                            {
                                case "gk_zt":
                                    gk_zf_zt_process(mdata);
                                    break;
                                case "gk_cj":
                                    mdata.get_CJ_data(cj_addr);
                                    gk_zf_cj_process(mdata);
                                    break;
                                case "gk_sf":
                                    mdata.get_SF_data(sf_addr);
                                    gk_zf_sf_process(mdata);
                                    break;
                                case "gk_qx":
                                    mdata.get_QXSF_data(qx_addr);
                                    mdata.get_CJ_data(cj_addr);
                                    mdata.get_SF_data(sf_addr);
                                    gk_zf_qx_process(mdata);
                                    break;
                                case "gk_xx":
                                    mdata.get_CJ_data(cj_addr);
                                    mdata.get_SF_data(sf_addr);
                                    gk_zf_xx_process(mdata);
                                    break;
                                case "gk_cus":
                                    gk_zf_custom_process(mdata);
                                    break;
                                default:
                                    break;
                            }
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
                                case "zk_xz":
                                    mdata.get_CJ_data(cj_addr);
                                    zk_xz_process(mdata);
                                    break;
                                case "hk_zt":
                                    hk_zt_process(mdata);
                                    break;
                                case "hk_script":
                                    hk_script_process(mdata);
                                    break;
                                case "gk_zt":
                                    gk_schedule(mdata, gk_zt_process, gk_zt_wl_process);
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
                                    gk_schedule(mdata, gk_qx_process, gk_qx_wl_process);
                                    break;
                                case "gk_xx":
                                    mdata.get_CJ_data(cj_addr);
                                    mdata.get_SF_data(sf_addr);
                                    gk_schedule(mdata, gk_xx_process, gk_xx_wl_process);
                                    break;
                                case "gk_cus":
                                    gk_cus_process(mdata);
                                    break;
                                case "gk_xz":
                                    mdata.get_CJ_data(cj_addr);
                                    gk_xz_process(mdata);
                                    break;
                                default:
                                    break;
                            }

                        }
                    //}
                    //catch (System.Threading.ThreadAbortException e)
                    //{
                    //}
                    //catch (Exception e)
                    //{
                    //    _form.ErrorM(exam_type, e.Message);
                    //}

                }
            }
            _form.ShowPro(exam_type, 2, "完成！");
        }
        public void gk_cj_comp_start()
        {
            _form.ShowPro("gk_cj_cp", 0, "对比开始处理...");
            MetaData mdata = new MetaData(cj_comp_year, "gk", "gk_xz");
            mdata.get_zf_data();
            mdata.get_CJ_data(Utils.CurrentDirectory + @"\config\gk_cj.xlsx");

            string[] urban_code = new string[mdata.CJ_list[0].Count - 1];
            for (int j = 1; j < mdata.CJ_list[0].Count; j++)
                urban_code[j - 1] = mdata.CJ_list[0][j].ToString().Trim();

            string[] country_code = new string[mdata.CJ_list[1].Count - 1];
            for (int j = 1; j < mdata.CJ_list[1].Count; j++)
                country_code[j - 1] = mdata.CJ_list[1][j].ToString().Trim();

            DataTable urban = mdata.basic.filteredtable("qxdm", urban_code);
            DataTable country = mdata.basic.filteredtable("qxdm", country_code);

            DataTable total;
            if (compare_cj_is_wk)
                total = mdata.basic.equalfilter("type", "w");
            else
                total = mdata.basic.equalfilter("type", "l");
            
            GKCompWordCreator wc = new GKCompWordCreator();
            wc.is_wk = compare_cj_is_wk;
            wc.year = cj_comp_year;
            wc.pre_process(total, urban, country);
            wc.creating_word();
            _form.ShowPro("gk_cj_cp", 2, "完成！");
        }
        public void gk_comp_start()
        {
            try
            {
                _form.ShowPro("gk_cp", 0, "历年对比开始处理...");
                Comparison comp = new Comparison();
                comp.start(compare_year1, compare_year2, "gk");
                Compare_wordcreator comp_wc = new Compare_wordcreator();
                comp_wc.year1_comb = comp.year1_comb;
                comp_wc.year2_comb = comp.year2_comb;
                comp_wc.year1_data = comp.year1_data;
                comp_wc.year2_data = comp.year2_data;
                comp_wc.creating_word();
                _form.ShowPro("gk_cp", 2, "完成！");
            }
            catch(Exception e)
            {
                _form.ErrorM("gk_cp", e.Message);
            }
            
        }
        public void zk_xz_start()
        {
            exam_type = "zk_xz";
            _exam = "zk";
            start();
        }
        public void zk_xz_process(MetaData mdata)
        {
            Configuration config = initConfig(mdata._sub, "行政", "中考");
            config.zk_first_level = zk_first_level;
            config.zk_second_level = zk_second_level;

            int urban = mdata.CJ_list[0][0].Equals("城区") ? 0 : 1;
            int country = mdata.CJ_list[0][1].Equals("郊区") ? 0 : 1;

            config.urban_code = new string[mdata.CJ_list[urban].Count];
            config.country_code = new string[mdata.CJ_list[country].Count];

            for (int j = 1; j < mdata.CJ_list[urban].Count; j++)
                config.urban_code[j - 1] = mdata.CJ_list[urban][j].ToString().Trim();

            for (int j = 1; j < mdata.CJ_list[country].Count; j++)
                config.country_code[j - 1] = mdata.CJ_list[country][j].ToString().Trim();

            _form.ShowPro("zk_xz", 1, mdata.log_name + "数据分析中...");
            AdminCal xzCal = new AdminCal(config, mdata.basic, mdata._fullmark, true);
            xzCal.zk_Calculate();

            _form.ShowPro("zk_xz", 1, mdata.log_name + "报告生成中...");
            ZK_Admin_WordCreator xzWordCreator = new ZK_Admin_WordCreator(config);
            xzWordCreator.creating_word(xzCal.final_result);
        }
        public void gk_xz_start()
        {
            exam_type = "gk_xz";
            _exam = "gk";
            start();
        }

        public void gk_xz_process(MetaData mdata)
        {
            Configuration config = initConfig(mdata._sub, "行政", "高考");
            config.wen_first_level = wen_first_level;
            config.wen_second_level = wen_second_level;
            config.wen_third_level = wen_third_level;

            config.li_first_level = li_first_level;
            config.li_second_level = li_second_level;
            config.li_third_level = li_third_level;

            int urban = mdata.CJ_list[0][0].Equals("城区") ? 0 : 1;
            int country = mdata.CJ_list[0][1].Equals("郊区") ? 0 : 1;

            config.urban_code = new string[mdata.CJ_list[urban].Count];
            config.country_code = new string[mdata.CJ_list[country].Count];

            for (int j = 1; j < mdata.CJ_list[urban].Count; j++)
                config.urban_code[j-1] = mdata.CJ_list[urban][j].ToString().Trim();

            for (int j = 1; j < mdata.CJ_list[country].Count; j++)
                config.country_code[j - 1] = mdata.CJ_list[country][j].ToString().Trim();

            _form.ShowPro("gk_xz", 1, mdata.log_name + "数据分析中...");
            AdminCal xzCal = new AdminCal(config, mdata.basic, mdata._fullmark);
            xzCal.Calculate();

            _form.ShowPro("gk_xz", 1, mdata.log_name + "报告生成中...");
            AdminWordCreator xzWordCreator = new AdminWordCreator(config);
            xzWordCreator.creating_word(xzCal.w_result, xzCal.l_result);
        }


        public List<DataRow> SearchIndex = new List<DataRow>();
        public List<long> NotFoundIndex = new List<long>();
        public void hk_script_start()
        {
            exam_type = "hk_script";
            _exam = "hk";
            start();
        }

        public void hk_script_process(MetaData mdata)
        {
            try
            {
                Configuration config = initConfig(mdata._sub, "报告单", "会考");

                if (hk_script_type == Utils.UnionType.ID)
                {
                    if (FindRow(stu_id, mdata.basic) == null)
                        _form.CheckStuID(2, "找不到该ID学生成绩！");
                }

                _form.ShowPro(exam_type, 1, mdata.log_name + "数据分析中...");
                //HK_worddata result = new HK_worddata(mdata.groups_group);
                //Total_statistic stat = new Total_statistic(result, mdata.basic, mdata._fullmark, mdata.ans, mdata.group, mdata.grp, mdata._group_num);
                //stat._config = config;
                //stat.statistic_process(false);
                //stat.HK_postprocess(hk_hierarchy);
                
                
                HKScriptCalculate cal = new HKScriptCalculate();
                cal.preprocess(mdata.group, hk_hierarchy);
                cal.execute(mdata.ans, mdata.grp, mdata.group);
                _form.ShowPro(exam_type, 1, mdata.log_name + "报告生成中...");
                HKScriptWordCreator wc = new HKScriptWordCreator();
                if (hk_script_type == Utils.UnionType.ID)
                {
                    DataRow dr = FindRow(stu_id, cal._data);
                    DataRow basic_dr = FindRow(stu_id, mdata.basic);
                    wc.create_word(cal.total, dr, mdata.grp, mdata.groups_group, basic_dr, CurrentDirectory, date);
                }
                else
                {

                    if (CurrentDirectory[CurrentDirectory.Length - 1] != Path.DirectorySeparatorChar)
                        CurrentDirectory += Path.DirectorySeparatorChar;
                    string save_adr = CurrentDirectory + "会考成绩报告单";
                    save_adr += Path.DirectorySeparatorChar;
                    if (!Directory.Exists(save_adr))
                        Directory.CreateDirectory(save_adr);
                    if (qx.Equals("0") || xx.Equals("0"))
                    {

                        if (qx.Equals("0"))
                        {
                            var qxtuples = from row in _form.schoolcode_table.AsEnumerable()
                                          group row by row.Field<string>("qxmc") into grp
                                          select new
                                          {
                                              qx = grp.Key.ToString().Trim(),
                                              code = string.Join(",", grp.GroupBy(p => p.Field<string>("qxdm")).Select(p => p.Key.ToString().Trim()).ToArray())
                                          };
                            foreach (var qxtuple in qxtuples)
                                QXCreateWord(cal, mdata, wc, qxtuple.qx, qxtuple.code, save_adr);
                        }
                        else
                        {
                            QXCreateWord(cal, mdata, wc, qx_name, qx, save_adr);
                        }

                    }
                    else
                    {
                        XXCreateWord(cal,mdata, wc, xx_name, xx, save_adr);
                    }
                }
            }
            catch (System.Threading.ThreadAbortException e)
            {
            }
            catch (Exception e)
            {
                _form.CheckStuID(2, e.Message.ToString());
            }
        }
        public void QXCreateWord(HKScriptCalculate cal, MetaData mdata, HKScriptWordCreator wc, string qx, string qx_code, string adr)
        {
            string save_adr = adr + qx;
            save_adr += Path.DirectorySeparatorChar;
            if (!Directory.Exists(save_adr))
                Directory.CreateDirectory(save_adr);
            var schooltuple = from row in _form.schoolcode_table.AsEnumerable()
                              where qx_code.Split(new char[1] { ',' }).Contains(row.Field<string>("xxdm").ToString().Trim().Substring(0, 2))
                              select new
                              {
                                  school = row.Field<string>("xxmc").Trim(),
                                  code = row.Field<string>("xxdm").Trim()
                              };
            foreach (var xx in schooltuple)
            {
                XXCreateWord(cal, mdata, wc, xx.school, xx.code, save_adr);
            }
            Utils.zip(save_adr, qx);
        }
        public void XXCreateWord(HKScriptCalculate cal, MetaData mdata, HKScriptWordCreator wc, string xx, string xx_code, string save_adr)
        {
            string adr = save_adr + xx;
            adr += Path.DirectorySeparatorChar;
            if (!Directory.Exists(adr))
                Directory.CreateDirectory(adr);
            List<string> stu_id = mdata.basic.AsEnumerable().Where(c => c.Field<string>("xxdm").Equals(xx_code)).Select(c => c.Field<string>("studentid")).ToList();
            foreach (string id in stu_id)
            {
                DataRow dr = FindRow(id, cal._data);
                DataRow basic_dr = FindRow(id, mdata.basic);
                wc.create_word(cal.total, dr, mdata.grp, mdata.groups_group, basic_dr, adr, date);
            }
            Utils.zip(adr, xx);
        }
        public DataRow FindRow(string stu_id, DataTable dt)
        {
            return dt.Rows.Find(stu_id);
        }
        public void FindRows(string stu_id_start, string stu_id_stop, DataTable dt)
        {
            char[] idChar = stu_id_start.Trim().ToCharArray();
            int size = idChar.Length;
            long start = Convert.ToInt64(stu_id_start);
            long stop = Convert.ToInt64(stu_id_stop);

            if (start > stop)
                throw new System.ArgumentException("起始学生ID不能大于终止学生ID!");

            SearchIndex.Clear();
            NotFoundIndex.Clear();

            for (long i = start; i <= stop; i++)
            {
                char[] temp = i.ToString().ToCharArray();
                StringBuilder sb = new StringBuilder(i.ToString());

                if (temp.Length < size)
                {
                    for (int j = temp.Length; j < size; j++)
                        sb.Insert(0, "0");
                }
                DataRow dr = FindRow(sb.ToString(), dt);
                if (dr == null)
                    NotFoundIndex.Add(i);
                else
                    SearchIndex.Add(dr);

            }
        }
        public void gk_custom_start()
        {
            exam_type = "gk_cus";
            _exam = "gk";
            start();
        }
        public void gk_zf_custom_process(MetaData mdata)
        {
            Configuration config = initConfig(mdata._sub, "自定义", "高考");
            List<ZF_statistic> result = new List<ZF_statistic>();
            _form.ShowPro(exam_type, 1, mdata.log_name + "数据分析中...");
            foreach (DataRow dr in custom_data.Rows)
            {
                DataTable temp = get_filterdata(mdata.basic, dr["condition"].ToString().Trim());
                ZF_statistic stat = new ZF_statistic(config, temp, mdata._fullmark, dr["name"].ToString().Trim());
                stat.partition_process();
                result.Add(stat);
            }
            _form.ShowPro(exam_type, 1, mdata.log_name + "报告生成中...");
            ZF_wordcreator create = new ZF_wordcreator(config);
            create.partition_wordcreate(result, "自定义");
        }
        public void gk_zh_cus_process(MetaData mdata)
        {
            Configuration config = initConfig(mdata._sub, "自定义", "高考");
            ArrayList sdata = new ArrayList();
            ArrayList ZH_data = new ArrayList();
            _form.ShowPro(exam_type, 1, mdata.log_name + "数据分析中...");

            string totalfilter = "";
            int count = 0;
            foreach (DataRow dr in custom_data.Rows)
            {
                string condition = dr["condition"].ToString().Trim();
                count++;
                totalfilter += "(" + condition + ")";
                if (count != custom_data.Rows.Count)
                    totalfilter += " OR ";
            }

            DataTable basic = SortTable(get_filterdata(mdata.basic, totalfilter), "ZH_totalmark");
            DataTable group = SortTable(get_filterdata(mdata.group, totalfilter), "ZH_totalmark");

            DataTable zh_basic = SortTable(get_filterdata(mdata.zh_basic, totalfilter), "totalmark");
            DataTable zh_group = SortTable(get_filterdata(mdata.zh_group, totalfilter), "totalmark");

            basic.SeperateGroups(mdata._grouptype, mdata._group_num, "groups");
            group.SeperateGroups(mdata._grouptype, mdata._group_num, "groups");

            zh_basic.SeperateGroups(mdata._grouptype, mdata._group_num, "groups");
            zh_group.SeperateGroups(mdata._grouptype, mdata._group_num, "groups");

            mdata.basic = basic;
            mdata.group = group;

            mdata.zh_basic = zh_basic;
            mdata.zh_group = zh_group;

            foreach (DataRow dr in custom_data.Rows)
            {
                string name = dr["name"].ToString().Trim();
                string condition = dr["condition"].ToString().Trim();

                CustomDataProcess(sdata, name, mdata, condition, config, mdata._sub_fullmark);
                ZHCustomDataProcess(ZH_data, name, mdata, condition, config);
            }

            CustomDataProcess(sdata, "分类整体", mdata, totalfilter, config, mdata._sub_fullmark);
            ZHCustomDataProcess(ZH_data, "分类整体", mdata, totalfilter, config);

            _form.ShowPro(exam_type, 1, mdata.log_name + "报告生成中...");
            Partition_wordcreator create = new Partition_wordcreator(sdata, mdata.grp, mdata.groups_group);
            create.SetConfig(config);
            create.creating_ZH_word(ZH_data, mdata.zh_grp, mdata.zh_groups_group);
        }
        public void ZHCustomDataProcess(ArrayList list, string name, MetaData mdata, string filter, Configuration config)
        {
            DataTable basic = get_filterdata(mdata.zh_basic, filter);
            DataTable group = get_filterdata(mdata.zh_group, filter);
            if (basic.Rows.Count == 0)
                throw new ArgumentException("条件 " + filter + " 没有数据！");
            //basic.SeperateGroups(mdata._grouptype, mdata._group_num, "groups");
            //group.SeperateGroups(mdata._grouptype, mdata._group_num, "groups");
            Partition_statistic total = new Partition_statistic(name, basic, mdata._fullmark, mdata.zh_ans, group, mdata.zh_grp, mdata._group_num);
            total._config = config;
            total.statistic_process(true);
            list.Add(total.result);
        }
        public void gk_cus_process(MetaData mdata)
        {
            Configuration config = initConfig(mdata._sub, "自定义", "高考");
            ArrayList list = new ArrayList();
            _form.ShowPro(exam_type, 1, mdata.log_name + "数据分析中...");

            string totalfilter = "";
            int count = 0;


            foreach (DataRow dr in custom_data.Rows)
            {
                string condition = dr["condition"].ToString().Trim();
                count++;
                totalfilter += "(" + condition + ")";
                if (count != custom_data.Rows.Count)
                    totalfilter += " OR ";
            }

            DataTable basic = SortTable(get_filterdata(mdata.basic, totalfilter), "totalmark");
            DataTable group = SortTable(get_filterdata(mdata.group, totalfilter), "totalmark");

            basic.SeperateGroups(mdata._grouptype, mdata._group_num, "groups");
            group.SeperateGroups(mdata._grouptype, mdata._group_num, "groups");

            mdata.basic = basic;
            mdata.group = group;

            foreach (DataRow dr in custom_data.Rows)
            {
                string name = dr["name"].ToString().Trim();
                string condition = dr["condition"].ToString().Trim();

                CustomDataProcess(list, name, mdata, condition, config, mdata._fullmark);
            }

            CustomDataProcess(list, "分类整体", mdata, totalfilter, config, mdata._fullmark);
            _form.ShowPro(exam_type, 1, mdata.log_name + "报告生成中...");
            Partition_wordcreator create = new Partition_wordcreator(list, mdata.grp, mdata.groups_group);
            create.SetConfig(config);
            create.creating_word();
        }
        public void CustomDataProcess(ArrayList list, string name, MetaData mdata, string filter, Configuration config, decimal fullmark)
        {
            DataTable basic = get_filterdata(mdata.basic, filter);
            DataTable group = get_filterdata(mdata.group, filter);
            if (basic.Rows.Count == 0)
                throw new ArgumentException("条件 " + filter + " 没有数据！");
            //basic.SeperateGroups(mdata._grouptype, mdata._group_num, "groups");
            //group.SeperateGroups(mdata._grouptype, mdata._group_num, "groups");
            //if (basic.Columns.Contains("XZ"))
            //    XZ_group_separate(basic, mdata);
            Partition_statistic total = new Partition_statistic(name, basic, fullmark, mdata.ans, group, mdata.grp, mdata._group_num);
            total._config = config;
            total.statistic_process(false);
            if (mdata.xz.Count > 0)
                total.xz_postprocess(mdata.xz);
            list.Add(total.result);
        }

        public DataTable get_filterdata(DataTable dt, string filter)
        {
            DataView dv = dt.AsDataView();
            dv.RowFilter = filter;
            return dv.ToTable();
        }
        public void hk_zt_start()
        {
            exam_type = "hk_zt";
            _exam = "hk";
            start();
        }
        public void hk_zt_process(MetaData mdata)
        {
            Configuration config = initConfig(mdata._sub, "总体", "会考");
            _form.ShowPro(exam_type, 1, mdata.log_name + "数据分析中...");
            HK_worddata result = new HK_worddata(mdata.groups_group);
            Total_statistic stat = new Total_statistic(result, mdata.basic, mdata._fullmark, mdata.ans, mdata.group, mdata.grp, mdata._group_num);
            stat._config = config;
            stat.statistic_process(false);
            stat.HK_postprocess(hk_hierarchy);
            if (mdata.xz.Count > 0)
                stat.xz_postprocess(mdata.xz);
            _form.ShowPro(exam_type, 1, mdata.log_name + "报告生成中...");
            WordCreator create = new WordCreator(result, config);
            create.creating_HK_word();
        }
        public void gk_zf_xx_process(MetaData mdata)
        {
            Configuration config = initConfig(mdata._sub, "学校", "高考");
            List<ZF_statistic> result = new List<ZF_statistic>();
            _form.ShowPro(exam_type, 1, mdata.log_name + "数据分析中...");
            ZF_statistic total = new ZF_statistic(config, mdata.basic, mdata._fullmark, "市整体");
            total.partition_process();
            result.Add(total);
            for (int i = 0; i < mdata.SF_list.Count; i++)
            {
                string[] SF_code = new string[mdata.SF_list[i].Count - 1];
                for (int j = 1; j < mdata.SF_list[i].Count; j++)
                    SF_code[j - 1] = mdata.SF_list[i][j].ToString().Trim();
                DataTable temp = mdata.basic.filteredtable("xxdm", SF_code);
                ZF_statistic stat = new ZF_statistic(config, temp, mdata._fullmark, mdata.SF_list[i][0].ToString().Trim());
                stat.partition_process();
                result.Add(stat);
            }
            for (int i = 0; i < mdata.CJ_list.Count; i++)
            {
                string[] cj_code = new string[mdata.CJ_list[i].Count - 1];
                for (int j = 1; j < mdata.CJ_list[i].Count; j++)
                    cj_code[j - 1] = mdata.CJ_list[i][j].ToString().Trim();
                DataTable temp = mdata.basic.filteredtable("qxdm", cj_code);
                ZF_statistic stat = new ZF_statistic(config, temp, mdata._fullmark, mdata.CJ_list[i][0].ToString().Trim());
                stat.partition_process();
                result.Add(stat);
            }
            DataTable bq_data = mdata.basic.filteredtable("qxdm", QXTransfer(school_qx[school_code]));
            ZF_statistic bq = new ZF_statistic(config, bq_data, mdata._fullmark, "本区");
            bq.partition_process();
            result.Add(bq);


            DataTable bx_data = mdata.basic.filteredtable("xxdm", new string[] { school_code });
            ZF_statistic bx = new ZF_statistic(config, bx_data, mdata._fullmark, "本校");
            bx.partition_process();
            List<ZF_statistic> temp_result = new List<ZF_statistic>();
            temp_result.AddRange(result);
            temp_result.Add(bx);
            _form.ShowPro(exam_type, 1, mdata.log_name + "报告生成中...");
            ZF_wordcreator create = new ZF_wordcreator(config, temp_result, school_name);
            create.XX_create();
            
            
        }
        public void gk_zf_qx_process(MetaData mdata)
        {
            Configuration config = initConfig(mdata._sub, "区县", "高考");
            _form.ShowPro(exam_type, 1, mdata.log_name + "数据分析中...");
            List<ZF_statistic> result = new List<ZF_statistic>();
            ZF_statistic total = new ZF_statistic(config, mdata.basic, mdata._fullmark, "市整体");
            total.partition_process();
            result.Add(total);
            for (int i = 0; i < mdata.SF_list.Count; i++)
            {
                string[] SF_code = new string[mdata.SF_list[i].Count - 1];
                for (int j = 1; j < mdata.SF_list[i].Count; j++)
                    SF_code[j - 1] = mdata.SF_list[i][j].ToString().Trim();
                DataTable temp = mdata.basic.filteredtable("xxdm", SF_code);
                ZF_statistic stat = new ZF_statistic(config, temp, mdata._fullmark, mdata.SF_list[i][0].ToString().Trim());
                stat.partition_process();
                result.Add(stat);
            }
            for (int i = 0; i < mdata.CJ_list.Count; i++)
            {
                string[] cj_code = new string[mdata.CJ_list[i].Count - 1];
                for (int j = 1; j < mdata.CJ_list[i].Count; j++)
                    cj_code[j - 1] = mdata.CJ_list[i][j].ToString().Trim();
                DataTable temp = mdata.basic.filteredtable("qxdm", cj_code);
                ZF_statistic stat = new ZF_statistic(config, temp, mdata._fullmark, mdata.CJ_list[i][0].ToString().Trim());
                stat.partition_process();
                result.Add(stat);
            }
            DataTable bq_data = mdata.basic.filteredtable("qxdm", QXTransfer(qx_code));
            ZF_statistic bq = new ZF_statistic(config, bq_data, mdata._fullmark, "本区");
            bq.partition_process();
            result.Add(bq);
            CalculateGKZF(config, mdata, bq_data, result);
            _form.ShowPro(exam_type, 1, mdata.log_name + "报告生成中...");
            ZF_wordcreator create = new ZF_wordcreator(config);
            create.partition_wordcreate(result, "区县");

        }
        void CalculateGKZF(Configuration config, MetaData mdata, DataTable total, List<ZF_statistic> result)
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

            DataTable flztdata = total.filteredtable("xxdm", SF_code);
            ZF_statistic flzt = new ZF_statistic(config, flztdata, mdata._fullmark, "分类整体");
            flzt.partition_process();
            result.Add(flzt);

            for (int i = 0; i < mdata.QXSF_list.Count; i++)
            {
                ArrayList temp = mdata.QXSF_list[i];
                string[] xx_code = new string[temp.Count - 1];
                for (int j = 1; j < temp.Count; j++)
                    xx_code[j - 1] = temp[j].ToString().Trim();
                DataTable data = flztdata.filteredtable("xxdm", xx_code);
                ZF_statistic stat = new ZF_statistic(config, data, mdata._fullmark, temp[0].ToString().Trim());
                stat.partition_process();
                result.Add(stat);

            }

        }
        public void gk_zf_sf_process(MetaData mdata)
        {
            Configuration config = initConfig(mdata._sub, "两类示范校", "高考");
            List<ZF_statistic> result = new List<ZF_statistic>();
            _form.ShowPro(exam_type, 1, mdata.log_name + "数据分析中...");
            for (int i = 0; i < mdata.SF_list.Count; i++)
            {
                string[] SF_code = new string[mdata.SF_list[i].Count - 1];
                for (int j = 1; j < mdata.SF_list[i].Count; j++)
                    SF_code[j - 1] = mdata.SF_list[i][j].ToString().Trim();
                DataTable temp = mdata.basic.filteredtable("xxdm", SF_code);
                ZF_statistic stat = new ZF_statistic(config, temp, mdata._fullmark, mdata.SF_list[i][0].ToString().Trim());
                stat.partition_process();
                result.Add(stat);
            }
            _form.ShowPro(exam_type, 1, mdata.log_name + "报告生成中...");
            ZF_wordcreator create = new ZF_wordcreator(config);
            create.partition_wordcreate(result, "两类示范校");
        }
        public void gk_zf_cj_process(MetaData mdata)
        {
            Configuration config = initConfig(mdata._sub, "城郊", "高考");
            List<ZF_statistic> result = new List<ZF_statistic>();
            _form.ShowPro(exam_type, 1, mdata.log_name + "数据分析中...");
            for (int i = 0; i < mdata.CJ_list.Count; i++)
            {
                string[] cj_code = new string[mdata.CJ_list[i].Count - 1];
                for (int j = 1; j < mdata.CJ_list[i].Count; j++)
                    cj_code[j - 1] = mdata.CJ_list[i][j].ToString().Trim();
                DataTable temp = mdata.basic.filteredtable("qxdm", cj_code);
                ZF_statistic stat = new ZF_statistic(config, temp, mdata._fullmark, mdata.CJ_list[i][0].ToString().Trim());
                stat.partition_process();
                result.Add(stat);
            }
            _form.ShowPro(exam_type, 1, mdata.log_name + "报告生成中...");
            ZF_wordcreator create = new ZF_wordcreator(config);
            create.partition_wordcreate(result, "城郊");
        }
        public void gk_zf_zt_process(MetaData mdata)
        {
            Configuration config = initConfig(mdata._sub, "总体", "高考");
            CacheData cdata = new CacheData();
            _form.ShowPro(exam_type, 1, mdata.log_name + "数据分析中...");
            ZF_statistic stat = new ZF_statistic(config, mdata.basic, mdata._fullmark, "总体");

            stat.partition_process();
            cdata.save_zf_data(mdata._year, mdata._exam, "wk", stat.w_result);
            cdata.save_zf_data(mdata._year, mdata._exam, "lk", stat.l_result);

            _form.ShowPro(exam_type, 1, mdata.log_name + "报告生成中...");
            ZF_wordcreator create = new ZF_wordcreator(config);
            create.total_create(stat);
        }
        public void gk_xx_start()
        {
            exam_type = "gk_xx";
            _exam = "gk";
            foreach (KeyValuePair<string, string> kv in school)
            {
                school_name = kv.Key;
                school_code = kv.Value;
                start();
            }
        }
        public void gk_zh_xx_process(MetaData mdata)
        {
            Configuration config = initConfig(mdata._sub, "学校", "高考");
            List<WSLG_partitiondata> total = new List<WSLG_partitiondata>();
            List<WSLG_partitiondata> single = new List<WSLG_partitiondata>();
            _form.ShowPro("gk_xx", 1, mdata.log_name + "数据分析中...");
            config.WSLG = true;
            PartitionXXDataProcess(config, mdata, total, mdata.zh_basic, mdata.zh_group, mdata._group_num, mdata.zh_ans, mdata.zh_grp, mdata._fullmark);
            PartitionXXDataProcess(config, mdata, single, mdata.basic, mdata.group, mdata._group_num, mdata.ans, mdata.grp, mdata._sub_fullmark);
            
                List<WSLG_partitiondata> t_total = new List<WSLG_partitiondata>();
                List<WSLG_partitiondata> t_single = new List<WSLG_partitiondata>();

                PartitionXX(config, mdata, t_total, mdata.zh_basic, mdata.zh_group, mdata._group_num, school_code, mdata.zh_ans, mdata.zh_grp, mdata._fullmark);
                PartitionXX(config, mdata, t_single, mdata.basic, mdata.group, mdata._group_num, school_code, mdata.ans, mdata.grp, mdata._sub_fullmark);
                t_total.AddRange(total);
                t_single.AddRange(single);
                WordData temp_total = TotalSchoolCal(config, mdata, mdata.zh_basic, mdata.zh_group, mdata._group_num, school_code, mdata.zh_ans, mdata.zh_grp, true, mdata._fullmark, "totalmark");
                WordData temp_single = TotalSchoolCal(config, mdata, mdata.basic, mdata.group, mdata._group_num, school_code, mdata.ans, mdata.grp, false, mdata._sub_fullmark, "ZH_totalmark");
                _form.ShowPro("gk_xx", 1, mdata.log_name + "报告生成中...");
                SchoolWordCreator swc = new SchoolWordCreator(temp_single, t_single, mdata.grp, school_name, mdata.groups_group);
                swc.SetUpZHparam(temp_total, t_total, mdata.zh_grp, mdata.zh_groups_group);
                swc._config = config;
            swc.creating_ZH_word();
                
            
        }
        public void gk_xx_wl_process(MetaData mdata)
        {
            Configuration config = initConfig(mdata._sub, "学校", "高考");
            config.WSLG = true;
            config.school = school_name;
            ArrayList WSLG = new ArrayList();
            _form.ShowPro("gk_xx", 1, mdata.log_name + "文理数据分析中...");
            DataTable XX_data = mdata.basic.filteredtable("xxdm", new string[] { school_code });
            DataTable XX_group = mdata.group.filteredtable("xxdm", new string[] { school_code });

            WSLGCal(config, mdata, XX_data, XX_group, WSLG);
            _form.ShowPro("gk_xx", 1, mdata.log_name + "文理报告生成中...");
            Partition_wordcreator create2 = new Partition_wordcreator(WSLG, mdata.grp, mdata.groups_group);
            create2.SetConfig(config);
            create2.creating_word();
        }
        public void gk_xx_process(MetaData mdata)
        {
            Configuration config = initConfig(mdata._sub, "学校", "高考");
            List<WSLG_partitiondata> total = new List<WSLG_partitiondata>();
            config.WSLG = true;
            _form.ShowPro("gk_xx", 1, mdata.log_name + "数据分析中...");
            PartitionXXDataProcess(config, mdata, total, mdata.basic, mdata.group, mdata._group_num, mdata.ans, mdata.grp, mdata._fullmark);
            List<WSLG_partitiondata> temp_list = new List<WSLG_partitiondata>();
            PartitionXX(config, mdata, temp_list, mdata.basic, mdata.group, mdata._group_num, school_code, mdata.ans, mdata.grp, mdata._fullmark);
            temp_list.AddRange(total);
            WordData temp = TotalSchoolCal(config, mdata, mdata.basic, mdata.group, mdata._group_num, school_code, mdata.ans, mdata.grp, false, mdata._fullmark, "totalmark");
            _form.ShowPro("gk_xx", 1, mdata.log_name + "报告生成中...");
            SchoolWordCreator swc = new SchoolWordCreator(temp, temp_list, mdata.grp, school_name, mdata.groups_group);
            swc._config = config;
            swc.creating_word();
        }
        WordData TotalSchoolCal(Configuration config, MetaData mdata, DataTable data, DataTable group, int groupnum, string school, DataTable my_ans, DataTable my_group, bool isZonghe, decimal my_mark, string totalmark)
        {
            DataTable XX = data.filteredtable("xxdm", new string[] { school });
            DataTable XX_group = group.filteredtable("xxdm", new string[] { school });

                XX = SortTable(XX, totalmark);
                XX_group = SortTable(XX_group, totalmark);

                XX.SeperateGroupsByColumnName(mdata._grouptype, Convert.ToDecimal(groupnum), totalmark);
                XX_group.SeperateGroupsByColumnName(mdata._grouptype, Convert.ToDecimal(groupnum), totalmark);
           
            WordData result = new WordData(mdata.groups_group);
            Total_statistic stat = new Total_statistic(result, XX, my_mark, my_ans, XX_group, my_group, groupnum);
            stat._config = config;
            stat.statistic_process(isZonghe);
            if (mdata.xz.Count > 0)
                stat.xz_postprocess(mdata.xz);

            return result;
        }
        void PartitionXX(Configuration config, MetaData mdata, List<WSLG_partitiondata> result, DataTable data, DataTable group, int groupnum, string school, DataTable my_ans, DataTable my_group, decimal my_mark)
        {
            DataTable XX = data.filteredtable("xxdm", new string[] { school });
            DataTable XX_group = group.filteredtable("xxdm", new string[] { school });
            Partition_statistic XX_stat = new Partition_statistic("本学校", XX, my_mark, my_ans, XX_group, my_group, groupnum);
            XX_stat._config = config;
            XX_stat.statistic_process(false);
            if (mdata.xz.Count > 0)
                XX_stat.xz_postprocess(mdata.xz);
            result.Insert(0, (WSLG_partitiondata)XX_stat.result);
        }
        void PartitionXXDataProcess(Configuration config, MetaData mdata, List<WSLG_partitiondata> result, DataTable data, DataTable group, int groupnum, DataTable my_ans, DataTable my_group, decimal my_mark)
        {

            DataTable QX = data.filteredtable("qxdm", QXTransfer(school_qx[school_code]));
            DataTable QX_group = group.filteredtable("qxdm", QXTransfer(school_qx[school_code]));
            Partition_statistic qx_stat = new Partition_statistic("区整体", QX, my_mark, my_ans, QX_group, my_group, groupnum);
            qx_stat._config = config;
            qx_stat.statistic_process(false);
            if (mdata.xz.Count > 0)
                qx_stat.xz_postprocess(mdata.xz);
            result.Add((WSLG_partitiondata)qx_stat.result);


            Partition_statistic total = new Partition_statistic("市整体", data, my_mark, my_ans, group, my_group, groupnum);
            total._config = config;
            total.statistic_process(false);
            if (mdata.xz.Count > 0)
                total.xz_postprocess(mdata.xz);
            result.Add((WSLG_partitiondata)total.result);

            for (int i = 0; i < mdata.CJ_list.Count; i++)
            {
                ArrayList cj = mdata.CJ_list[i];
                string[] xx_code = new string[cj.Count - 1];
                for (int j = 1; j < cj.Count; j++)
                    xx_code[j - 1] = cj[j].ToString().Trim();
                DataTable temp = data.filteredtable("qxdm", xx_code);
                DataTable temp_group = group.filteredtable("qxdm", xx_code);
                Partition_statistic stat = new Partition_statistic(cj[0].ToString(), temp, my_mark, my_ans, temp_group, my_group, groupnum);
                stat._config = config;
                stat.statistic_process(false);
                if (mdata.xz.Count > 0)
                    stat.xz_postprocess(mdata.xz);
                result.Add((WSLG_partitiondata)stat.result);
            }

            for (int i = 0; i < mdata.SF_list.Count; i++)
            {
                ArrayList sf = mdata.SF_list[i];
                string[] xx_code = new string[sf.Count - 1];
                for (int j = 1; j < sf.Count; j++)
                    xx_code[j - 1] = sf[j].ToString().Trim();
                DataTable temp = data.filteredtable("xxdm", xx_code);
                DataTable temp_group = group.filteredtable("xxdm", xx_code);
                Partition_statistic stat = new Partition_statistic(sf[0].ToString(), temp, my_mark, my_ans, temp_group, my_group, groupnum);
                stat._config = config;
                stat.statistic_process(false);
                if (mdata.xz.Count > 0)
                    stat.xz_postprocess(mdata.xz);
                result.Add((WSLG_partitiondata)stat.result);
            }

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
            config.QX = QX;
            config.fullmark = _fullmark;
            config.sub_fullmark = _sub_fullmark;
            config.is_sub_cor = is_sub_cor;
            config._group_num = _group_num;
            config._grouptype = _grouptype;
            config.year = curryear;
            config.month = currmonth;
            return config;
        }
        public void gk_qx_start()
        {
            exam_type = "gk_qx";
            _exam = "gk";
            start();
        }
        
        public void gk_zh_qx_process(MetaData mdata)
        {
            Configuration config = initConfig(mdata._sub, "区县", "高考");
            ArrayList total = new ArrayList();
            ArrayList QX = new ArrayList();
            ArrayList ZH_total = new ArrayList();
            ArrayList ZH_QX = new ArrayList();
            _form.ShowPro("gk_qx", 1, mdata.log_name + "数据分析中...");
            CalculatePartition(config, ZH_total, "市整体", mdata.zh_basic, mdata.zh_group, mdata._fullmark, mdata.zh_grp, mdata._group_num, true, mdata.zh_ans);
            //decimal ZH_fullmark = (decimal)((PartitionData)ZH_total[0]).groups_analysis.Rows.Find(sub)["fullmark"];
            CalculatePartition(config, total, "市整体", mdata.basic, mdata.group, mdata._sub_fullmark, mdata.grp, mdata._group_num, false, mdata.ans);
            for (int i = 0; i < mdata.SF_list.Count; i++)
            {
                string[] SF_code = new string[mdata.SF_list[i].Count - 1];
                for (int j = 1; j < mdata.SF_list[i].Count; j++)
                    SF_code[j - 1] = mdata.SF_list[i][j].ToString().Trim();
                DataTable temp = mdata.zh_basic.filteredtable("xxdm", SF_code);
                DataTable temp_group = mdata.zh_group.filteredtable("xxdm", SF_code);

                DataTable single = mdata.basic.filteredtable("xxdm", SF_code);
                DataTable single_table = mdata.group.filteredtable("xxdm", SF_code);
                CalculatePartition(config, ZH_total, mdata.SF_list[i][0].ToString(), temp, temp_group, mdata._fullmark, mdata.zh_grp, mdata._group_num, true, mdata.zh_ans);
                CalculatePartition(config, total, mdata.SF_list[i][0].ToString(), single, single_table, mdata._sub_fullmark, mdata.grp, mdata._group_num, false, mdata.ans);
            }
            for (int i = 0; i < mdata.CJ_list.Count; i++)
            {
                string[] SF_code = new string[mdata.CJ_list[i].Count - 1];
                for (int j = 1; j < mdata.CJ_list[i].Count; j++)
                    SF_code[j - 1] = mdata.CJ_list[i][j].ToString().Trim();
                DataTable temp = mdata.zh_basic.filteredtable("qxdm", SF_code);
                DataTable temp_group = mdata.zh_group.filteredtable("qxdm", SF_code);

                DataTable single = mdata.basic.filteredtable("qxdm", SF_code);
                DataTable single_table = mdata.group.filteredtable("qxdm", SF_code);
                CalculatePartition(config, ZH_total, mdata.CJ_list[i][0].ToString(), temp, temp_group, mdata._fullmark, mdata.zh_grp, mdata._group_num, true, mdata.zh_ans);
                CalculatePartition(config, total, mdata.CJ_list[i][0].ToString(), single, single_table, mdata._sub_fullmark, mdata.grp, mdata._group_num, false, mdata.ans);
            }
            DataTable QX_ZH_data = mdata.zh_basic.filteredtable("qxdm", QXTransfer(qx_code));
            DataTable QX_ZH_group = mdata.zh_group.filteredtable("qxdm", QXTransfer(qx_code));

            DataTable QX_data = mdata.basic.filteredtable("qxdm", QXTransfer(qx_code));
            DataTable QX_group = mdata.group.filteredtable("qxdm", QXTransfer(qx_code));

            CalculatePartition(config, ZH_total, "区整体", QX_ZH_data, QX_ZH_group, mdata._fullmark, mdata.zh_grp, mdata._group_num, true, mdata.zh_ans);
            CalculatePartition(config, total, "区整体", QX_data, QX_group, mdata._sub_fullmark, mdata.grp, mdata._group_num, false, mdata.ans);

            string[] qxsf_code = CalculateTotal(mdata.QXSF_list);
            DataTable qxsf_zh_data = QX_ZH_data.filteredtable("xxdm", qxsf_code);
            DataTable qxsf_zh_group = QX_ZH_group.filteredtable("xxdm", qxsf_code);
            DataTable qxsf_data = SortTable(QX_data.filteredtable("xxdm", qxsf_code), "ZH_totalmark");
            DataTable qxsf_group = SortTable(QX_group.filteredtable("xxdm", qxsf_code), "ZH_totalmark");

            qxsf_zh_data.SeperateGroups(mdata._grouptype, mdata._group_num, "groups");
            qxsf_zh_group.SeperateGroups(mdata._grouptype, mdata._group_num, "groups");

            
            qxsf_data.SeperateGroupsByColumnName(mdata._grouptype, mdata._group_num, "ZH_totalmark");
            qxsf_group.SeperateGroupsByColumnName(mdata._grouptype, mdata._group_num, "ZH_totalmark");

            CalculatePartition(config, ZH_total, "分类整体", qxsf_zh_data, qxsf_zh_group, mdata._fullmark, mdata.zh_grp, mdata._group_num, true, mdata.zh_ans);
            CalculatePartition(config, total, "分类整体", qxsf_data, qxsf_group, mdata._sub_fullmark, mdata.grp, mdata._group_num, false, mdata.ans);
            for (int i = 0; i < mdata.QXSF_list.Count; i++)
            {
                string[] SF_code = new string[mdata.QXSF_list[i].Count - 1];
                for (int j = 1; j < mdata.QXSF_list[i].Count; j++)
                    SF_code[j - 1] = mdata.QXSF_list[i][j].ToString().Trim();
                DataTable temp = qxsf_zh_data.filteredtable("xxdm", SF_code);
                DataTable temp_group = qxsf_zh_group.filteredtable("xxdm", SF_code);

                DataTable single = qxsf_data.filteredtable("xxdm", SF_code);
                DataTable single_table = qxsf_group.filteredtable("xxdm", SF_code);
                CalculatePartition(config, ZH_total, mdata.QXSF_list[i][0].ToString(), temp, temp_group, mdata._fullmark, mdata.zh_grp, mdata._group_num, true, mdata.zh_ans);
                CalculatePartition(config, total, mdata.QXSF_list[i][0].ToString(), single, single_table, mdata._sub_fullmark, mdata.grp, mdata._group_num, false, mdata.ans);
                CalculatePartition(config, ZH_QX, mdata.QXSF_list[i][0].ToString(), temp, temp_group, mdata._fullmark, mdata.zh_grp, mdata._group_num, true, mdata.zh_ans);
                CalculatePartition(config, QX, mdata.QXSF_list[i][0].ToString(), single, single_table, mdata._sub_fullmark, mdata.grp, mdata._group_num, false, mdata.ans);
            }
            CalculatePartition(config, ZH_QX, "分类整体", qxsf_zh_data, qxsf_zh_group, mdata._fullmark, mdata.zh_grp, mdata._group_num, true, mdata.zh_ans);
            CalculatePartition(config, QX, "分类整体", qxsf_data, qxsf_group, mdata._sub_fullmark, mdata.grp, mdata._group_num, false, mdata.ans);
            _form.ShowPro("gk_qx", 1, mdata.log_name + "报告生成中...");
            Partition_wordcreator create = new Partition_wordcreator(total, QX, mdata.grp, mdata.groups_group);
            create.SetConfig(config);
            create.creating_ZH_QX_word(ZH_total, ZH_QX, mdata.zh_grp, mdata.zh_groups_group);
        }
        void CalculatePartition(Configuration config, ArrayList list, String title, DataTable total, DataTable group, decimal fullmark, DataTable group_ans, int groupnum, bool isZonghe, DataTable ans)
        {
            Partition_statistic stat = new Partition_statistic(title, total, fullmark, ans, group, group_ans, groupnum);
            stat._config = config;
            stat.statistic_process(isZonghe);
            list.Add(stat.result);
        }
        public void gk_qx_wl_process(MetaData mdata)
        {
            _form.ShowPro("gk_qx", 1, mdata.log_name + "文理数据分析中...");
            Configuration config = initConfig(mdata._sub, "区县", "高考");
            config.WSLG = true;
            ArrayList WSLG = new ArrayList();
            DataTable QX_data = mdata.basic.filteredtable("qxdm", QXTransfer(qx_code));
            DataTable QX_group = mdata.group.filteredtable("qxdm", QXTransfer(qx_code));

            WSLGCal(config, mdata, QX_data, QX_group, WSLG);
            _form.ShowPro("gk_qx", 1, mdata.log_name + "文理报告生成中...");
            Partition_wordcreator create2 = new Partition_wordcreator(WSLG, mdata.grp, mdata.groups_group);
            create2.SetConfig(config);
            create2.creating_word();
        }
        void WSLGCal(Configuration config, MetaData mdata, DataTable QX_data, DataTable QX_group, ArrayList WSLG)
        {

            int group = QX_data.SeperateGroups(mdata._grouptype, mdata._group_num, "groups");
            QX_group.SeperateGroups(mdata._grouptype, mdata._group_num, "groups");
            //if (QX_data.Columns.Contains("XZ"))
            //    XZ_group_separate(QX_data, mdata);
            DataTable W_data = QX_data.Likefilter("zkzh", "'1*'");
            DataTable W_group = QX_group.Likefilter("zkzh", "'1*'");

            Partition_statistic w_stat = new Partition_statistic("文科", W_data, mdata._fullmark, mdata.ans, W_group, mdata.grp, group);
            w_stat._config = config;
            w_stat.statistic_process(false);
            if (mdata.xz.Count > 0)
                w_stat.xz_postprocess(mdata.xz);
            WSLG.Add(w_stat.result);

            DataTable l_data = QX_data.Likefilter("zkzh", "'5*'");
            DataTable l_group = QX_group.Likefilter("zkzh", "'5*'");

            Partition_statistic l_stat = new Partition_statistic("理科", l_data, mdata._fullmark, mdata.ans, l_group, mdata.grp, group);
            l_stat._config = config;
            l_stat.statistic_process(false);
            if (mdata.xz.Count > 0)
                l_stat.xz_postprocess(mdata.xz);
            WSLG.Add(l_stat.result);

            Partition_statistic total_stat = new Partition_statistic("分类整体", QX_data, mdata._fullmark, mdata.ans, QX_group, mdata.grp, group);
            total_stat._config = config;
            total_stat.statistic_process(false);
            if (mdata.xz.Count > 0)
                total_stat.xz_postprocess(mdata.xz);
            WSLG.Add(total_stat.result);


        }
        public void gk_qx_process(MetaData mdata)
        {
            Configuration config = initConfig(mdata._sub, "区县", "高考");
            ArrayList QX = new ArrayList();
            ArrayList total = new ArrayList();
            _form.ShowPro("gk_qx", 1, mdata.log_name + "数据分析中...");
            PartitionQXDataProcess(config, mdata, total, QX, mdata.basic, mdata.group, mdata._group_num);
            _form.ShowPro("gk_qx", 1, mdata.log_name + "报告生成中...");
            Partition_wordcreator create = new Partition_wordcreator(total, QX, mdata.grp, mdata.groups_group);
            create.SetConfig(config);
            create.creating_word();


        }
        void PartitionQXDataProcess(Configuration config, MetaData mdata, ArrayList result, ArrayList sresult, DataTable data, DataTable group, int groupnum)
        {
            Partition_statistic total = new Partition_statistic("市整体", data, mdata._fullmark, mdata.ans, group, mdata.grp, groupnum);
            total._config = config;
            total.statistic_process(false);
            if (mdata.xz.Count > 0)
                total.xz_postprocess(mdata.xz);
            result.Add(total.result);

            for (int i = 0; i < mdata.SF_list.Count; i++)
            {
                ArrayList sf = mdata.SF_list[i];
                string[] xx_code = new string[sf.Count - 1];
                for (int j = 1; j < sf.Count; j++)
                    xx_code[j - 1] = sf[j].ToString().Trim();
                DataTable temp = data.filteredtable("xxdm", xx_code);
                DataTable temp_group = group.filteredtable("xxdm", xx_code);
                Partition_statistic stat = new Partition_statistic(sf[0].ToString(), temp, mdata._fullmark, mdata.ans, temp_group, mdata.grp, groupnum);
                stat._config = config;
                stat.statistic_process(false);
                if (mdata.xz.Count > 0)
                    stat.xz_postprocess(mdata.xz);
                result.Add(stat.result);
            }

            for (int i = 0; i < mdata.CJ_list.Count; i++)
            {
                ArrayList cj = mdata.CJ_list[i];
                string[] xx_code = new string[cj.Count - 1];
                for (int j = 1; j < cj.Count; j++)
                    xx_code[j - 1] = cj[j].ToString().Trim();
                DataTable temp = data.filteredtable("qxdm", xx_code);
                DataTable temp_group = group.filteredtable("qxdm", xx_code);
                Partition_statistic stat = new Partition_statistic(cj[0].ToString(), temp, mdata._fullmark, mdata.ans, temp_group, mdata.grp, groupnum);
                stat._config = config;
                stat.statistic_process(false);
                if (mdata.xz.Count > 0)
                    stat.xz_postprocess(mdata.xz);
                result.Add(stat.result);
            }

            DataTable QX = data.filteredtable("qxdm", QXTransfer(qx_code));
            DataTable QX_group = group.filteredtable("qxdm", QXTransfer(qx_code));
            Partition_statistic qx_stat = new Partition_statistic("区整体", QX, mdata._fullmark, mdata.ans, QX_group, mdata.grp, groupnum);
            qx_stat._config = config;
            qx_stat.statistic_process(false);
            if (mdata.xz.Count > 0)
                qx_stat.xz_postprocess(mdata.xz);
            result.Add(qx_stat.result);
            if (!Utils.OnlyQZT)
            {
                PartitionDataProcess(config, mdata, result, mdata.QXSF_list, "xxdm", QX, QX_group, groupnum, true);
                PartitionDataProcess(config, mdata, sresult, mdata.QXSF_list, "xxdm", QX, QX_group, groupnum, false);
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
            PartitionDataProcess(config, mdata, list, mdata.CJ_list, "xxdm", mdata.basic, mdata.group, mdata._group_num, false);
            _form.ShowPro("gk_sf", 1, mdata.log_name + "报告生成中...");
            Partition_wordcreator create = new Partition_wordcreator(list, mdata.grp, mdata.groups_group);
            create.SetConfig(config);
            create.creating_word();
        }
        public void gk_zh_sf_process(MetaData mdata)
        {
            Configuration config = initConfig(mdata._sub, "两类示范校", "高考");
            ArrayList sdata = new ArrayList();
            ArrayList ZH_data = new ArrayList();
            _form.ShowPro("gk_sf", 1, mdata.log_name + "数据分析中...");
            string[] total_code = CalculateTotal(mdata.SF_list);
            DataTable total = mdata.zh_basic.filteredtable("xxdm", total_code);
            DataTable total_group = mdata.zh_group.filteredtable("xxdm", total_code);

            int groupnum = total.SeperateGroups(mdata._grouptype, mdata._group_num, "groups");
            total_group.SeperateGroups(mdata._grouptype, mdata._group_num, "groups");

            DataTable single_total = SortTable(mdata.basic.filteredtable("xxdm", total_code), "ZH_totalmark");
            DataTable single_total_group = SortTable(mdata.group.filteredtable("xxdm", total_code), "ZH_totalmark");


            single_total.SeperateGroupsByColumnName(mdata._grouptype, mdata._group_num, "ZH_totalmark");
            single_total_group.SeperateGroupsByColumnName(mdata._grouptype, mdata._group_num, "ZH_totalmark");
            for (int i = 0; i < mdata.SF_list.Count; i++)
            {
                string[] SF_code = new string[mdata.SF_list[i].Count - 1];
                for (int j = 1; j < mdata.SF_list[i].Count; j++)
                    SF_code[j - 1] = mdata.SF_list[i][j].ToString().Trim();
                DataTable temp = total.filteredtable("xxdm", SF_code);
                DataTable temp_group = total_group.filteredtable("xxdm", SF_code);
                Partition_statistic stat = new Partition_statistic(mdata.SF_list[i][0].ToString().Trim(), temp, mdata._fullmark, mdata.zh_ans, temp_group, mdata.zh_grp, groupnum);
                stat._config = config;
                stat.statistic_process(true);
                ZH_data.Add(stat.result);

                DataTable single = single_total.filteredtable("xxdm", SF_code);
                DataTable single_group = single_total_group.filteredtable("xxdm", SF_code);
                Partition_statistic single_stat = new Partition_statistic(mdata.SF_list[i][0].ToString().Trim(), single, mdata._sub_fullmark, mdata.ans, single_group, mdata.grp, groupnum);
                single_stat._config = config;
                single_stat.statistic_process(false);
                sdata.Add(single_stat.result);
            }

            Partition_statistic total_stat = new Partition_statistic("分类整体", total, mdata._fullmark, mdata.zh_ans, total_group, mdata.zh_grp, groupnum);
            total_stat._config = config;
            total_stat.statistic_process(true);
            ZH_data.Add(total_stat.result);
            Partition_statistic single_total_stat = new Partition_statistic("分类整体", single_total, mdata._sub_fullmark, mdata.ans, single_total_group, mdata.grp, groupnum);
            single_total_stat._config = config;
            single_total_stat.statistic_process(false);
            sdata.Add(single_total_stat.result);
            _form.ShowPro("gk_sf", 1, mdata.log_name + "报告生成中...");
            Partition_wordcreator create = new Partition_wordcreator(sdata, mdata.grp, mdata.groups_group);
            create.SetConfig(config);
            create.creating_ZH_word(ZH_data, mdata.zh_grp, mdata.zh_groups_group);
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
            PartitionDataProcess(config, mdata, list, mdata.CJ_list, "qxdm", mdata.basic, mdata.group, mdata._group_num, false);
            _form.ShowPro("gk_cj", 1, mdata.log_name + "报告生成中...");
            Partition_wordcreator create = new Partition_wordcreator(list, mdata.grp, mdata.groups_group);
            create.SetConfig(config);
            create.creating_word();
        }
        public void gk_zh_cj_process(MetaData mdata)
        {
            Configuration config = initConfig(mdata._sub, "城郊", "高考");
            ArrayList sdata = new ArrayList();
            ArrayList ZH_data = new ArrayList();
            _form.ShowPro("gk_cj", 1, mdata.log_name + "数据分析中...");
            string[] total_code = CalculateTotal(mdata.CJ_list);

            DataTable total = mdata.zh_basic.filteredtable("qxdm", total_code);
            DataTable total_group = mdata.zh_group.filteredtable("qxdm", total_code);

            int groupnum = total.SeperateGroups(mdata._grouptype, mdata._group_num, "groups");
            total_group.SeperateGroups(mdata._grouptype, mdata._group_num, "groups");

            DataTable single_total = SortTable(mdata.basic.filteredtable("qxdm", total_code), "ZH_totalmark");
            DataTable single_total_group = SortTable(mdata.group.filteredtable("qxdm", total_code), "ZH_totalmark");;

            
            single_total.SeperateGroupsByColumnName(mdata._grouptype, mdata._group_num, "ZH_totalmark");
            single_total_group.SeperateGroupsByColumnName(mdata._grouptype, mdata._group_num, "ZH_totalmark");
            for (int i = 0; i < mdata.CJ_list.Count; i++)
            {
                string[] SF_code = new string[mdata.CJ_list[i].Count - 1];
                for (int j = 1; j < mdata.CJ_list[i].Count; j++)
                    SF_code[j - 1] = mdata.CJ_list[i][j].ToString().Trim();
                DataTable temp = total.filteredtable("qxdm", SF_code);
                DataTable temp_group = total_group.filteredtable("qxdm", SF_code);
                Partition_statistic stat = new Partition_statistic(mdata.CJ_list[i][0].ToString().Trim(), temp, mdata._fullmark, mdata.zh_ans, temp_group, mdata.zh_grp, groupnum);
                stat._config = config;
                stat.statistic_process(true);
                ZH_data.Add(stat.result);

                DataTable single = single_total.filteredtable("qxdm", SF_code);
                DataTable single_group = single_total_group.filteredtable("qxdm", SF_code);
                Partition_statistic single_stat = new Partition_statistic(mdata.CJ_list[i][0].ToString().Trim(), single, mdata._sub_fullmark, mdata.ans, single_group, mdata.grp, groupnum);
                single_stat._config = config;
                single_stat.statistic_process(false);
                sdata.Add(single_stat.result);
            }

            Partition_statistic total_stat = new Partition_statistic("分类整体", total, mdata._fullmark, mdata.zh_ans, total_group, mdata.zh_grp, groupnum);
            total_stat._config = config;
            total_stat.statistic_process(true);
            ZH_data.Add(total_stat.result);
            Partition_statistic single_total_stat = new Partition_statistic("分类整体", single_total, mdata._sub_fullmark, mdata.ans, single_total_group, mdata.grp, groupnum);
            single_total_stat._config = config;
            single_total_stat.statistic_process(false);
            sdata.Add(single_total_stat.result);
            _form.ShowPro("gk_cj", 1, mdata.log_name + "报告生成中...");
            Partition_wordcreator create = new Partition_wordcreator(sdata, mdata.grp, mdata.groups_group);
            create.SetConfig(config);
            create.creating_ZH_word(ZH_data, mdata.zh_grp, mdata.zh_groups_group);
        }
        void PartitionDataProcess(Configuration config, MetaData mdata, ArrayList result, List<ArrayList> list, String filter, DataTable data, DataTable group, int groupnum, bool isQXSF)
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
            //if (dt.Columns.Contains("XZ"))
            //    XZ_group_separate(dt, mdata);
            Partition_statistic total = new Partition_statistic("分类整体", dt, mdata._fullmark, mdata.ans, dt_group, mdata.grp, groupnum);
            total._config = config;
            total.statistic_process(false);
            if (mdata.xz.Count > 0)
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
                stat._config = config;
                stat.statistic_process(false);
                if (mdata.xz.Count > 0)
                    stat.xz_postprocess(mdata.xz);
                result.Add(stat.result);
            }
            if (!isQXSF)
                result.Add(total.result);
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
        public void gk_schedule(MetaData mdata, gk_process zt_process, gk_process wl_process)
        {
            if (mdata._sub.Equals("语文") || mdata._sub.Equals("英语"))
            {
                switch (mdata.ywyy_choice)
                {
                    case "类型报告":
                        zt_process(mdata);
                        break;
                    case "文理报告":
                        wl_process(mdata);
                        break;
                    case "两者均有":
                        zt_process(mdata);
                        wl_process(mdata);
                        break;
                    default:
                        break;

                }
            }
            else
                zt_process(mdata);
        }
        public void gk_zt_process(MetaData mdata)
        {
            Configuration config = initConfig(mdata._sub, "总体", "高考");
            CacheData cdata = new CacheData();

            WordData data = new WordData(mdata.groups_group);
            _form.ShowPro("gk_zt", 1, mdata.log_name + "数据分析中...");
            
            Total_statistic stat = new Total_statistic(data, mdata.basic, mdata._fullmark, mdata.ans, mdata.group, mdata.grp, mdata._group_num);
            stat._config = config;
            stat.statistic_process(false);
            if (mdata.xz.Count > 0)
                stat.xz_postprocess(mdata.xz);

            cdata.save_totaldata(mdata._year, mdata._exam, mdata.get_sub(), data);

            _form.ShowPro("gk_zt", 1, mdata.log_name + "报告生成中...");
            WordCreator create = new WordCreator(data, config);
            create.creating_word();
        }
        public void gk_zt_wl_process(MetaData mdata)
        {
            Configuration config = initConfig(mdata._sub, "总体", "高考");
            config.WSLG = true;
            ArrayList WSLG = new ArrayList();

            CacheData cdata = new CacheData();
            string subname = mdata.get_sub();

            _form.ShowPro("gk_zt", 1, mdata.log_name + "文理数据分析中...");

            DataTable W_data = mdata.basic.Likefilter("zkzh", "'1*'");
            DataTable W_group = mdata.group.Likefilter("zkzh", "'1*'");

            Partition_statistic w_stat = new Partition_statistic("文科", W_data, mdata._fullmark, mdata.ans, W_group, mdata.grp, mdata._group_num);
            w_stat._config = config;
            w_stat.statistic_process(false);
            if (mdata.xz.Count > 0)
                w_stat.xz_postprocess(mdata.xz);
            WSLG.Add(w_stat.result);

            cdata.save_partitiondata(mdata._year, mdata._exam, subname + "w", w_stat.result);

            DataTable l_data = mdata.basic.Likefilter("zkzh", "'5*'");
            DataTable l_group = mdata.group.Likefilter("zkzh", "'5*'");

            Partition_statistic l_stat = new Partition_statistic("理科", l_data, mdata._fullmark, mdata.ans, l_group, mdata.grp, mdata._group_num);
            l_stat._config = config;
            l_stat.statistic_process(false);
            if (mdata.xz.Count > 0)
                l_stat.xz_postprocess(mdata.xz);
            WSLG.Add(l_stat.result);

            cdata.save_partitiondata(mdata._year, mdata._exam, subname + "l", l_stat.result);

            Partition_statistic total_stat = new Partition_statistic("分类整体", mdata.basic, mdata._fullmark, mdata.ans, mdata.group, mdata.grp, mdata._group_num);
            total_stat._config = config;
            total_stat.statistic_process(false);
            if (mdata.xz.Count > 0)
                total_stat.xz_postprocess(mdata.xz);
            WSLG.Add(total_stat.result);
            _form.ShowPro("gk_zt", 1, mdata.log_name + "文理报告生成中...");
            Partition_wordcreator create2 = new Partition_wordcreator(WSLG, mdata.grp, mdata.groups_group);
            create2.SetConfig(config);
            create2.creating_word();
        }
        public void gk_zh_zt_process(MetaData mdata)
        {
            Configuration config = initConfig(mdata._sub, "总体", "高考");
            _form.ShowPro("gk_zt", 1, mdata.log_name + "数据分析中...");
            CacheData cdata = new CacheData();
            WordData total = new WordData(mdata.zh_groups_group);
            Total_statistic total_stat = new Total_statistic(total, mdata.zh_basic, mdata._fullmark, mdata.zh_ans, mdata.zh_group, mdata.zh_grp, mdata._group_num);
            total_stat._config = config;
            total_stat.statistic_process(true);

            cdata.save_totaldata(mdata._year, mdata._exam, mdata.get_zh(), total);

            WordData single = new WordData(mdata.groups_group);

            Total_statistic single_stat = new Total_statistic(single, mdata.basic, mdata._sub_fullmark, mdata.ans, mdata.group, mdata.grp, mdata._group_num);
            single_stat._config = config;
            single_stat.statistic_process(false);

            cdata.save_totaldata(mdata._year, mdata._exam, mdata.get_sub(), single);

            _form.ShowPro("gk_zt", 1, mdata.log_name + "报告生成中...");
            WordCreator create = new WordCreator(single, total, config);
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
            if (mdata.xz.Count > 0)
                stat.xz_postprocess(mdata.xz);
            _form.ShowPro("zk_zt", 1, mdata.log_name + "报告生成中...");
            WordCreator creator = new WordCreator(result, config);
            creator.creating_word();
        }
        
        
        public void zk_qx_start()
        {
            _exam = "zk";
            exam_type = "zk_qx";
            start();
        }



        public void zk_qx_process(MetaData mdata)
        {
            Configuration config = initConfig(mdata._sub, "区县", "中考");
            ArrayList sdata = new ArrayList();
            ArrayList totaldata = new ArrayList();
            _form.ShowPro(exam_type, 1, mdata.log_name + "数据分析中...");
            Partition_statistic total = new Partition_statistic("市整体", mdata.basic, mdata._fullmark, mdata.ans, mdata.group, mdata.grp, mdata._group_num);
            total._config = config;
            total.statistic_process(false);
            if (mdata.xz.Count > 0)
                total.xz_postprocess(mdata.xz);
            totaldata.Add(total.result);

            for (int mark = 0; mark < mdata.CJ_list.Count; mark++)
            {
                string[] CQ_code = new string[mdata.CJ_list[mark].Count - 1];

                for (int i = 1; i < mdata.CJ_list[mark].Count; i++)
                {
                    CQ_code[i - 1] = mdata.CJ_list[mark][i].ToString().Trim();
                }
                DataTable CQ_data = mdata.basic.filteredtable("qxdm", CQ_code);
                DataTable CQ_groups_data = mdata.group.filteredtable("qxdm", CQ_code);

                Partition_statistic CQ = new Partition_statistic(mdata.CJ_list[mark][0].ToString().Trim(), CQ_data, mdata._fullmark, mdata.ans, CQ_groups_data, mdata.grp, mdata._group_num);
                CQ._config = config;
                CQ.statistic_process(false);
                if (mdata.xz.Count > 0)
                    CQ.xz_postprocess(mdata.xz);
                totaldata.Add(CQ.result);
            }

            DataTable QX_total_data = mdata.basic.filteredtable("qxdm", QXTransfer(qx_code));
            DataTable QX_groups_data = mdata.group.filteredtable("qxdm", QXTransfer(qx_code));

            Partition_statistic QX_total = new Partition_statistic("区整体", QX_total_data, mdata._fullmark, mdata.ans, QX_groups_data, mdata.grp, mdata._group_num);
            QX_total._config = config;
            QX_total.statistic_process(false);
            if (mdata.xz.Count > 0)
                QX_total.xz_postprocess(mdata.xz);
            totaldata.Add(QX_total.result);

            CalculateClassTotal(config, QX_total_data, QX_groups_data, totaldata, sdata, mdata);
            _form.ShowPro(exam_type, 1, mdata.log_name + "报告生成中...");
            Partition_wordcreator create = new Partition_wordcreator(totaldata, sdata, mdata.grp, mdata.groups_group);
            create.SetConfig(config);
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

        void CalculateClassTotal(Configuration config, DataTable total, DataTable groups_data, ArrayList totaldata, ArrayList sdata, MetaData mdata)
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
            DataTable ClassTotal_data = total.filteredtable("xxdm", SF_code);
            DataTable ClassGroupTotal_data = groups_data.filteredtable("xxdm", SF_code);

            int groupnum = ClassTotal_data.SeperateGroups(mdata._grouptype, mdata._group_num, "groups");
            ClassGroupTotal_data.SeperateGroups(mdata._grouptype, mdata._group_num, "groups");
            //if (ClassTotal_data.Columns.Contains("XZ"))
            //    XZ_group_separate(ClassTotal_data, mdata);
            Partition_statistic ClassTotal = new Partition_statistic("分类整体", ClassTotal_data, mdata._fullmark, mdata.ans, ClassGroupTotal_data, mdata.grp, mdata._group_num);
            ClassTotal._config = config;
            ClassTotal.statistic_process(false);
            if (mdata.xz.Count > 0)
                ClassTotal.xz_postprocess(mdata.xz);
            totaldata.Add(ClassTotal.result);

            for (int i = 0; i < mdata.QXSF_list.Count; i++)
            {
                ArrayList temp = mdata.QXSF_list[i];
                string[] xx_code = new string[temp.Count - 1];
                for (int j = 1; j < temp.Count; j++)
                    xx_code[j - 1] = temp[j].ToString().Trim();
                DataTable xx_data = ClassTotal_data.filteredtable("xxdm", xx_code);
                DataTable xx_group_data = ClassGroupTotal_data.filteredtable("xxdm", xx_code);

                Partition_statistic XXTotal = new Partition_statistic(temp[0].ToString().Trim(), xx_data, mdata._fullmark, mdata.ans, xx_group_data, mdata.grp, mdata._group_num);
                XXTotal._config = config;
                XXTotal.statistic_process(false);
                if (mdata.xz.Count > 0)
                    XXTotal.xz_postprocess(mdata.xz);
                totaldata.Add(XXTotal.result);
                sdata.Add(XXTotal.result);

            }
            sdata.Add(ClassTotal.result);
        }
        public DataTable SortTable(DataTable dt, string ColumnName)
        {
            DataView dv = dt.DefaultView;
            dv.Sort = ColumnName;
            DataTable res = dv.ToTable();
            return res;
        }
        //void XZ_group_separate(DataTable temp_dt, MetaData mdata)
        //{
        //    if (!temp_dt.Columns.Contains("xz_groups"))
        //        temp_dt.Columns.Add("xz_groups", typeof(string));
        //    var xz_tuple = from row in temp_dt.AsEnumerable()
        //                   group row by row.Field<string>("XZ") into grp
        //                   select new
        //                   {
        //                       name = grp.Key
        //                   };
        //    foreach (var item in xz_tuple)
        //    {
        //        DataView dv = temp_dt.equalfilter("XZ", item.name).DefaultView;
        //        DataTable inter_table = dv.ToTable();
        //        inter_table.SeperateGroups(mdata._grouptype, mdata._group_num, "xz_groups");
        //        var temp = from row in temp_dt.AsEnumerable()
        //                   join row2 in inter_table.AsEnumerable() on row.Field<string>("kh") equals row2.Field<string>("kh")
        //                   where row.Field<string>("XZ") == item.name
        //                   select new
        //                   {
        //                       row1 = row,
        //                       groups = row2.Field<string>("xz_groups")
        //                   };
        //        foreach (var inner_item in temp)
        //        {
        //            inner_item.row1.SetField<string>("xz_groups", inner_item.groups);
        //        }
        //    }
        //}
    }
}
