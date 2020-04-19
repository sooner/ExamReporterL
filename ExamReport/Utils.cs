using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.OleDb;
using System.Configuration;
using System.Threading;
using MySql.Data;
using MySql.Data.MySqlClient;
using Ionic.Zip;

namespace ExamReport
{
    public class Utils
    {
        public enum UnionType { QX_XX, ID };
        public static string save_address;
        public static string exam;
        public static string subject;
        public static string report_style = "";
        public static string template_address;
        public static string zh_template_address;
        public static string CurrentDirectory = System.IO.Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath);
        public static bool isVisible = false;
        public static bool saveMidData = false;
        public static string QX;
        public static bool WSLG = false;
        public static bool sub_iszero = false;
        public static bool fullmark_iszero = false;
        public static decimal PartialRight = 0;
        public static string year;
        public static string month;
        public static ZK_database.GroupType group_type = ZK_database.GroupType.population;
        public static string school_name;
        public static bool OnlyQZT = false;

        public static int smooth_degree = 10;

        public static decimal fullmark = 0;

        public static List<decimal> GroupMark = new List<decimal>();

        public static decimal shengwu_zhengzhi;
        public static decimal wuli_lishi;
        public static decimal huaxue_dili;

        public static string[] ywyy_combo = new string[] { "文理报告", "类型报告", "两者均有"};
        public static string[] zh_combo = new string[] { "总体总分相关", "科目总分相关" };
        public static string[] null_combo = new string[] { };

        public static Mutex mutex_clipboard = new Mutex();

        public static string ZK_title_1 = "北京市高级中等学校招生考试";
        public static string ZK_title_2 = "实测数据统计分析报告";
        public static string ZK_QX_title_2 = "分类校数据统计分析报告";
        public static string HK_title_1 = "北京市高中会考数据统计分析报告";
        public static string GK_title_1 = "北京市普通高等学校招生全国统一考试";
        public static string GK_CJ_title_2 = "城区、郊区数据统计分析报告";
        public static string GK_SF_title_2 = "示范校数据统计分析报告";
        public static string GK_QX_title_2 = "分类校数据统计分析报告";
        public static string GK_CUS_title_2 = "自选数据统计分析报告";
        public static string GK_title_2 = "实测数据统计分析报告";
        public static string GK_ZF_title_1 = "年北京市普通高考";
        public static string GK_ZF_title_2 = "试卷总分统计分析报告";
        public static string GK_ZF_title_xz_1 = "年高考数据分析报告";
        public static string GK_ZF_title_xz_2 = "(行政版)";
        public static string ZK_ZF_title_xz_1 = "年中考数据分析报告";
        public static string ZK_ZF_title_xz_2 = "（行政版）";
        public static string GK_WSLG_title_2 = "文史、理工类数据统计分析报告";
        public static string XX_title = "学校数据统计分析报告";

        public static string[] qx_in_order = { "东城区", "西城区", "海淀区", "朝阳区", "石景山区", "丰台区", "燕 山", "通州区", "顺义区", "昌平区", "门头沟区", "房山区", "大兴区", "怀柔区", "平谷区", "密云县", "延庆县" };
        public static string[] qxdm_in_order = { "01", "02","08","05","07","06", "10","12", "13","14","09","11","15","16", "17", "18","19"};

        public static string[] hk_subject = { "yw", "sx", "yy", "wl", "hx", "sw", "ls", "dl" };
        public static Dictionary<int, string> sub_choice = new Dictionary<int,string>() {
        {1, "物理+生化+地理"}, 
        {2, "物理+生化+历史"}, 
        {3, "物理+生化+思品"}, 
        {4, "物理+地理+历史"}, 
        {5, "物理+地理+思品"},
        {6, "物理+历史+思品"}, 
        {7, "生化+地理+历史"}, 
        {8, "生化+地理+思品"}, 
        {9, "生化+历史+思品"}};
        public static void WSLG_WriteFrontPage(Configuration config, Microsoft.Office.Interop.Word._Document oDoc)
        {
            WriteIntoDocument(oDoc, "date", config.year + "年" + config.month);
            WriteIntoDocument(oDoc, "title_1", GK_title_1);
            WriteIntoDocument(oDoc, "title_2", GK_WSLG_title_2);
            if (config.report_style.Equals("区县"))
            {
                WriteIntoDocument(oDoc, "QX", config.QX);
                WriteIntoDocument(oDoc, "QX_subject", config.subject);
            }
            else if (config.report_style.Equals("总体") || config.report_style.Equals("自定义"))
            {
                WriteIntoDocument(oDoc, "QX", "全市");
                WriteIntoDocument(oDoc, "QX_subject", config.subject);
            }
            else if (config.report_style.Equals("学校"))
            {
                WriteIntoDocument(oDoc, "QX", config.school);
                WriteIntoDocument(oDoc, "QX_subject", config.subject);
            }
        }

        public static void WriteFrontPage(Configuration config, Microsoft.Office.Interop.Word._Document oDoc, string school)
        {
            WriteIntoDocument(oDoc, "date", config.year + "年" + config.month);
            if (config.subject.Equals("总分"))
            {
                WriteIntoDocument(oDoc, "title_1", config.year + GK_ZF_title_1);
                WriteIntoDocument(oDoc, "title_2", GK_ZF_title_2);
                WriteIntoDocument(oDoc, "subject", school);
            }

            else
            {
                WriteIntoDocument(oDoc, "title_2", XX_title);
                if (config.subject.Contains("理综"))
                {
                    WriteIntoDocument(oDoc, "QX", school);
                    WriteIntoDocument(oDoc, "ZH", "理科综合");
                    WriteIntoDocument(oDoc, "QX_ZH_subject", config.subject.Substring(3));
                }
                else if (config.subject.Contains("文综"))
                {
                    WriteIntoDocument(oDoc, "QX", school);
                    WriteIntoDocument(oDoc, "ZH", "文科综合");
                    WriteIntoDocument(oDoc, "QX_ZH_subject", config.subject.Substring(3));
                }
                else
                {
                    WriteIntoDocument(oDoc, "QX", school);
                    WriteIntoDocument(oDoc, "QX_subject", config.subject);
                }
            }
        }

        public static void WriteFrontPage(Configuration config, Microsoft.Office.Interop.Word._Document oDoc)
        {
            WriteIntoDocument(oDoc, "date", config.year + "年" + config.month);

            if (config.exam.Equals("中考"))
                {
                    
                    if (config.report_style.Equals("总体"))
                    {
                        WriteIntoDocument(oDoc, "title_1", ZK_title_1);
                        WriteIntoDocument(oDoc, "title_2", ZK_title_2);
                        WriteIntoDocument(oDoc, "subject", config.subject);
                    }
                    else if (config.report_style.Equals("区县"))
                    {
                        WriteIntoDocument(oDoc, "title_1", ZK_title_1);
                        WriteIntoDocument(oDoc, "title_2", ZK_QX_title_2);
                        WriteIntoDocument(oDoc, "QX", config.QX);
                        WriteIntoDocument(oDoc, "QX_subject", config.subject);
                    }
                    else if (config.report_style.Equals("城郊"))
                    {
                        WriteIntoDocument(oDoc, "title_1", ZK_title_1);
                        WriteIntoDocument(oDoc, "title_2", GK_CJ_title_2);
                        WriteIntoDocument(oDoc, "QX", "全市");
                        WriteIntoDocument(oDoc, "QX_subject", config.subject);
                    }
                    else if (config.subject.Equals("中考行政版"))
                    {
                        WriteIntoDocument(oDoc, "title_1", "");
                        WriteIntoDocument(oDoc, "title_2", config.year + ZK_ZF_title_xz_1);
                        WriteIntoDocument(oDoc, "subject", ZK_ZF_title_xz_2);
                        WriteIntoDocument(oDoc, "company", "北京教育考试院科研办");
                    }
                }
            else if (config.exam.Equals("会考"))
                {
                    WriteIntoDocument(oDoc, "HK_title_1", HK_title_1);
                    WriteIntoDocument(oDoc, "subject", config.subject);
                }
            else if (config.exam.Equals("高考"))
                {
                    if (config.subject.Equals("总分"))
                    {
                        WriteIntoDocument(oDoc, "title_1", config.year + GK_ZF_title_1);
                        WriteIntoDocument(oDoc, "title_2", GK_ZF_title_2);
                        if (config.report_style.Equals("城郊"))
                            WriteIntoDocument(oDoc, "subject", "城区与郊区");
                        else if (config.report_style.Equals("两类示范校"))
                            WriteIntoDocument(oDoc, "subject", "两类示范校");
                        else if (config.report_style.Equals("区县"))
                            WriteIntoDocument(oDoc, "subject", config.QX);
                        else if (config.report_style.Equals("总体"))
                            WriteIntoDocument(oDoc, "subject", "全市");
                        
                    }
                    else if (config.subject.Equals("高考行政版"))
                    {
                        WriteIntoDocument(oDoc, "title_1", "");
                        WriteIntoDocument(oDoc, "title_2", config.year + GK_ZF_title_xz_1);
                        WriteIntoDocument(oDoc, "subject", GK_ZF_title_xz_2);
                        WriteIntoDocument(oDoc, "company", "北京教育考试院科研办");
                    }
                    else
                    {
                        WriteIntoDocument(oDoc, "title_1", GK_title_1);
                        if (config.report_style.Equals("城郊"))
                        {
                            WriteIntoDocument(oDoc, "title_2", GK_CJ_title_2);
                            if (config.subject.Contains("理综"))
                            {
                                WriteIntoDocument(oDoc, "CJ_ZH", "理科综合");
                                WriteIntoDocument(oDoc, "CJ_ZH_subject", config.subject.Substring(3));
                            }
                            else if (config.subject.Contains("文综"))
                            {
                                WriteIntoDocument(oDoc, "CJ_ZH", "文科综合");
                                WriteIntoDocument(oDoc, "CJ_ZH_subject", config.subject.Substring(3));
                            }
                            else
                            {
                                WriteIntoDocument(oDoc, "QX", "全市");
                                WriteIntoDocument(oDoc, "QX_subject", config.subject);
                            }
                        }
                        else if (config.report_style.Equals("两类示范校"))
                        {
                            WriteIntoDocument(oDoc, "title_2", GK_SF_title_2);
                            if (config.subject.Contains("理综"))
                            {
                                WriteIntoDocument(oDoc, "CJ_ZH", "理科综合");
                                WriteIntoDocument(oDoc, "CJ_ZH_subject", config.subject.Substring(3));
                            }
                            else if (config.subject.Contains("文综"))
                            {
                                WriteIntoDocument(oDoc, "CJ_ZH", "文科综合");
                                WriteIntoDocument(oDoc, "CJ_ZH_subject", config.subject.Substring(3));
                            }
                            else
                            {
                                WriteIntoDocument(oDoc, "QX", "全市");
                                WriteIntoDocument(oDoc, "QX_subject", config.subject);
                            }
                        }
                        else if (config.report_style.Equals("区县"))
                        {
                            WriteIntoDocument(oDoc, "title_2", GK_QX_title_2);
                            if (config.subject.Contains("理综"))
                            {
                                WriteIntoDocument(oDoc, "QX", config.QX);
                                WriteIntoDocument(oDoc, "ZH", "理科综合");
                                WriteIntoDocument(oDoc, "QX_ZH_subject", config.subject.Substring(3));
                            }
                            else if (config.subject.Contains("文综"))
                            {
                                WriteIntoDocument(oDoc, "QX", config.QX);
                                WriteIntoDocument(oDoc, "ZH", "文科综合");
                                WriteIntoDocument(oDoc, "QX_ZH_subject", config.subject.Substring(3));
                            }
                            else
                            {
                                WriteIntoDocument(oDoc, "QX", config.QX);
                                WriteIntoDocument(oDoc, "QX_subject", config.subject);
                            }
                        }
                        else if (config.report_style.Equals("总体"))
                        {
                            WriteIntoDocument(oDoc, "title_2", ZK_title_2);
                            if (config.subject.Contains("理综"))
                            {
                                WriteIntoDocument(oDoc, "CJ_ZH", "理科综合");
                                WriteIntoDocument(oDoc, "CJ_ZH_subject", config.subject.Substring(3));
                            }
                            else if (config.subject.Contains("文综"))
                            {
                                WriteIntoDocument(oDoc, "CJ_ZH", "文科综合");
                                WriteIntoDocument(oDoc, "CJ_ZH_subject", config.subject.Substring(3));
                            }
                            else
                            {
                                WriteIntoDocument(oDoc, "subject", config.subject);
                            }
                        }
                        else if (config.report_style.Equals("自定义"))
                        {
                            WriteIntoDocument(oDoc, "title_2", GK_CUS_title_2);
                            if (config.subject.Contains("理综"))
                            {
                                WriteIntoDocument(oDoc, "CJ_ZH", "理科综合");
                                WriteIntoDocument(oDoc, "CJ_ZH_subject", config.subject.Substring(3));
                            }
                            else if (config.subject.Contains("文综"))
                            {
                                WriteIntoDocument(oDoc, "CJ_ZH", "文科综合");
                                WriteIntoDocument(oDoc, "CJ_ZH_subject", config.subject.Substring(3));
                            }
                            else
                            {
                                WriteIntoDocument(oDoc, "QX", "全市");
                                WriteIntoDocument(oDoc, "QX_subject", config.subject);
                            }
                        }

                    }
                
            }

        }

        public static void WriteIntoDocument(Microsoft.Office.Interop.Word._Document oDoc, string BookmarkName, string FillName)
        {
            object bookmarkName = BookmarkName;
            Microsoft.Office.Interop.Word.Bookmark bm = oDoc.Bookmarks.get_Item(ref bookmarkName);//返回书签 
            bm.Range.Text = FillName;//设置书签域的内容
        }
        public static void WSLG_Save(Configuration config, Microsoft.Office.Interop.Word._Document oDoc, Microsoft.Office.Interop.Word._Application oWord)
        {
            insertAddons(config, oDoc, oWord);
            object oMissing = System.Reflection.Missing.Value;
            string addr = config.save_address + @"\";
            string final = config.year + "年" + config.subject + "文史、理工类数据统计分析报告(最终版）.docx"; ;
            final = addr + final;
            oDoc.SaveAs(final, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
            oDoc.Close(oMissing, oMissing, oMissing);
            oWord.Quit(oMissing, oMissing, oMissing);

        }
        public static void Save(Configuration config, Microsoft.Office.Interop.Word._Document oDoc, Microsoft.Office.Interop.Word._Application oWord, string school)
        {
            insertAddons(config, oDoc, oWord);
            object oMissing = System.Reflection.Missing.Value;
            string addr = config.save_address + @"\";
            string final = "a.docx";

            if (config.subject.Equals("总分"))
            {
                final = config.year + "年总分统计分析报告(" + school + ").docx";
            }
            else
            {
                if (config.subject.Contains("理综") || config.subject.Contains("文综"))
                    final = config.year + "年北京市" + school + config.subject.Substring(3) + "学校数据统计分析报告.docx";
                else
                    final = config.year + "年北京市" + school + config.subject + "学校数据统计分析报告.docx";
            }

            final = addr + final;
            oDoc.SaveAs(final, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
            oDoc.Close(oMissing, oMissing, oMissing);
            oWord.Quit(oMissing, oMissing, oMissing);
        }
        public static void Save(Configuration config, Microsoft.Office.Interop.Word._Document oDoc, Microsoft.Office.Interop.Word._Application oWord)
        {
            if(!config.subject.Contains("行政版"))
                insertAddons(config, oDoc, oWord);
            object oMissing = System.Reflection.Missing.Value;
            string addr = config.save_address + @"\";
            string final = "a.docx";
            if (config.exam.Equals("中考"))
            {
                if (config.report_style.Equals("总体"))
                {
                    final = config.year + "年北京市高级中等学校招生考试" + config.subject.ToString() + "数据统计分析报告.docx";
                }
                else if (config.report_style.Equals("区县"))
                {
                    final = config.year + "年" + config.QX + config.subject.ToString() + "分类校数据统计分析报告.docx";
                }
                else if (config.subject.Equals("中考行政版"))
                {
                    final = config.year + "年北京市中考试卷行政版分析报告.docx";
                }
            }
            else if (config.exam.Equals("会考"))
            {
                final = config.year + "年" + config.subject.ToString() + "北京市普通高中会考统计报告.docx";
            }
            else if (config.exam.Equals("高考"))
            {
                if (config.subject.Equals("总分"))
                {
                    if (config.report_style.Equals("城郊"))
                        final = config.year + "年北京市普通高考试卷总分统计分析报告(城区与郊区).docx";
                    else if (config.report_style.Equals("两类示范校"))
                        final = config.year + "年北京市普通高考试卷总分统计分析报告(两类示范校).docx";
                    else if (config.report_style.Equals("区县"))
                        final = config.year + "北京市普通高考试卷总分统计分析报告（" + config.QX + "）.docx";
                    else if (config.report_style.Equals("总体"))
                        final = config.year + "年北京市普通高考试卷总分统计分析报告(全市).docx";
                    
                }
                if(config.subject.Equals("高考行政版"))
                {
                    final = config.year + "年北京市普通高考试卷行政版分析报告.docx";
                }
                else
                {
                    if (config.report_style.Equals("城郊"))
                    {
                        if (config.subject.Contains("理综") || config.subject.Contains("文综"))
                            final = config.year + "年" + config.subject.Substring(3) + "城区、郊区数据统计分析报告.docx";
                        else
                            final = config.year + "年" + config.subject + "城区、郊区数据统计分析报告.docx";
                    }
                    else if (config.report_style.Equals("两类示范校"))
                    {
                        if (config.subject.Contains("理综") || config.subject.Contains("文综"))
                            final = config.year + "年" + config.subject.Substring(3) + "示范校数据统计分析报告.docx";
                        else
                            final = config.year + "年" + config.subject + "示范校数据统计分析报告.docx";
                    }
                    else if (config.report_style.Equals("区县"))
                    {
                        if (config.subject.Contains("理综") || config.subject.Contains("文综"))
                            final = config.year + "年" + config.QX + config.subject.Substring(3) + "分类校数据统计分析报告.docx";
                        else
                            final = config.year + "年" + config.QX + config.subject + "分类校数据统计分析报告.docx";
                    }
                    else if (config.report_style.Equals("总体"))
                    {
                        if (config.subject.Contains("理综") || config.subject.Contains("文综"))
                            final = config.year + "年" + config.subject.Substring(3) + "数据统计分析报告(最终版）.docx";
                        else
                            final = config.year + "年" + config.subject + "数据统计分析报告(最终版）.docx";
                    }
                    else if (config.report_style.Equals("自定义"))
                    {
                        if (config.subject.Contains("理综") || config.subject.Contains("文综"))
                            final = config.year + "年" + config.subject.Substring(3) + "自选数据统计分析报告(最终版）.docx";
                        else
                            final = config.year + "年" + config.subject + "自选数据统计分析报告(最终版）.docx";
                    }
                }
            }
            final = addr + final;
            oDoc.SaveAs(final, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
            oDoc.Close(oMissing, oMissing, oMissing);
            oWord.Quit(oMissing, oMissing, oMissing);
        }
        public static void insertAddons(Configuration config, Microsoft.Office.Interop.Word._Document doc, Microsoft.Office.Interop.Word._Application oWord)
        {
            doc.Characters.Last.InsertBreak(Microsoft.Office.Interop.Word.WdBreakType.wdPageBreak);
            object oEndOfDoc = "\\endofdoc";
            object oMissing = System.Reflection.Missing.Value;

            Microsoft.Office.Interop.Word.Range first = doc.Paragraphs.Add(ref oMissing).Range;


            if (config.subject.Contains("理综") || config.subject.Contains("文综") || (config.subject.Equals("总分") && !config.report_style.Equals("总体")))
                first.set_Style("ExamTitle0");
            else
                first.set_Style("ExamTitle1");
            first.InsertBefore("附录" + "\n");

            doc.Characters.Last.Select();
            oWord.Selection.HomeKey(Microsoft.Office.Interop.Word.WdUnits.wdLine, oMissing);
            oWord.Selection.Delete(Microsoft.Office.Interop.Word.WdUnits.wdCharacter, oMissing);
            oWord.Selection.Range.set_Style("ExamBodyText");

            Microsoft.Office.Interop.Word.Range range = doc.Bookmarks.get_Item(ref oEndOfDoc).Range;

            range.InsertFile(@config.CurrentDirectory + @"\addon.doc", oMissing, false, false, false);
            foreach (Microsoft.Office.Interop.Word.TableOfContents table in doc.TablesOfContents)
                table.Update();
        }
        public static String ToSBC(String input)
        {
            // 半角转全角：
            char[] c = input.ToCharArray();
            for (int i = 0; i < c.Length; i++)
            {
                if (c[i] == 32)
                {
                    c[i] = (char)12288;
                    continue;
                }
                if (c[i] < 127)
                    c[i] = (char)(c[i] + 65248);
            }
            return new String(c);
        }
        public static void XZ_group_separate(DataTable temp_dt, Configuration config, string th)
        {
            if (!temp_dt.Columns.Contains("xz_groups"))
                temp_dt.Columns.Add("xz_groups", typeof(string));
            var xz_tuple = from row in temp_dt.AsEnumerable()
                           group row by row.Field<string>(th) into grp
                           select new
                           {
                               name = grp.Key
                           };
            foreach (var item in xz_tuple)
            {
                DataView dv = temp_dt.equalfilter(th, item.name).DefaultView;
                DataTable inter_table = dv.ToTable();
                inter_table.SeperateGroups(config._grouptype, config._group_num, "xz_groups");
                var temp = from row in temp_dt.AsEnumerable()
                           join row2 in inter_table.AsEnumerable() on row.Field<string>("kh") equals row2.Field<string>("kh")
                           where row.Field<string>(th) == item.name
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
        public static string choiceTransfer(string choice)
        {
            switch (choice.Trim())
            {
                case "0":
                    return "未选";
                case "1":
                    return "A";
                case "2":
                    return "B";
                case "4":
                    return "C";
                case "8":
                    return "D";
                case "@":
                    return "E";
                case "P":
                    return "F";
                case "p":
                    return "G";
                case "3":
                    return "AB";
                case "5":
                    return "AC";
                case "6":
                    return "BC";
                case "7":
                    return "ABC";
                case "9":
                    return "AD";
                case ":":
                    return "BD";
                case ";":
                    return "ABD";
                case "<":
                    return "CD";
                case "=":
                    return "ACD";
                case ">":
                    return "BCD";
                case "?":
                    return "ABCD";
                case "A":
                    return "AE";
                case "B":
                    return "BE";
                case "C":
                    return "ABE";
                case "D":
                    return "CE";
                case "E":
                    return "ACE";
                case "F":
                    return "BCE";
                case "G":
                    return "ABCE";
                case "H":
                    return "DE";
                case "I":
                    return "ADE";
                case "J":
                    return "BDE";
                case "K":
                    return "ABDE";
                case "L":
                    return "CDE";
                case "M":
                    return "ACDE";
                case "N":
                    return "BCDE";
                case "O":
                    return "ABCDE";
                case "":
                    return "";
                default:
                    return null;

            }
        }

        public static bool isContain(string da, string ans)
        {
            char[] ans_ = ans.ToCharArray();
            foreach (char temp in ans_)
            {
                if (!da.Contains(temp))
                    return false;
            }
            return true;
        }
        public static string hk_en_trans(string name)
        {
            switch (name)
            {
                case "语文":
                    return "Chinese";
                case "数学":
                    return "Mathematics";
                case "英语":
                    return "English";
                case "化学":
                    return "Chemistry";
                case "物理":
                    return "Physics";
                case "生物":
                    return "Biology";
                case "政治":
                    return "Politics";
                case "地理":
                    return "Geography";
                case "历史":
                    return "History";
                default:
                    return "";
            }
        }
        public static string hk_en_trans_dt(string name)
        {
            switch (name)
            {
                case "语文":
                    return "chinese";
                case "数学":
                    return "math";
                case "英语":
                    return "english";
                case "化学":
                    return "chemistry";
                case "物理":
                    return "physics";
                case "生物":
                    return "biology";
                case "政治":
                    return "politics";
                case "地理":
                    return "geography";
                case "历史":
                    return "history";
                default:
                    return "";
            }
        }
        public static string hk_lang_trans(string name)
        {
            switch (name)
            {
                case "语文":
                    return "yw";
                case "数学":
                    return "sx";
                case "英语":
                    return "yy";
                case "化学":
                    return "hx";
                case "物理":
                    return "wl";
                case "生物":
                    return "sw";
                case "政治":
                    return "zz";
                case "地理":
                    return "dl";
                case "历史":
                    return "ls";
                case "中考":
                    return "zk";
                case "高考":
                    return "gk";
                case "会考":
                    return "hk";
                case "数学理":
                    return "sxl";
                case "数学文":
                    return "sxw";
                case "总分":
                    return "zf";
                case "高考行政版":
                    return "gk_xz";
                case "中考行政版":
                    return "zk_xz";
                case "yw":
                    return "语文";
                case "sx":
                    return "数学";
                case "yy":
                    return "英语";
                case "hx":
                    return "化学";
                case "wl":
                    return "物理";
                case "sw":
                    return "生物";
                case "zz":
                    return "政治";
                case "dl":
                    return "地理";
                case "ls":
                    return "历史";
                case "sxl":
                    return "数学理";
                case "sxw":
                    return "数学文"; 
                case "zf":
                    return "总分";
                case "gk_xz":
                    return "高考行政版";
                case "zk_xz":
                    return "中考行政版";
                case "思想品德":
                    return "sxpd";
                case "sxpd":
                    return "思想品德";
                default:
                    return "";
            }
        }
        public static string language_trans(string name)
        {
            switch (name)
            {
                case "语文":
                    return "yw";
                case "数学":
                    return "sx";
                case "英语":
                    return "yy";
                case "理综-化学":
                case "化学":
                    return "hx";
                case "理综-物理":
                case "物理":
                    return "wl";
                case "理综-生物":
                case "生物":
                    return "sw";
                case "文综-政治":
                case "政治":
                    return "zz";
                case "文综-地理":
                case "地理":
                    return "dl";
                case "文综-历史":
                case "历史":
                    return "ls";
                case "中考":
                    return "zk";
                case "高考":
                    return "gk";
                case "2020新高考":
                    return "ngk";
                case "会考":
                    return "hk";
                case "数学理":
                    return "sxl";
                case "数学文":
                    return "sxw";
                case "总分":
                    return "zf";
                case "高考行政版":
                    return "gk_xz";
                case "中考行政版":
                    return "zk_xz";
                case "思想品德":
                    return "sxpd";

                case "yw":
                    return "语文";
                case "sx":
                    return "数学";
                case "yy":
                    return "英语";
                case "hx":
                    return "理综-化学";
                case "wl":
                    return "理综-物理";
                case "sw":
                    return "理综-生物";
                case "zz":
                    return "文综-政治";
                case "dl":
                    return "文综-地理";
                case "ls":
                    return "文综-历史";
                case "sxl":
                    return "数学理";
                case "sxw":
                    return "数学文";
                case "zf":
                    return "总分";
                case "gk_xz":
                    return "高考行政版";
                case "zk_xz":
                    return "中考行政版";
                case "sxpd":
                    return "思想品德";
                default:
                    return "";
            }
        }
        public static bool is_gk_zh(string exam, string sub)
        {
            if (exam.Equals("高考") || exam.Equals("gk"))
            {
                if (sub.Contains("理综") || sub.Contains("文综"))
                    return true;
                else
                    return false;
            }
            return false;
        }
        public static string get_zt_tablename(string year, string exam, string sub)
        {
            return year + "_" + exam + "_" + sub;
        }
        public static string get_tablename(string year, string exam, string sub)
        {
            return year + "_" + exam + "_" + sub;
        }
        public static string get_basic_tablename(string year, string exam, string sub)
        {
            return year + "_" + exam + "_" + sub + "_basic";
        }
        public static string get_group_tablename(string year, string exam, string sub)
        {
            return year + "_" + exam + "_" + sub + "_group";
        }
        public static string get_ans_tablename(string year, string exam, string sub)
        {
            return get_tablename(year, exam, sub) + "_ans";
        }
        public static string get_fz_tablename(string year, string exam, string sub)
        {
            return get_tablename(year, exam, sub) + "_fz";
        }
        public static void create_groups_table(DataTable groups_data, string filename)
        {
            string conn = @"Provider=vfpoledb;Data Source=" + save_address + ";Collating Sequence=machine;";
            string charsize = ConfigurationManager.AppSettings["charsize"].ToString().Trim();
            OleDbConnection dbfConnection = new OleDbConnection(conn);
            StringBuilder objectdata = new StringBuilder();
            objectdata.Clear();
            int i = 0;
            
            objectdata.Append("CREATE TABLE `" + filename + "` (\n");
            int count = 0;
            foreach (DataColumn dc in groups_data.Columns)
            {
                objectdata.Append("\t`" + dc.ColumnName + "` ");
                if (dc.DataType.ToString().Equals("System.String"))
                    objectdata.Append("c(" + charsize + ")");
                else if (dc.DataType.ToString().Equals("System.Decimal"))
                    objectdata.Append("n(4,1)");
                else
                    i++;
                count++;
                if (count != groups_data.Columns.Count)
                    objectdata.Append(",\n");
                else
                    objectdata.Append(");");
            }
            
            OleDbCommand group_create = new OleDbCommand(objectdata.ToString(), dbfConnection);
            dbfConnection.Open();
            group_create.ExecuteNonQuery();
            OleDbCommand group_insert = new OleDbCommand();
            group_insert.Connection = dbfConnection;
            OleDbTransaction group_trans = null;
            group_trans = group_insert.Connection.BeginTransaction();
            group_insert.Transaction = group_trans;

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
                    
                    if (i != groups_data.Columns.Count - 1)
                        objectdata.Append(",");
                    else
                        objectdata.Append(");");
                    
                }

                group_insert.CommandText = objectdata.ToString();
                group_insert.ExecuteNonQuery();

            }
            group_trans.Commit();
            dbfConnection.Close();
        }
        public static void zip(string dir, string zipedFile)
        {
            using (Ionic.Zip.ZipFile zip = new Ionic.Zip.ZipFile(zipedFile + ".zip", Encoding.Default))
            {
                zip.AddDirectory(dir, new FileInfo(dir).Name);
                zip.Save();
            }

            System.IO.File.Move(zipedFile + ".zip", Directory.GetParent(dir).Parent.FullName + Path.DirectorySeparatorChar + zipedFile + ".zip");

            Directory.Delete(dir, true);
        }

        public static void unZip(string zipFile, string outputdir)
        {
            ReadOptions options = new ReadOptions();
            options.Encoding = Encoding.Default;
            using (ZipFile zip = ZipFile.Read(zipFile, options))
            {
                foreach (ZipEntry z in zip)
                {
                    FileInfo f = new FileInfo(outputdir + "/" + z.FileName);
                    if (f.Exists)
                    {
                        string parent = f.Directory.Name;
                        string name = f.Name.Replace(f.Extension, "");
                        string str = parent + "|" + name;
                        //if (MessageBox.Show("文件(" + str + ")已经存在，是否替换？", "确认", MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.Yes)
                        //{
                        //    f.Delete();
                        //}
                        //else
                        //{
                        //    continue;
                        //}
                    }
                    z.Extract(outputdir);
                }
            }

        }

        public static string OperatorTrans(string oper)
        {
            switch (oper)
            {
                case "大于":
                    return ">";
                case "小于":
                    return "<";
                case "等于":
                    return "=";
                case "大于等于":
                    return ">=";
                case "小于等于":
                    return "<=";
                case "不等于":
                    return "<>";
                case "近似于":
                    return "like";
                case "并且":
                    return "and";
                case "或者":
                    return "or";
                default:
                    return "";
            }
        }

        public static string QXTrans(string qx)
        {
            switch (qx)
            {
                case "dc":
                    return "东城";
                case "xc":
                    return "西城";
                case "cy":
                    return "朝阳";
                case "ft":
                    return "丰台";
                case "sjs":
                    return "石景山";
                case "hd":
                    return "海淀";
                case "mtg":
                    return "门头沟";
                case "ys":
                    return "燕山";
                case "fs":
                    return "房山";
                case "tz":
                    return "通州";
                case "sy":
                    return "顺义";
                case "cp":
                    return "昌平";
                case "dx":
                    return "大兴";
                case "hr":
                    return "怀柔";
                case "pg":
                    return "平谷";
                case "my":
                    return "密云";
                case "yq":
                    return "延庆";
                default:
                    return "";
            }
        }

    }
}
