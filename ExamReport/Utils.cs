using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.OleDb;
using System.Configuration;

namespace ExamReport
{
    public static class Utils
    {
        public static string save_address;
        public static string exam;
        public static string subject;
        public static string report_style;
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

        public static int smooth_degree = 10;

        public static decimal fullmark = 0;

        public static List<decimal> GroupMark = new List<decimal>();

        public static decimal shengwu_zhengzhi;
        public static decimal wuli_lishi;
        public static decimal huaxue_dili;

        public static object clipboardLock;

        public static string ZK_title_1 = "北京市高级中等学校招生考试";
        public static string ZK_title_2 = "实测数据统计分析报告";
        public static string ZK_QX_title_2 = "分类校数据统计分析报告";
        public static string HK_title_1 = "北京市高中会考数据统计分析报告";
        public static string GK_title_1 = "北京市普通高等学校招生全国统一考试";
        public static string GK_CJ_title_2 = "城区、郊区数据统计分析报告";
        public static string GK_SF_title_2 = "示范校数据统计分析报告";
        public static string GK_QX_title_2 = "分类校数据统计分析报告";
        public static string GK_title_2 = "实测数据统计分析报告";
        public static string GK_ZF_title_1 = "年北京市普通高考";
        public static string GK_ZF_title_2 = "试卷总分统计分析报告";
        public static string GK_WSLG_title_2 = "文史、理工类数据统计分析报告";
        public static string XX_title = "学校数据统计分析报告";

        public static void WSLG_WriteFrontPage(Microsoft.Office.Interop.Word._Document oDoc)
        {
            WriteIntoDocument(oDoc, "title_1", GK_title_1);
            WriteIntoDocument(oDoc, "title_2", GK_WSLG_title_2);
            if (report_style.Equals("区县"))
            {
                WriteIntoDocument(oDoc, "QX", QX);
                WriteIntoDocument(oDoc, "QX_subject", subject);
            }
            else if (report_style.Equals("总体"))
            {
                WriteIntoDocument(oDoc, "QX", "全市");
                WriteIntoDocument(oDoc, "QX_subject", subject);
            }
        }

        public static void WriteFrontPage(Microsoft.Office.Interop.Word._Document oDoc, string school)
        {
            if (subject.Equals("总分"))
            {
                WriteIntoDocument(oDoc, "title_1", year + GK_ZF_title_1);
                WriteIntoDocument(oDoc, "title_2", GK_ZF_title_2);
                WriteIntoDocument(oDoc, "subject", school);
            }

            else
            {
                WriteIntoDocument(oDoc, "title_2", XX_title);
                if (subject.Contains("理综"))
                {
                    WriteIntoDocument(oDoc, "QX", school);
                    WriteIntoDocument(oDoc, "ZH", "理科综合");
                    WriteIntoDocument(oDoc, "QX_ZH_subject", subject.Substring(3));
                }
                else if (subject.Contains("文综"))
                {
                    WriteIntoDocument(oDoc, "QX", school);
                    WriteIntoDocument(oDoc, "ZH", "文科综合");
                    WriteIntoDocument(oDoc, "QX_ZH_subject", subject.Substring(3));
                }
                else
                {
                    WriteIntoDocument(oDoc, "QX", school);
                    WriteIntoDocument(oDoc, "QX_subject", subject);
                }
            }
        }

        public static void WriteFrontPage(Microsoft.Office.Interop.Word._Document oDoc)
        {
            WriteIntoDocument(oDoc, "date", year + "年" + month);
            
                if (exam.Equals("中考"))
                {
                    WriteIntoDocument(oDoc, "title_1", ZK_title_1);
                    if (report_style.Equals("总体"))
                    {

                        WriteIntoDocument(oDoc, "title_2", ZK_title_2);
                        WriteIntoDocument(oDoc, "subject", subject);
                    }
                    else if (report_style.Equals("区县"))
                    {
                        WriteIntoDocument(oDoc, "title_2", ZK_QX_title_2);
                        WriteIntoDocument(oDoc, "QX", QX);
                        WriteIntoDocument(oDoc, "QX_subject", subject);
                    }
                }
                else if (exam.Equals("会考"))
                {
                    WriteIntoDocument(oDoc, "HK_title_1", HK_title_1);
                    WriteIntoDocument(oDoc, "subject", subject);
                }
                else if (exam.Equals("高考"))
                {
                    if (subject.Equals("总分"))
                    {
                        WriteIntoDocument(oDoc, "title_1", year + GK_ZF_title_1);
                        WriteIntoDocument(oDoc, "title_2", GK_ZF_title_2);
                        if (report_style.Equals("城郊"))
                            WriteIntoDocument(oDoc, "subject", "城区与郊区");
                        else if (report_style.Equals("两类示范校"))
                            WriteIntoDocument(oDoc, "subject", "两类示范校");
                        else if (report_style.Equals("区县"))
                            WriteIntoDocument(oDoc, "subject", QX);
                        else if (report_style.Equals("总体"))
                            WriteIntoDocument(oDoc, "subject", "全市");
                        
                    }
                    else
                    {
                        WriteIntoDocument(oDoc, "title_1", GK_title_1);
                        if (report_style.Equals("城郊"))
                        {
                            WriteIntoDocument(oDoc, "title_2", GK_CJ_title_2);
                            if (subject.Contains("理综"))
                            {
                                WriteIntoDocument(oDoc, "CJ_ZH", "理科综合");
                                WriteIntoDocument(oDoc, "CJ_ZH_subject", subject.Substring(3));
                            }
                            else if (subject.Contains("文综"))
                            {
                                WriteIntoDocument(oDoc, "CJ_ZH", "文科综合");
                                WriteIntoDocument(oDoc, "CJ_ZH_subject", subject.Substring(3));
                            }
                            else
                            {
                                WriteIntoDocument(oDoc, "QX", "全市");
                                WriteIntoDocument(oDoc, "QX_subject", subject);
                            }
                        }
                        else if (report_style.Equals("两类示范校"))
                        {
                            WriteIntoDocument(oDoc, "title_2", GK_SF_title_2);
                            if (subject.Contains("理综"))
                            {
                                WriteIntoDocument(oDoc, "CJ_ZH", "理科综合");
                                WriteIntoDocument(oDoc, "CJ_ZH_subject", subject.Substring(3));
                            }
                            else if (subject.Contains("文综"))
                            {
                                WriteIntoDocument(oDoc, "CJ_ZH", "文科综合");
                                WriteIntoDocument(oDoc, "CJ_ZH_subject", subject.Substring(3));
                            }
                            else
                            {
                                WriteIntoDocument(oDoc, "QX", "全市");
                                WriteIntoDocument(oDoc, "QX_subject", subject);
                            }
                        }
                        else if (report_style.Equals("区县"))
                        {
                            WriteIntoDocument(oDoc, "title_2", GK_QX_title_2);
                            if (subject.Contains("理综"))
                            {
                                WriteIntoDocument(oDoc, "QX", QX);
                                WriteIntoDocument(oDoc, "ZH", "理科综合");
                                WriteIntoDocument(oDoc, "QX_ZH_subject", subject.Substring(3));
                            }
                            else if (subject.Contains("文综"))
                            {
                                WriteIntoDocument(oDoc, "QX", QX);
                                WriteIntoDocument(oDoc, "ZH", "文科综合");
                                WriteIntoDocument(oDoc, "QX_ZH_subject", subject.Substring(3));
                            }
                            else
                            {
                                WriteIntoDocument(oDoc, "QX", QX);
                                WriteIntoDocument(oDoc, "QX_subject", subject);
                            }
                        }
                        else if (report_style.Equals("总体"))
                        {
                            WriteIntoDocument(oDoc, "title_2", ZK_title_2);
                            if (subject.Contains("理综"))
                            {
                                WriteIntoDocument(oDoc, "CJ_ZH", "理科综合");
                                WriteIntoDocument(oDoc, "CJ_ZH_subject", subject.Substring(3));
                            }
                            else if (subject.Contains("文综"))
                            {
                                WriteIntoDocument(oDoc, "CJ_ZH", "文科综合");
                                WriteIntoDocument(oDoc, "CJ_ZH_subject", subject.Substring(3));
                            }
                            else
                            {
                                WriteIntoDocument(oDoc, "subject", subject);
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
        public static void WSLG_Save(Microsoft.Office.Interop.Word._Document oDoc, Microsoft.Office.Interop.Word._Application oWord)
        {
            insertAddons(oDoc);
            object oMissing = System.Reflection.Missing.Value;
            string addr = save_address + @"\";
            string final = year + "年" + subject + "文史、理工类数据统计分析报告(最终版）.docx";;
            final = addr + final;
            oDoc.SaveAs(final, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
            oDoc.Close(oMissing, oMissing, oMissing);
            oWord.Quit(oMissing, oMissing, oMissing);

        }
        public static void Save(Microsoft.Office.Interop.Word._Document oDoc, Microsoft.Office.Interop.Word._Application oWord, string school)
        {
            insertAddons(oDoc);
            object oMissing = System.Reflection.Missing.Value;
            string addr = save_address + @"\";
            string final = "a.docx";

            if (subject.Equals("总分"))
            {
                final = year + "年总分统计分析报告(" + school + ").docx";
            }
            else
            {
                if (subject.Contains("理综") || subject.Contains("文综"))
                    final = year + "年北京市" + school + subject.Substring(3) + "学校数据统计分析报告.docx";
                else
                    final = year + "年北京市" + school + subject + "学校数据统计分析报告.docx";
            }

            final = addr + final;
            oDoc.SaveAs(final, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
            oDoc.Close(oMissing, oMissing, oMissing);
            oWord.Quit(oMissing, oMissing, oMissing);
        }
        public static void Save(Microsoft.Office.Interop.Word._Document oDoc, Microsoft.Office.Interop.Word._Application oWord)
        {
            insertAddons(oDoc);
            object oMissing = System.Reflection.Missing.Value;
            string addr = save_address + @"\";
            string final = "a.docx";
            if (exam.Equals("中考"))
            {
                if (report_style.Equals("总体"))
                {
                    final = year + "年北京市高级中等学校招生考试" + subject.ToString() + "数据统计分析报告.docx";
                }
                else if (report_style.Equals("区县"))
                {
                    final = year + "年" + QX + subject.ToString() + "分类校数据统计分析报告.docx";
                }
            }
            else if (exam.Equals("会考"))
            {
                final = year + "年" + subject.ToString() + "北京市普通高中会考统计报告.docx";
            }
            else if (exam.Equals("高考"))
            {
                if (subject.Equals("总分"))
                {
                    if (report_style.Equals("城郊"))
                        final = year + "年北京市普通高考试卷总分统计分析报告(城区与郊区).docx";
                    else if (report_style.Equals("两类示范校"))
                        final = year + "年北京市普通高考试卷总分统计分析报告(两类示范校).docx";
                    else if (report_style.Equals("区县"))
                        final = year + "北京市普通高考试卷总分统计分析报告（" + QX + "）.docx";
                    else if (report_style.Equals("总体"))
                        final = year + "年北京市普通高考试卷总分统计分析报告(全市).docx";
                    
                }
                else
                {
                    if (report_style.Equals("城郊"))
                    {
                        if (subject.Contains("理综") || subject.Contains("文综"))
                            final = year + "年" + subject.Substring(3) + "城区、郊区数据统计分析报告.docx";
                        else
                            final = year + "年" + subject + "城区、郊区数据统计分析报告.docx";
                    }
                    else if (report_style.Equals("两类示范校"))
                    {
                        if (subject.Contains("理综") || subject.Contains("文综"))
                            final = year + "年" + subject.Substring(3) + "示范校数据统计分析报告.docx";
                        else
                            final = year + "年" + subject + "示范校数据统计分析报告.docx";
                    }
                    else if (report_style.Equals("区县"))
                    {
                        if (subject.Contains("理综") || subject.Contains("文综"))
                            final = year + "年" + QX + subject.Substring(3) + "分类校数据统计分析报告.docx";
                        else
                            final = year + "年" + QX + subject + "分类校数据统计分析报告.docx";
                    }
                    else if (report_style.Equals("总体"))
                    {
                        if (subject.Contains("理综") || subject.Contains("文综"))
                            final = year + "年" + subject.Substring(3) + "数据统计分析报告(最终版）.docx";
                        else
                            final = year + "年" + subject + "数据统计分析报告(最终版）.docx";
                    }
                    
                }
            }
            final = addr + final;
            oDoc.SaveAs(final, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
            oDoc.Close(oMissing, oMissing, oMissing);
            oWord.Quit(oMissing, oMissing, oMissing);
        }
        public static void insertAddons(Microsoft.Office.Interop.Word._Document doc)
        {
            object oEndOfDoc = "\\endofdoc";
            object oMissing = System.Reflection.Missing.Value;
            
            Microsoft.Office.Interop.Word.Range range = doc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            range.InsertFile(@CurrentDirectory + @"\addon.doc", oMissing, false, false, false);
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
            char[] ans_ = choiceTransfer(ans).ToCharArray();
            foreach (char temp in ans_)
            {
                if (!choiceTransfer(da).Contains(temp))
                    return false;
            }
            return true;
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

        
    }
}
