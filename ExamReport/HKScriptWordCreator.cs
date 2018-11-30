using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Drawing;
using System.IO;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using ZedGraph;
using System.Windows.Forms;
using Microsoft.International.Converters.PinYinConverter;

namespace ExamReport
{
    class HKScriptWordCreator
    {
        Word._Application oWord;
        Word._Document oDoc;
        object oMissing = System.Reflection.Missing.Value;

        public string get_pinyin_of_name(string name)
        {
            string r = string.Empty;
            if (name.Length == 2)
            {
                ChineseChar familyname = new ChineseChar(name[0]);
                string fname = familyname.Pinyins[0].ToString().TrimEnd("0123456789".ToCharArray());

                ChineseChar lastname = new ChineseChar(name[1]);
                string lname = lastname.Pinyins[0].ToString().TrimEnd("0123456789".ToCharArray());

                r = lname + " " + fname;
            }

            if (name.Length == 3)
            {
                ChineseChar familyname = new ChineseChar(name[0]);
                string fname = familyname.Pinyins[0].ToString().TrimEnd("0123456789".ToCharArray());

                ChineseChar lastname1 = new ChineseChar(name[1]);
                string lname1 = lastname1.Pinyins[0].ToString().TrimEnd("0123456789".ToCharArray());

                ChineseChar lastname2 = new ChineseChar(name[2]);
                string lname2 = lastname2.Pinyins[0].ToString().TrimEnd("0123456789".ToCharArray());

                r = lname1 + lname2 + " " + fname;

            }
            return r;
        }

        public string insertHistGraph(string title, DataRow dr, DataTable group, string subject_en, string key, Dictionary<string, List<string>> group_dict, int start)
        {
            List<string> colnames = group_dict[key];
            DataTable dt = new DataTable();
            dt.Columns.Add("tz", typeof(string));
            dt.Columns.Add("num", typeof(decimal));

            for (int i = 0; i < colnames.Count; i++)
            {
                DataRow temp = dt.NewRow();
                temp[0] = colnames[i];
                temp[1] = dr["FZ" + (start + i + 1).ToString()];
                dt.Rows.Add(temp);
            }
            return DotNetCharting.CreateColumn_wh(dt, 360, 250, title + "分维度评价的百分位\nPercentile of " + subject_en + "\nevaluated in different Dimensions of the Candidate", false, 40);
            //ZedGraph.createSubDiffBar(data_list);

            //Word.Range dist_rng = oDoc.Bookmarks.get_Item(oEndOfDoc).Range;
            //dist_rng.Paste();
            //Utils.mutex_clipboard.ReleaseMutex();
            //dist_rng.InsertCaption(oWord.CaptionLabels["图"], title, oMissing, Word.WdCaptionPosition.wdCaptionPositionBelow, oMissing);
            //dist_rng.MoveEnd(Word.WdUnits.wdParagraph, 1);
            //dist_rng.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            //dist_rng = oDoc.Bookmarks.get_Item(oEndOfDoc).Range;
            //dist_rng.MoveStart(Word.WdUnits.wdParagraph, 1);
            //dist_rng.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            //dist_rng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            //dist_rng.InsertParagraphAfter();
        }
        public string insertZFHistGraph(DataRow dr)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("tz", typeof(string));
            dt.Columns.Add("num", typeof(decimal));

            for (int i = Utils.hk_subject.Length - 1; i >= 0; i-- )
            {
                string sub = Utils.hk_subject[i];
                DataRow temp = dt.NewRow();
                temp[0] = Utils.hk_lang_trans(sub) + Utils.hk_en_trans(Utils.hk_lang_trans(sub));
                temp[1] = dr[sub];
                dt.Rows.Add(temp);
            }
            return DotNetCharting.CreateColumn_wh(dt, 400, 300, "考生各学科成绩百分位\nPercentile of the Candidate's Scores in Each Subject", true, 15);
        }
        public string insertsubHistGraph(DataRow dr, DataTable group, string subject_en, string key, Dictionary<string, List<string>> group_dict, int start)
        {
            List<string> colnames = group_dict[key];
            DataTable dt = new DataTable();
            dt.Columns.Add("tz", typeof(string));
            dt.Columns.Add("num", typeof(decimal));

            for (int i = 0; i < colnames.Count; i++)
            {
                DataRow temp = dt.NewRow();
                temp[0] = colnames[i];
                temp[1] = dr["FZ" + (start + i + 1).ToString()];
                dt.Rows.Add(temp);
            }
            return DotNetCharting.CreateColumn_wh(dt, 360, 220, key + "百分位", true, 20);
            
        }
        public void create_word_zf(DataRow realdata, Configuration config, DataRow dr, DataTable group, Dictionary<string, List<string>> group_dict, DataRow basic_dr, String adr, string date, DataTable schools)
        {
            object filepath = @Utils.CurrentDirectory + @"\HKscript_zf_template.docx";
            oWord = new Word.Application();
            oWord.Visible = Utils.isVisible;
            oDoc = oWord.Documents.Add(ref filepath, ref oMissing,
            ref oMissing, ref oMissing);
            schools.PrimaryKey = new DataColumn[] { schools.Columns["xxdm"] };
            DataRow school = schools.Rows.Find((string)basic_dr["xxdm"]);
            string schoolname = "";
            string schoolen = "";
            if (school != null)
            {
                schoolen = (string)school["school_en"];
                schoolname = (string)school["school_nam"];
            }
            WriteIntoDocument("name", (string)basic_dr["xm"]);
            WriteIntoDocument("name_en", get_pinyin_of_name((string)basic_dr["xm"]));
            WriteIntoDocument("id", (string)basic_dr["sfzjh"]);
            WriteIntoDocument("id2", (string)basic_dr["sfzjh"]);
            //WriteIntoDocument("school", schoolname);
            //WriteIntoDocument("school_en", schoolen);
            WriteIntoDocument("year", config.year);
            WriteIntoDocument("year2", config.year);
            WriteIntoDocument("exam_id", dr["zkzh"].ToString().Trim());
            WriteIntoDocument("exam_id2", dr["zkzh"].ToString().Trim());
            WriteIntoDocument("date", date.Trim());
            WriteIntoDocument("date2", date.Trim());

            foreach (string sub in Utils.hk_subject)
            {
                string subject_name = Utils.hk_en_trans_dt(Utils.hk_lang_trans(sub));
                switch (realdata[subject_name].ToString().Trim())
                {
                    case "合格":
                        WriteIntoDocument(sub, "合格");
                        WriteIntoDocument(sub + "_en", "Pass");
                        break;
                    case "不合格":
                        WriteIntoDocument(sub, "不合格");
                        WriteIntoDocument(sub + "_en", "Fail");
                        break;
                    default:
                        WriteIntoDocument(sub, "");
                        WriteIntoDocument(sub + "_en", "");
                        break;
                }

                
            }
            string imagedir = insertZFHistGraph(dr);
            //Word.Range dist_rng = oDoc.Bookmarks.get_Item("figure1").Range;
            Word.InlineShape shape = oDoc.Bookmarks.get_Item("figure").Range.InlineShapes.AddPicture(imagedir);
            shape.Width = 360;
            shape.Height = 260;
            adr = adr + Path.DirectorySeparatorChar + dr["zkzh"].ToString().Trim();
            if (!Directory.Exists(adr))
                Directory.CreateDirectory(adr);
            string name = "合格考全科成绩_" + dr["zkzh"];
            string addr = adr + @"\" + name;
            object fileformat = Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatPDF;
            //object fileformat = Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatDocument;
            oDoc.SaveAs(addr, fileformat, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
            oDoc.Close(false, oMissing, oMissing);
            oWord.Quit(false, oMissing, oMissing);
        }
        public void create_word2(DataRow realdata, Configuration config, DataRow dr, DataTable group, Dictionary<string, List<string>> group_dict, DataRow basic_dr, String adr, string date, DataTable schools)
        {
            object filepath = @Utils.CurrentDirectory + @"\HKscript_subject_template.docx";
            oWord = new Word.Application();
            oWord.Visible = Utils.isVisible;
            oDoc = oWord.Documents.Add(ref filepath, ref oMissing,
            ref oMissing, ref oMissing);
            schools.PrimaryKey = new DataColumn[] { schools.Columns["xxdm"] };
            DataRow school = schools.Rows.Find((string)basic_dr["xxdm"]);
            string schoolname = "";
            string schoolen = "";
            if (school != null)
            {
                schoolen = (string)school["school_en"];
                schoolname = (string)school["school_nam"];
            }

            WriteIntoDocument("name", (string)basic_dr["xm"]);
            WriteIntoDocument("name_en", get_pinyin_of_name((string)basic_dr["xm"]));
            WriteIntoDocument("id", (string)basic_dr["sfzjh"]);
            WriteIntoDocument("id2", (string)basic_dr["sfzjh"]);
            //WriteIntoDocument("school", schoolname);
            //WriteIntoDocument("school_en", schoolen);
            WriteIntoDocument("year", config.year);
            WriteIntoDocument("year2", config.year);
            WriteIntoDocument("subject", config.subject);
            WriteIntoDocument("subject2", config.subject);
            WriteIntoDocument("subject_en", Utils.hk_en_trans(config.subject));
            WriteIntoDocument("subject_en2", Utils.hk_en_trans(config.subject));
            WriteIntoDocument("exam_id", dr["zkzh"].ToString().Trim());
            WriteIntoDocument("exam_id2", dr["zkzh"].ToString().Trim());
            WriteIntoDocument("date", date.Trim());
            WriteIntoDocument("date2", date.Trim());
            string subject_name = Utils.hk_en_trans_dt(config.subject);
            switch (realdata[subject_name].ToString().Trim())
            {
                case "合格":
                    WriteIntoDocument("mark", "合格");
                    WriteIntoDocument("mark" + "_en", "Pass");
                    break;
                case "不合格":
                    WriteIntoDocument("mark", "不合格");
                    WriteIntoDocument("mark" + "_en", "Fail");
                    break;
                default:
                    WriteIntoDocument("mark", "");
                    WriteIntoDocument("mark" + "_en", "");
                    break;
            }
            //WriteIntoDocument("mark", (decimal)dr["totalmark"] >= 60 ? "合格" : "不合格");
            //WriteIntoDocument("mark_en", (decimal)dr["totalmark"] >= 60 ? "Pass" : "Fail");

            //draw_bar(group, dr, config.subject, Utils.hk_en_trans(config.subject), adr);
            string imagedir = insertHistGraph(config.subject, dr, group, Utils.hk_en_trans(config.subject), group_dict.Keys.ElementAt(0), group_dict, 0);
            //Word.Range dist_rng = oDoc.Bookmarks.get_Item("figure1").Range;
            Word.InlineShape shape = oDoc.Bookmarks.get_Item("figure1").Range.InlineShapes.AddPicture(imagedir);
            shape.Width = 200;
            shape.Height = 140;
            int count1 = group_dict[group_dict.Keys.ElementAt(0)].Count;
            string imagedir2 = insertsubHistGraph(dr, group, Utils.hk_en_trans(config.subject), group_dict.Keys.ElementAt(1), group_dict, count1);
            Word.InlineShape shape2 = oDoc.Bookmarks.get_Item("figure2").Range.InlineShapes.AddPicture(imagedir2);
            //dist_rng.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            //dist_rng.Paste();
            int count2 = count1 + group_dict[group_dict.Keys.ElementAt(1)].Count;
            string imagedir3 = insertsubHistGraph(dr, group, Utils.hk_en_trans(config.subject), group_dict.Keys.ElementAt(2), group_dict, count2);
            Word.InlineShape shape3 = oDoc.Bookmarks.get_Item("figure3").Range.InlineShapes.AddPicture(imagedir3);

            adr = adr + Path.DirectorySeparatorChar + dr["zkzh"].ToString().Trim();
            if (!Directory.Exists(adr))
                Directory.CreateDirectory(adr);
            string name = "合格考成绩_" + config.subject+ "_" + dr["zkzh"];
            string addr = adr + @"\" + name;
            //object fileformat = Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatPDF;
            object fileformat = Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatPDF;
            oDoc.SaveAs(addr, fileformat, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
            oDoc.Close(false, oMissing, oMissing);
            oWord.Quit(false, oMissing, oMissing);

        }

        public void draw_bar(DataTable dt, DataRow basic, string subject, string subject_en, string adr)
        {
            int fontsize = 20;
            ZedGraphControl zgc = new ZedGraphControl();
            GraphPane myPane = zgc.GraphPane;

            zgc.Width = 531;
            zgc.Height = 271;
            //zgc.Height = 450;

            List<double[]> data = new List<double[]>();

            string[] xlabels = new string[dt.Rows.Count];
            int tz_name_max = 0;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                string tz_name = dt.Rows[i]["tz"].ToString().Trim();
                if (tz_name.Length > tz_name_max)
                    tz_name_max = tz_name.Length;
                //xlabels[i] = AxisTransfer(tz_name);
                xlabels[i] = tz_name;
            }
            //zgc.Height = 300 + tz_name_max * 15;
            AddData(basic, data, dt.Rows.Count);

            AddBar(data[0], "本人平均得分率", ref myPane, SymbolType.Diamond, Color.Red, 3);

            myPane.XAxis.Scale.TextLabels = xlabels;
            myPane.XAxis.Scale.FontSpec.Size = fontsize;
            //myPane.XAxis.Scale.FontSpec.Angle = 90;
            myPane.XAxis.Type = AxisType.Text;
            myPane.XAxis.Scale.Align = AlignP.Inside;
            //myPane.XAxis.Scale.AlignH = AlignH.Left;
            myPane.IsFontsScaled = true;
            myPane.XAxis.Title.Text = "";
            myPane.YAxis.Title.Text = AxisTransfer("得分率");
            myPane.YAxis.Title.FontSpec.Size = fontsize;
            myPane.YAxis.Title.FontSpec.Angle = 90;

            myPane.Title.Text = subject + "分维度评价的百分位\nPercentile of "+subject_en+"\nevaluated in different Dimensions of the Candidate";

            myPane.XAxis.Scale.Max = dt.Rows.Count + 1;
            myPane.XAxis.Scale.MajorStep = 1;
            myPane.YAxis.Scale.Max = 100;
            myPane.YAxis.Scale.MajorStep = 25;
            myPane.YAxis.Scale.Min = 0;
            myPane.YAxis.MinorTic.IsAllTics = false;
            myPane.XAxis.MinorTic.IsAllTics = false;
            myPane.YAxis.MajorTic.IsOpposite = false;
            myPane.XAxis.MajorTic.IsOpposite = false;

            myPane.Legend.IsVisible = false;
            //myPane.Legend.Position = LegendPos.BottomFlushLeft;
            //myPane.Legend.FontSpec.Size = 11;

            myPane.Title.IsVisible = true;
            myPane.Chart.Fill = new Fill(Color.White);
            zgc.AxisChange();
            //Bitmap sourceBitmap = new Bitmap(zgc.Width, zgc.Height, System.Drawing.Imaging.PixelFormat.Format48bppRgb);

            //Bitmap sourceBitmap = myPane.GetImage(zgc.Width, zgc.Height, 300);
            //myPane.GetImage().Save(adr + "a.jpg");
            //zgc.DrawToBitmap(sourceBitmap, new Rectangle(0, 0, zgc.Width, zgc.Height));
            //Thread.CurrentThread.SetApartmentState(ApartmentState.STA);
            Bitmap sourceBitmap = new Bitmap(zgc.Width, zgc.Height);
            zgc.DrawToBitmap(sourceBitmap, new Rectangle(0, 0, zgc.Width, zgc.Height));
            
            Clipboard.Clear();
            Clipboard.SetImage(sourceBitmap);
        }
        public void create_word(DataTable dt, DataRow dr, DataTable group, Dictionary<string, List<string>> group_dict, DataRow basic_dr, String adr, string date)
        {
            object filepath = @Utils.CurrentDirectory + @"\template.doc";
            oWord = new Word.Application();
            oWord.Visible = Utils.isVisible;
            oDoc = oWord.Documents.Add(ref filepath, ref oMissing,
            ref oMissing, ref oMissing);
            WriteIntoDocument("subject", Utils.subject);
            WriteIntoDocument("stu_id", dr["studentid"].ToString().Trim());
            WriteIntoDocument("date", date.Trim());
            WriteIntoDocument("totalmark", string.Format("{0:F1}", (decimal)dr["totalmark"]));
            WriteIntoDocument("name", basic_dr["studentname"].ToString().Trim());
            WriteIntoDocument("school", basic_dr["schoolname"].ToString().Trim());

            DataRow rankDR = null;
            DataRow totalDR = dt.Rows.Find("total");
            string rank_name = "";
            switch ((string)dr["rank"])
            {
                case "excellent":
                    WriteIntoDocument("rank", "优秀");
                    rankDR = dt.Rows.Find("excellent");
                    rank_name = "优秀等级";
                    break;
                case "well":
                    WriteIntoDocument("rank", "良好");
                    rankDR = dt.Rows.Find("well");
                    rank_name = "良好等级";
                    break;
                case "pass":
                    WriteIntoDocument("rank", "合格");
                    rankDR = dt.Rows.Find("pass");
                    rank_name = "合格等级";
                    break;
                case "fail":
                    WriteIntoDocument("rank", "不合格");
                    rankDR = dt.Rows.Find("fail");
                    rank_name = "不合格等级";
                    break;
                default:
                    break;
            }

            WriteIntoDocument("PR", string.Format("{0:F1}", (decimal)dr["PR_total"]));

            Word.Table table = oDoc.Tables[1];

            table.Cell(2, 3).Range.Text = rank_name;
            int previousRow = 3;
            int line = 0, count = 1;

            List<string> strong = new List<string>();
            List<string> weak = new List<string>();
            List<string> bad = new List<string>();
            List<string> average = new List<string>();
            bool once = true;
            float row_height = 0;
            foreach (string key in group_dict.Keys)
            {
                //table.Cell(line + 2, 3).Range.Rows.Add(oMissing);

                //for (int j = 1; j < 6; j++)
                //{
                //    table.Cell(line + 3, j).Range.Font.Bold = 0;
                //    table.Cell(line + 3, j).Range.Font.Size = 10;
                //    table.Cell(line + 3, j).Range.Shading.BackgroundPatternColor = table.Cell(1, 1).Range.Shading.BackgroundPatternColor;
                //    table.Cell(line + 3, j).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                //    table.Cell(line + 3, j).SetHeight(0.48f, Word.WdRowHeightRule.wdRowHeightAtLeast);
                //}

                //table.Cell(line + 3, 1).Range.Text = key;
                //table.Cell(line + 3, 1).Range.Font.Bold = 1;
                //table.Cell(line + 3, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                //previousRow = line + 3;
                //for (int j = 2; j < 6; j++)
                //{
                //    table.Cell(line + 3, j).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                //    table.Cell(line + 3, j).Range.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray10;
                //}

                //line++;
                bool first = true;
                List<string> names = group_dict[key];
                int name_count = names.Count;
                foreach (string group_name in names)
                {
                    table.Cell(line + 2, 3).Range.Rows.Add(oMissing);
                    if (first)
                    {

                        if (once)
                        {
                            table.Cell(line + 3, 1).Split(1, 2);
                            float temp_wid = table.Cell(line + 3, 1).Width;
                            table.Cell(line + 3, 1).Width = temp_wid / 2;
                            table.Cell(line + 3, 2).Width += temp_wid / 2;
                            once = false;
                        }
                        table.Cell(line + 3, 1).Range.Text = AxisTransfer(key);
                        table.Cell(line + 3, 1).Range.Font.Bold = 1;
                        table.Cell(line + 3, 1).Range.Font.Size = 10;
                        table.Cell(line + 3, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        table.Cell(line + 3, 1).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                        previousRow = line + 3;
                        first = false;

                        int char_count = key.ToCharArray().Length;
                        if (char_count > name_count)
                            row_height = (char_count * 16f) / name_count;
                        else
                            row_height = 16f;
                    }
                    else
                    {
                        table.Cell(previousRow, 1).Merge(table.Cell(line + 3, 1));

                    }
                    for (int j = 2; j < 7; j++)
                    {
                        table.Cell(line + 3, j).Range.Font.Bold = 0;
                        table.Cell(line + 3, j).Range.Font.Size = 10;
                        table.Cell(line + 3, j).Range.Shading.BackgroundPatternColor = table.Cell(1, 1).Range.Shading.BackgroundPatternColor;
                        table.Cell(line + 3, j).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        table.Cell(line + 3, j).SetHeight(row_height, Word.WdRowHeightRule.wdRowHeightExactly);
                    }
                    table.Cell(line + 3, 2).Range.Text = group_name;

                    table.Cell(line + 3, 3).Range.Text = string.Format("{0:F2}", (decimal)dr["FZ" + count.ToString()]);
                    table.Cell(line + 3, 4).Range.Text = string.Format("{0:F2}", (decimal)rankDR["FZ" + count.ToString()]);
                    table.Cell(line + 3, 5).Range.Text = string.Format("{0:F2}", (decimal)totalDR["FZ" + count.ToString()]);
                    table.Cell(line + 3, 6).Range.Text = string.Format("{0:F1}", (decimal)dr["PR" + count.ToString()]);




                    decimal score = (decimal)dr["FZ" + count.ToString()] - (decimal)rankDR["FZ" + count.ToString()];
                    if (score >= 0.05m)
                        strong.Add(group_name);
                    else if (score <= -0.01m && score > -0.05m)
                        weak.Add(group_name);
                    else if (score <= -0.05m)
                        bad.Add(group_name);
                    else if (score >= 0.01m && score < 0.05m)
                        average.Add(group_name);

                    line++;
                    count++;
                }

            }

            //draw(group, dr, rankDR, totalDR);
            draw_horizontal(group, dr, rankDR, totalDR);
            Word.Range dist_rng = oDoc.Bookmarks.get_Item("pic").Range;
            dist_rng.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            dist_rng.Paste();

            if (average.Count == group.Rows.Count)
                WriteIntoDocument("comment", "您" + Utils.subject + "学科的高中会考成绩为" + rank_name + "。与您相同等级的学生群体相比，您对本学科应掌握的知识点和能力点都达到本等级群体的平均水平。\n");
            else
            {
                StringBuilder sb = new StringBuilder("您" + Utils.subject + "学科的高中会考成绩为" + rank_name + "。与您相同等级的学生群体相比，您在");

                for (int i = 0; i < strong.Count; i++)
                {
                    sb.Append(strong[i]);
                    if (i != strong.Count - 1)
                        sb.Append("、");
                }
                if (strong.Count != 0)
                {
                    sb.Append("方面较强，希望保持您的优势");
                    if (weak.Count == 0 && bad.Count == 0)
                        sb.Append("。");
                    else
                        sb.Append("；");
                }
                for (int i = 0; i < weak.Count; i++)
                {
                    sb.Append(weak[i]);
                    if (i != weak.Count - 1)
                        sb.Append("、");
                }
                if (weak.Count != 0)
                {
                    sb.Append("方面较弱，有待进一步提升");
                    if (bad.Count == 0)
                        sb.Append("。");
                    else
                        sb.Append("；");
                }
                for (int i = 0; i < bad.Count; i++)
                {
                    sb.Append(bad[i]);
                    if (i != bad.Count - 1)
                        sb.Append("、");
                }
                if (bad.Count != 0)
                    sb.Append("方面薄弱，需要找出原因，加强这方面的学习和提高。");
                sb.Append("\n");
                WriteIntoDocument("comment", sb.ToString());
            }

            //Word.Table count_table = oDoc.Tables[2];
            //count_table.Cell(2, 2).Range.Text = getPercent((int)dt.Rows[1]["num"], (int)dt.Rows[0]["num"]);
            //count_table.Cell(2, 3).Range.Text = getPercent((int)dt.Rows[2]["num"], (int)dt.Rows[0]["num"]);
            //count_table.Cell(2, 4).Range.Text = getPercent((int)dt.Rows[3]["num"], (int)dt.Rows[0]["num"]);
            //count_table.Cell(2, 5).Range.Text = getPercent((int)dt.Rows[4]["num"], (int)dt.Rows[0]["num"]);

            Word.Table TH_table = oDoc.Tables[2];
            previousRow = 1;
            line = 0;
            count = 0;
            once = true;
            row_height = 0;
            foreach (string key in group_dict.Keys)
            {
                //for (int j = 1; j < 3; j++)
                //{
                //    TH_table.Cell(line + 2, j).Range.Font.Bold = 0;
                //    TH_table.Cell(line + 2, j).Range.Font.Size = 10;
                //    TH_table.Cell(line + 2, j).Range.Shading.BackgroundPatternColor = table.Cell(1, 1).Range.Shading.BackgroundPatternColor;
                //    TH_table.Cell(line + 2, j).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                //    TH_table.Cell(line + 2, j).SetHeight(0.48f, Word.WdRowHeightRule.wdRowHeightAtLeast);
                //}
                //TH_table.Cell(line + 2, 1).Range.Text = key;

                //previousRow = line + 2;
                //TH_table.Cell(line + 2, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                //TH_table.Cell(line + 2, 1).Range.Font.Bold = 1;
                //TH_table.Cell(line + 2, 2).Range.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray10;

                //if (count < group.Rows.Count - 1)
                //    TH_table.Cell(line + 2, 2).Range.Rows.Add(oMissing);
                //line++;
                bool first = true;
                List<string> names = group_dict[key];
                int name_count = names.Count;
                foreach (string group_name in names)
                {
                    //TH_table.Cell(line + 1, 2).Range.Rows.Add(oMissing);
                    if (first)
                    {
                        if (once)
                        {
                            TH_table.Cell(line + 2, 1).Split(1, 2);
                            float temp_wid = TH_table.Cell(line + 2, 1).Width;
                            TH_table.Cell(line + 2, 1).Width = temp_wid / 2;
                            TH_table.Cell(line + 2, 2).Width += temp_wid / 2;
                            once = false;
                        }
                        TH_table.Cell(line + 2, 1).Range.Text = AxisTransfer(key);
                        TH_table.Cell(line + 2, 1).Range.Font.Bold = 1;
                        TH_table.Cell(line + 2, 1).Range.Font.Size = 10;
                        TH_table.Cell(line + 2, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        TH_table.Cell(line + 2, 1).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                        previousRow = line + 2;
                        first = false;

                        int char_count = key.ToCharArray().Length;
                        if (char_count > name_count)
                            row_height = (char_count * 16f) / name_count;
                        else
                            row_height = 16f;
                    }
                    else
                        TH_table.Cell(previousRow, 1).Merge(TH_table.Cell(line + 2, 1));
                    for (int j = 2; j < 4; j++)
                    {
                        TH_table.Cell(line + 2, j).Range.Font.Bold = 0;
                        TH_table.Cell(line + 2, j).Range.Font.Size = 10;
                        TH_table.Cell(line + 2, j).Range.Shading.BackgroundPatternColor = table.Cell(1, 1).Range.Shading.BackgroundPatternColor;
                        TH_table.Cell(line + 2, j).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        TH_table.Cell(line + 2, j).SetHeight(row_height, Word.WdRowHeightRule.wdRowHeightExactly);
                    }
                    TH_table.Cell(line + 2, 2).Range.Text = group_name;
                    TH_table.Cell(line + 2, 3).Range.Text = getTH(group.Rows[count]["th"].ToString().Trim());
                    TH_table.Cell(line + 2, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;


                    if (count < group.Rows.Count - 1)
                        TH_table.Cell(line + 2, 2).Range.Rows.Add(oMissing);

                    line++;
                    count++;
                }
            }
            string name = "会考成绩_" + dr["studentid"];
            string addr = adr + @"\" + name;
            object fileformat = Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatPDF;
            oDoc.SaveAs(addr, fileformat, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
            oDoc.Close(false, oMissing, oMissing);
            oWord.Quit(false, oMissing, oMissing);
        }
        public string getPercent(int num, int total)
        {
            return string.Format("{0:F2}", num / Convert.ToDouble(total) * 100);
        }
        public string getTH(string th)
        {
            string[] th_string = th.ToString().Trim().Split(new char[2] { ',', '，' });
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < th_string.Length; i++)
            {
                if (System.Text.RegularExpressions.Regex.IsMatch(th_string[i], "^\\d+~\\d+$"))
                {
                    string[] num = th_string[i].Split('~');
                    sb.Append(num[0]);
                    sb.Append("～");
                    sb.Append(num[1]);
                }
                else
                    sb.Append(th_string[i]);
                if (i < th_string.Length - 1)
                    sb.Append("、");
            }
            return sb.ToString();
        }
        public void WriteIntoDocument(string BookmarkName, string FillName)
        {
            object bookmarkName = BookmarkName;
            Microsoft.Office.Interop.Word.Bookmark bm = oDoc.Bookmarks.get_Item(ref bookmarkName);//返回书签 
            bm.Range.Text = FillName;//设置书签域的内容
        }

        public void testCase()
        {
            object filepath = @"C:\Users\sooner\Documents\Visual Studio 2010\Projects\new2\WindowsFormsApplication1\WindowsFormsApplication1\bin\Debug\template.doc";
            oWord = new Word.Application();
            oWord.Visible = true;
            oDoc = oWord.Documents.Add(ref filepath, ref oMissing,
            ref oMissing, ref oMissing);
            Word.Table table = oDoc.Tables[2];

            //object beforerow = table.Rows[1];

            //table.Rows.Add(beforerow);
            table.Cell(2, 1).Range.Text = "听力";
            table.Cell(2, 2).Range.Text = "写作";
            //table.Cell(2, 3).Range.Text = "阅读";
            table.Cell(2, 2).Range.Rows.Add(oMissing);

        }
        public void draw_horizontal(DataTable dt, DataRow basic, DataRow rank, DataRow total)
        {
            ZedGraphControl zgc = new ZedGraphControl();
            GraphPane myPane = zgc.GraphPane;
            zgc.Width = 680;
            zgc.Height = 600;

            List<double[]> data = new List<double[]>();
            string[] ylabels = new string[dt.Rows.Count];
            for (int i = 0; i < dt.Rows.Count; i++)
                ylabels[i] = dt.Rows[dt.Rows.Count - i - 1]["tz"].ToString().Trim();

            AddHorizontalData(basic, data, dt.Rows.Count);
            AddHorizontalData(rank, data, dt.Rows.Count);
            AddHorizontalData(total, data, dt.Rows.Count);

            AddHorizontalBar(data[0], "本人平均得分率", ref myPane, SymbolType.Diamond, Color.Red, 3);
            AddHorizontalCurve(data[1], "同等级平均得分率", ref myPane, SymbolType.Square, Color.Red, 2);
            AddHorizontalCurve(data[2], "总体平均得分率", ref myPane, SymbolType.Triangle, Color.DarkBlue, 1);

            myPane.CurveList.Sort(new CurveItemTagComparer());
            zgc.Refresh();

            myPane.YAxis.Scale.TextLabels = ylabels;
            myPane.YAxis.Scale.FontSpec.Size = 11;
            myPane.YAxis.Title.Text = "";
            myPane.YAxis.Type = AxisType.Text;
            myPane.YAxis.Scale.Align = AlignP.Inside;

            myPane.XAxis.Title.Text = "得分率";
            myPane.XAxis.Title.FontSpec.Size = 11;
            myPane.XAxis.Scale.FontSpec.Size = 11;
            myPane.Title.Text = "";

            myPane.YAxis.Scale.Max = dt.Rows.Count + 1;
            myPane.YAxis.Scale.MajorStep = 1;
            myPane.XAxis.Scale.Max = 1.0;
            myPane.XAxis.Scale.MajorStep = 0.5;
            myPane.XAxis.Scale.Min = 0;
            myPane.YAxis.MinorTic.IsAllTics = false;
            myPane.XAxis.MinorTic.IsAllTics = false;
            myPane.YAxis.MajorTic.IsOpposite = false;
            myPane.XAxis.MajorTic.IsOpposite = false;

            myPane.Legend.IsVisible = true;
            myPane.Legend.Position = LegendPos.Bottom;
            myPane.Legend.FontSpec.Size = 9;

            myPane.Title.IsVisible = true;
            myPane.Chart.Fill = new Fill(Color.White);
            zgc.AxisChange();
            Bitmap sourceBitmap = new Bitmap(zgc.Width, zgc.Height, System.Drawing.Imaging.PixelFormat.Format48bppRgb);
            zgc.DrawToBitmap(sourceBitmap, new Rectangle(0, 0, zgc.Width, zgc.Height));
            //Thread.CurrentThread.SetApartmentState(ApartmentState.STA);
            Clipboard.Clear();
            Clipboard.SetImage(sourceBitmap);
        }
        public void draw(DataTable dt, DataRow basic, DataRow rank, DataRow total)
        {
            ZedGraphControl zgc = new ZedGraphControl();
            GraphPane myPane = zgc.GraphPane;

            zgc.Width = 590;
            //zgc.Height = 450;

            List<double[]> data = new List<double[]>();

            string[] xlabels = new string[dt.Rows.Count];
            int tz_name_max = 0;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                string tz_name = dt.Rows[i]["tz"].ToString().Trim();
                if (tz_name.Length > tz_name_max)
                    tz_name_max = tz_name.Length;
                xlabels[i] = AxisTransfer(tz_name);
            }
            zgc.Height = 300 + tz_name_max * 15;
            AddData(basic, data, dt.Rows.Count);
            AddData(rank, data, dt.Rows.Count);
            AddData(total, data, dt.Rows.Count);

            AddBar(data[0], "本人平均得分率", ref myPane, SymbolType.Diamond, Color.Red, 3);
            AddCurve(data[1], "同等级平均得分率", ref myPane, SymbolType.Square, Color.Red, 2);
            AddCurve(data[2], "总体平均得分率", ref myPane, SymbolType.Triangle, Color.DarkGreen, 1);

            myPane.CurveList.Sort(new CurveItemTagComparer());
            zgc.Refresh();
            myPane.XAxis.Scale.TextLabels = xlabels;
            myPane.XAxis.Scale.FontSpec.Size = 12;
            //myPane.XAxis.Scale.FontSpec.Angle = 90;
            myPane.XAxis.Type = AxisType.Text;
            myPane.XAxis.Scale.Align = AlignP.Inside;
            //myPane.XAxis.Scale.AlignH = AlignH.Left;
            myPane.IsFontsScaled = true;
            myPane.XAxis.Title.Text = "";
            myPane.YAxis.Title.Text = AxisTransfer("得分率");
            myPane.YAxis.Title.FontSpec.Size = 12;
            myPane.YAxis.Title.FontSpec.Angle = 90;

            myPane.Title.Text = "";

            myPane.XAxis.Scale.Max = dt.Rows.Count + 1;
            myPane.XAxis.Scale.MajorStep = 1;
            myPane.YAxis.Scale.Max = 1.0;
            myPane.YAxis.Scale.MajorStep = 0.5;
            myPane.YAxis.Scale.Min = 0;
            myPane.YAxis.MinorTic.IsAllTics = false;
            myPane.XAxis.MinorTic.IsAllTics = false;
            myPane.YAxis.MajorTic.IsOpposite = false;
            myPane.XAxis.MajorTic.IsOpposite = false;

            myPane.Legend.IsVisible = true;
            myPane.Legend.Position = LegendPos.BottomFlushLeft;
            myPane.Legend.FontSpec.Size = 11;

            myPane.Title.IsVisible = true;
            myPane.Chart.Fill = new Fill(Color.White);
            zgc.AxisChange();
            //Bitmap sourceBitmap = new Bitmap(zgc.Width, zgc.Height, System.Drawing.Imaging.PixelFormat.Format48bppRgb);

            Bitmap sourceBitmap = myPane.GetImage(zgc.Width, zgc.Height, 1000);
            //zgc.DrawToBitmap(sourceBitmap, new Rectangle(0, 0, zgc.Width, zgc.Height));
            //Thread.CurrentThread.SetApartmentState(ApartmentState.STA);
            Clipboard.Clear();
            Clipboard.SetImage(sourceBitmap);
        }
        class CurveItemTagComparer : IComparer<CurveItem>
        {
            public int Compare(CurveItem x, CurveItem y)
            {
                return ((int)x.Tag).CompareTo((int)y.Tag);
            }
        }
        public void AddData(DataRow dr, List<double[]> list, int count)
        {
            double[] temp = new double[count];
            for (int i = 0; i < count; i++)
            {
                temp[i] = Convert.ToDouble(dr["FZ" + (i + 1).ToString()]);
            }

            list.Add(temp);

        }
        public void AddHorizontalData(DataRow dr, List<double[]> list, int count)
        {
            double[] temp = new double[count];
            for (int i = 0; i < count; i++)
            {
                temp[count - i - 1] = Convert.ToDouble(dr["FZ" + (i + 1).ToString()]);
            }

            list.Add(temp);

        }

        public void AddCurve(double[] data, string name, ref GraphPane pane, SymbolType type, Color color, int tag)
        {
            LineItem myCurve = pane.AddCurve(name, null, data, color, type);
            myCurve.Line.IsVisible = false;
            myCurve.Symbol.Size = 10;
            myCurve.Symbol.Fill = new Fill(color);

            myCurve.Tag = tag;
        }

        public void AddHorizontalCurve(double[] data, string name, ref GraphPane pane, SymbolType type, Color color, int tag)
        {
            double[] ydata = new double[data.Length];
            for (int i = 0; i < data.Length; i++)
                ydata[i] = i + 1;
            LineItem myCurve = pane.AddCurve(name, data, ydata, color, type);

            myCurve.Line.IsVisible = false;
            myCurve.Symbol.Size = 9;
            myCurve.Symbol.Fill = new Fill(color);

            myCurve.Tag = tag;
        }
        public void AddBar(double[] data, string name, ref GraphPane pane, SymbolType type, Color color, int tag)
        {
            double[] ydata = new double[data.Length];
            for (int i = 0; i < data.Length; i++)
                ydata[i] = i + 1;
            PointPairList ppBar = new PointPairList(ydata, data);
            BarItem myCurve1 = pane.AddBar(name, ppBar, color);
            myCurve1.Bar.Fill = new Fill(Color.FromArgb(0, 255, 255), Color.FromArgb(0, 255, 255));
            myCurve1.Tag = tag;
        }
        public void AddHorizontalBar(double[] data, string name, ref GraphPane pane, SymbolType type, Color color, int tag)
        {
            double[] ydata = new double[data.Length];
            for (int i = 0; i < data.Length; i++)
                ydata[i] = i + 1;
            //LineItem myCurve = pane.AddCurve(name, data, ydata, color, type);
            pane.BarSettings.Base = BarBase.Y;
            PointPairList ppBar = new PointPairList(data, ydata);
            BarItem myCurve1 = pane.AddBar(name, ppBar, color);
            myCurve1.Bar.Fill = new Fill(Color.FromArgb(0, 255, 255), Color.FromArgb(0, 255, 255));
            myCurve1.Tag = tag;
            //myCurve.Line.IsVisible = false;
            //myCurve.Symbol.Size = 12;
            //myCurve.Symbol.Fill = new Fill(color);
        }


        public static string AxisTransfer(string name)
        {
            char[] names = name.ToCharArray();
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < names.Length; i++)
            {
                if (names[i].Equals('（') || names[i].Equals('('))
                {
                    while (i < names.Length)
                    {
                        if (names[i].Equals('）') || names[i].Equals(')'))
                        {
                            sb.Append(names[i]);
                            break;
                        }

                        sb.Append(name[i]);
                        i++;
                    }
                    if (i != names.Length - 1)
                        sb.Append("\n");
                }
                else
                {
                    sb.Append(names[i]);
                    if (i != names.Length - 1)
                        sb.Append("\n");
                }

            }
            return sb.ToString();
        }
    }
}
