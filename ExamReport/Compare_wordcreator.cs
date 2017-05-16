using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Collections;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using Graph = Microsoft.Office.Interop.Graph;

namespace ExamReport
{
    public class Compare_wordcreator
    {
        private object CaptionTitle = "CaptionTitle";
        private object ExamTitle0 = "ExamTitle0";
        private object ExamTitle1 = "ExamTitle1";
        private object ExamTitle2 = "ExamTitle2";
        private object ExamTitle3 = "ExamTitle3";
        private object ExamBodyText = "ExamBodyText";
        private object TableContent = "TableContent";
        private object TableContent2 = "TableContent2";
        object oPagebreak = Microsoft.Office.Interop.Word.WdBreakType.wdPageBreak;
        object oMissing = System.Reflection.Missing.Value;
        object oEndOfDoc = "\\endofdoc"; /* \endofdoc is a predefined bookmark */
        Word._Application oWord;
        Word._Document oDoc;
        ArrayList _sdata;
        object oParagrahbreak = Microsoft.Office.Interop.Word.WdBreakType.wdLineBreak;
        Object oTrue = true;
        Object oFalse = false;

        private Configuration _config;
        DataTable _groups;
        object oClassType = "Excel.Chart.8";
        string _addr;

        public string year1;
        public string year2;
        Dictionary<string, WordData> year1_data = new Dictionary<string, WordData>();
        Dictionary<string, WordData> year2_data = new Dictionary<string, WordData>();
        public Partition_wordcreator.ChartCombine year1_comb;
        public Partition_wordcreator.ChartCombine year2_comb;
        public DataTable summary;

        Dictionary<string, string> subname = new Dictionary<string, string> {
            {"wk","文科"},{"yww","语文（文）"},{"sxw","数学（文）"}, {"yyw","英语（文）"},
            {"wz","文科综合"},{"ls","历史"},{"dl","地理"},{"zz","政治"},
            {"lz","理科综合"},{"ywl","语文（理）"},{"sxl","数学（理）"}, {"yyl","英语（理）"},
            {"wl","物理"},{"hx","化学"},{"sw","生物"},
            {"yw","语文"},{"yy","英语"}
        };

        public void creating_word()
        {
            string subject = _config.subject;
            object filepath = @_config.CurrentDirectory + @"\template2.dotx";
            //Start Word and create a new document.
            _addr = _config.save_address + @"\" + subject + ".docx";
            oWord = new Word.Application();

            oWord.Visible = _config.isVisible;
            oDoc = oWord.Documents.Add(ref filepath, ref oMissing,
            ref oMissing, ref oMissing);
            Utils.WriteFrontPage(_config, oDoc);

            insertText(ExamTitle0, "  "+ year1 + "、" + year2 +"年高考北京卷总分分布曲线图");
            insertChart("    "+ year1 + "总分分布曲线图", year1_comb.target, "分数", "比率", Excel.XlChartType.xlLineMarkers, 750);
            insertChart("    " + year2 + "总分分布曲线图", year2_comb.target, "分数", "比率", Excel.XlChartType.xlLineMarkers, 750);

            insertText(ExamTitle0, "  " + year1 + "、" + year2 + "年高考北京卷整体对比分析");
            insertTotalTable("    试卷（文理科）总分分析表");

            insertText(ExamTitle0, "  " + year2 + "年数据分析（简缩版）发现的主要学科问题");
            oDoc.Characters.Last.InsertBreak(oPagebreak);
            insertText(ExamTitle0, "  " + year1 + "、" + year2 + "年分学科数据分析（简缩版）");
            foreach (string sub in year1_data.Keys)
            {
                int rowcount;
                string grouptitle;
                string sub_cn = subname[sub];
                insertText(ExamTitle1, sub_cn + "学科数据分析（简缩版）");
                insertTotalTable("    " + year2 + "年" + sub_cn + "总分分析表", year2_data[sub]);
                insertTotalChart("    " + year1 + "年" + sub_cn + "总分分布曲线图", year1_data[sub]);
                insertTotalChart("    " + year2 + "年" + sub_cn + "总分分布曲线图", year2_data[sub]);
                if (sub.Equals("wz") || sub.Equals("lz"))
                {
                    rowcount = 4;
                    grouptitle = "    " + year2 + "年" + sub_cn + "科目整体分析表";
                }
                else
                {
                    grouptitle = "    " + year2 + "年" + sub_cn + "题组整体分析表";
                    rowcount = year2_data[sub].group_analysis.Rows.Count + 1;
                    insertTotalDifficultyChart("    " + year1 + "年" + sub_cn + "题目难度与区分度坐标图", year1_data[sub]);
                    insertTotalDifficultyChart("    " + year2 + "年" + sub_cn + "题目难度与区分度坐标图", year2_data[sub]);
                    insertSingleTable("    " + year2 + "年" + sub_cn + "题目整体分析表", year2_data[sub]);
                }
                insertTotalGroupTable(grouptitle, year2_data[sub], rowcount);
            }

            foreach (Word.TableOfContents table in oDoc.TablesOfContents)
                table.Update();
            Utils.Save(_config, oDoc, oWord);




        }
        public void insertTotalGroupTable(string title, WordData sdata, int rowcount)
        {
            Word.Table groups_table;
            Word.Range groups_range = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;

            groups_table = oDoc.Tables.Add(groups_range, rowcount, 10, ref oMissing, oTrue);
            groups_table.Rows[1].HeadingFormat = -1;
            groups_table.Range.InsertCaption(oWord.CaptionLabels["表"], title, oMissing, Word.WdCaptionPosition.wdCaptionPositionAbove, oMissing);
            groups_range.MoveEnd(Word.WdUnits.wdParagraph, 1);
            groups_range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            //groups_table.Range.set_Style(ref TableContent2);
            //groups_table.Range.Select();
            //oWord.Selection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            //oWord.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            //groups_table.Range.ParagraphFormat.Space1();
            groups_table.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            groups_table.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;

            //groups_table.Range.Paragraphs.OutlineLevel = Word.WdOutlineLevel.wdOutlineLevelBodyText;
            //groups_table.Range.Font.Size = 10;
            //groups_table.Range.Font.Name = "黑体";

            groups_table.Cell(1, 1).Range.Text = "题目";
            groups_table.Cell(1, 2).Range.Text = "满分值";
            groups_table.Cell(1, 3).Range.Text = "最大值";
            groups_table.Cell(1, 4).Range.Text = "最小值";
            groups_table.Cell(1, 5).Range.Text = "平均值";
            groups_table.Cell(1, 6).Range.Text = "标准差";
            groups_table.Cell(1, 7).Range.Text = "差异系数";
            groups_table.Cell(1, 8).Range.Text = "难度";
            groups_table.Cell(1, 9).Range.Text = "相关系数";
            groups_table.Cell(1, 10).Range.Text = "鉴别指数";

            for (int i = 0; i < rowcount - 1; i++)
            {
                groups_table.Cell(i + 2, 1).Range.Text = sdata._groups_ans.Rows[i][0].ToString().Trim();
                groups_table.Cell(i + 2, 2).Range.Text = FullmarkFormat((decimal)sdata.group_analysis.Rows[i]["fullmark"]);
                groups_table.Cell(i + 2, 3).Range.Text = string.Format("{0:F1}", sdata.group_analysis.Rows[i]["max"]);
                groups_table.Cell(i + 2, 4).Range.Text = string.Format("{0:F1}", sdata.group_analysis.Rows[i]["min"]);
                groups_table.Cell(i + 2, 5).Range.Text = string.Format("{0:F2}", sdata.group_analysis.Rows[i]["avg"]);
                groups_table.Cell(i + 2, 6).Range.Text = string.Format("{0:F2}", sdata.group_analysis.Rows[i]["standardErr"]);
                groups_table.Cell(i + 2, 7).Range.Text = string.Format("{0:F2}", sdata.group_analysis.Rows[i]["dfactor"]);
                groups_table.Cell(i + 2, 8).Range.Text = string.Format("{0:F2}", sdata.group_analysis.Rows[i]["difficulty"]);
                groups_table.Cell(i + 2, 9).Range.Text = string.Format("{0:F2}", sdata.group_analysis.Rows[i]["correlation"]);
                groups_table.Cell(i + 2, 10).Range.Text = string.Format("{0:F2}", sdata.group_analysis.Rows[i]["discriminant"]);
                if (Math.Abs((decimal)sdata.group_analysis.Rows[i]["correlation"]) != (decimal)sdata.group_analysis.Rows[i]["correlation"] ||
                    Math.Abs((decimal)sdata.group_analysis.Rows[i]["discriminant"]) != (decimal)sdata.group_analysis.Rows[i]["discriminant"])
                    for (int j = 1; j < 11; j++)
                        groups_table.Cell(i + 2, j).Range.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray10;
            }
            groups_table.Select();
            oWord.Selection.set_Style(ref TableContent2);

            groups_range = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            groups_range.InsertParagraphAfter();
        }
        string FullmarkFormat(decimal remark)
        {
            return Math.Ceiling(Convert.ToDouble(remark)) == Convert.ToDouble(remark) ? Convert.ToInt32(remark).ToString() : string.Format("{0:F1}", remark);
        }
        public void insertSingleTable(string title, WordData sdata)
        {
            Word.Table topic_table;
            Word.Range topic_Rng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;

            topic_table = oDoc.Tables.Add(topic_Rng, sdata.total_analysis.Rows.Count + 1, 10, ref oMissing, oTrue);
            topic_table.Rows[1].HeadingFormat = -1;
            
            topic_table.Range.InsertCaption(oWord.CaptionLabels["表"], title, oMissing, Word.WdCaptionPosition.wdCaptionPositionAbove, oMissing);
            topic_Rng.MoveEnd(Word.WdUnits.wdParagraph, 1);
            topic_Rng.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

            //topic_table.Range.Select();
            //oWord.Selection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            //oWord.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            //topic_table.Range.ParagraphFormat.Space1();
            topic_table.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            topic_table.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;

            //topic_table.Range.Paragraphs.OutlineLevel = Word.WdOutlineLevel.wdOutlineLevelBodyText;
            //topic_table.Range.Font.Size = 10;
            //topic_table.Range.Font.Name = "黑体";

            topic_table.Cell(1, 1).Range.Text = "题目";
            topic_table.Cell(1, 2).Range.Text = "满分值";
            topic_table.Cell(1, 3).Range.Text = "最大值";
            topic_table.Cell(1, 4).Range.Text = "最小值";
            topic_table.Cell(1, 5).Range.Text = "平均值";
            topic_table.Cell(1, 6).Range.Text = "标准差";
            topic_table.Cell(1, 7).Range.Text = "差异系数";
            topic_table.Cell(1, 8).Range.Text = "难度";
            topic_table.Cell(1, 9).Range.Text = "相关系数";
            topic_table.Cell(1, 10).Range.Text = "鉴别指数";

            for (int i = 0; i < sdata.total_analysis.Rows.Count; i++)
            {
                topic_table.Cell(i + 2, 1).Range.Text = sdata.total_analysis.Rows[i]["number"].ToString().Substring(1);
                topic_table.Cell(i + 2, 2).Range.Text = FullmarkFormat((decimal)sdata.total_analysis.Rows[i]["fullmark"]);
                topic_table.Cell(i + 2, 3).Range.Text = string.Format("{0:F1}", sdata.total_analysis.Rows[i]["max"]);
                topic_table.Cell(i + 2, 4).Range.Text = string.Format("{0:F1}", sdata.total_analysis.Rows[i]["min"]);
                topic_table.Cell(i + 2, 5).Range.Text = string.Format("{0:F2}", sdata.total_analysis.Rows[i]["avg"]);
                topic_table.Cell(i + 2, 6).Range.Text = string.Format("{0:F2}", sdata.total_analysis.Rows[i]["standardErr"]);
                topic_table.Cell(i + 2, 7).Range.Text = string.Format("{0:F2}", sdata.total_analysis.Rows[i]["dfactor"]);
                topic_table.Cell(i + 2, 8).Range.Text = string.Format("{0:F2}", sdata.total_analysis.Rows[i]["difficulty"]);
                topic_table.Cell(i + 2, 9).Range.Text = string.Format("{0:F2}", sdata.total_analysis.Rows[i]["correlation"]);
                topic_table.Cell(i + 2, 10).Range.Text = string.Format("{0:F2}", sdata.total_analysis.Rows[i]["discriminant"]);
                if (Math.Abs((decimal)sdata.total_analysis.Rows[i]["correlation"]) != (decimal)sdata.total_analysis.Rows[i]["correlation"] ||
                    Math.Abs((decimal)sdata.total_analysis.Rows[i]["discriminant"]) != (decimal)sdata.total_analysis.Rows[i]["discriminant"])
                    for (int j = 1; j < 11; j++)
                        topic_table.Cell(i + 2, j).Range.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray10;
            }
            //oWord.Selection.MoveEnd();
            //oWord.Selection.MoveDown(WdLine, _sdata.total_analysis.Rows.Count + 1, oMissing);
            //oWord.Selection.InsertParagraphAfter();

            //topic_table.Range.set_Style(ref TableContent2);
            topic_table.Select();
            oWord.Selection.set_Style(ref TableContent2);

            topic_Rng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            topic_Rng.InsertParagraphAfter();
        }
        public void insertTotalDifficultyChart(string title, WordData sdata)
        {
            double[][] data = new double[sdata.total_analysis.Rows.Count][];
            for (int i = 0; i < sdata.total_analysis.Rows.Count; i++)
            {
                data[i] = new double[2];
                data[i][0] = Convert.ToDouble(sdata.total_analysis.Rows[i]["difficulty"]);
                //data[i][1] = Convert.ToDouble(_sdata.total_analysis.Rows[i]["discriminant"]);
                data[i][1] = Convert.ToDouble(sdata.total_analysis.Rows[i]["correlation"]);
            }

            ZedGraph.createGradient(data);
            Word.Range dist_rng = oDoc.Bookmarks.get_Item(oEndOfDoc).Range;
            dist_rng.Paste();
            Utils.mutex_clipboard.ReleaseMutex();
            dist_rng.InsertCaption(oWord.CaptionLabels["图"], title, oMissing, Word.WdCaptionPosition.wdCaptionPositionAbove, oMissing);
            dist_rng.MoveEnd(Word.WdUnits.wdParagraph, 1);
            dist_rng.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            dist_rng = oDoc.Bookmarks.get_Item(oEndOfDoc).Range;
            dist_rng.MoveStart(Word.WdUnits.wdParagraph, 1);
            dist_rng.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            dist_rng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            dist_rng.InsertParagraphAfter();
        }
        public void insertTotalChart(string title, WordData sdata)
        {
            DataTable dt = sdata.totalmark_dist;
            double[][] data = new double[dt.Rows.Count][];
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                data[i] = new double[2];
                data[i][0] = Convert.ToDouble((decimal)dt.Rows[i][0]);
                data[i][1] = Convert.ToDouble((int)dt.Rows[i][1]);

            }
            double[] cuvedata = new double[2];
            cuvedata[0] = Convert.ToDouble(sdata.avg);
            cuvedata[1] = Convert.ToDouble(sdata.stDev);
            ZedGraph.createCuveAndBar(_config, cuvedata, data, Convert.ToDouble(sdata.max));
            Word.Range dist_rng = oDoc.Bookmarks.get_Item(oEndOfDoc).Range;
            dist_rng.Paste();
            Utils.mutex_clipboard.ReleaseMutex();
            dist_rng.InsertCaption(oWord.CaptionLabels["图"], title, oMissing, Word.WdCaptionPosition.wdCaptionPositionAbove, oMissing);
            dist_rng.MoveEnd(Word.WdUnits.wdParagraph, 1);
            dist_rng.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            dist_rng = oDoc.Bookmarks.get_Item(oEndOfDoc).Range;
            dist_rng.MoveStart(Word.WdUnits.wdParagraph, 1);
            dist_rng.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            dist_rng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            dist_rng.InsertParagraphAfter();

        }
        public void insertTotalTable(string name, WordData sdata)
        {
            Word.Table Total_Table;
            Word.Range total_rng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;

            Total_Table = oDoc.Tables.Add(total_rng, 4, 7, ref oMissing, oTrue);
            object Total_title = name;
            Total_Table.Range.InsertCaption(oWord.CaptionLabels["表"], Total_title, oMissing, Word.WdCaptionPosition.wdCaptionPositionAbove, oMissing);
            total_rng.MoveEnd(Word.WdUnits.wdParagraph, 1);
            total_rng.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            //Total_Table.Range.set_Style(ref TableContent);

            //oWord.Selection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            //oWord.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            //Total_Table.Range.ParagraphFormat.Space1();
            Total_Table.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            Total_Table.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;

            //Total_Table.Range.Paragraphs.OutlineLevel = Word.WdOutlineLevel.wdOutlineLevelBodyText;
            //Total_Table.Range.Font.Size = 10;
            //Total_Table.Range.Font.Name = "黑体";
            Total_Table.Cell(1, 1).Range.Text = "总人数";
            Total_Table.Cell(1, 2).Range.Text = "满分值";
            Total_Table.Cell(1, 3).Range.Text = "最大值";
            Total_Table.Cell(1, 4).Range.Text = "最小值";
            Total_Table.Cell(1, 5).Range.Text = "平均值";
            Total_Table.Cell(1, 6).Range.Text = "标准差";
            Total_Table.Cell(1, 7).Range.Text = "差异系数";

            Total_Table.Cell(2, 1).Range.Text = sdata.total_num.ToString();
            Total_Table.Cell(2, 2).Range.Text = Convert.ToInt32(sdata.fullmark).ToString();
            Total_Table.Cell(2, 3).Range.Text = string.Format("{0:F1}", sdata.max);
            Total_Table.Cell(2, 4).Range.Text = string.Format("{0:F1}", sdata.min);
            Total_Table.Cell(2, 5).Range.Text = string.Format("{0:F2}", sdata.avg);
            Total_Table.Cell(2, 6).Range.Text = string.Format("{0:F2}", sdata.stDev);
            Total_Table.Cell(2, 7).Range.Text = string.Format("{0:F2}", sdata.Dfactor);

            Total_Table.Cell(3, 1).Range.Text = "难度";
            Total_Table.Cell(3, 2).Range.Text = "alpha系数";
            Total_Table.Cell(3, 3).Range.Text = "标准误";
            Total_Table.Cell(3, 4).Range.Text = "中数";
            Total_Table.Cell(3, 5).Range.Text = "众数";
            Total_Table.Cell(3, 6).Range.Text = "偏度";
            Total_Table.Cell(3, 7).Range.Text = "峰度";

            Total_Table.Cell(4, 1).Range.Text = string.Format("{0:F2}", sdata.difficulty);
            Total_Table.Cell(4, 2).Range.Text = string.Format("{0:F2}", sdata.alfa);
            Total_Table.Cell(4, 3).Range.Text = string.Format("{0:F2}", sdata.standardErr);
            Total_Table.Cell(4, 4).Range.Text = string.Format("{0:F2}", sdata.mean);
            Total_Table.Cell(4, 5).Range.Text = string.Format("{0:F1}", sdata.mode);
            Total_Table.Cell(4, 6).Range.Text = string.Format("{0:F2}", sdata.skewness);
            Total_Table.Cell(4, 7).Range.Text = string.Format("{0:F2}", sdata.kertosis);

            Total_Table.Range.Select();
            oWord.Selection.set_Style(ref TableContent2);

            total_rng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            total_rng.InsertParagraphAfter();
            //object WdLine = Microsoft.Office.Interop.Word.WdUnits.wdLine;
            //oWord.Selection.MoveDown(WdLine, 4, oMissing);
            //oWord.Selection.TypeParagraph();
            //oDoc.Characters.Last.InsertBreak(oParagrahbreak);
        }
        public void insertText(object type, string content)
        {
            Word.Range first = oDoc.Paragraphs.Add(ref oMissing).Range;
            first.set_Style(type);
            first.InsertBefore(content + "\n");

            oDoc.Characters.Last.Select();
            oWord.Selection.HomeKey(Word.WdUnits.wdLine, oMissing);
            oWord.Selection.Delete(Word.WdUnits.wdCharacter, oMissing);
            oWord.Selection.Range.set_Style(ExamBodyText);
        }
        public void insertTotalTable(string title)
        {
            Word.Table table;
            Word.Range range = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;

            table = oDoc.Tables.Add(range, 18, 11, ref oMissing, oTrue);
            table.Range.InsertCaption(oWord.CaptionLabels["表"], title, oMissing, Word.WdCaptionPosition.wdCaptionPositionAbove, oMissing);
            range.MoveEnd(Word.WdUnits.wdParagraph, 1);
            range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            table.Rows[1].HeadingFormat = -1;


            table.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            table.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;

            table.Cell(1, 1).Range.Text = "科目";
            table.Cell(1, 2).Range.Text = "年份";
            table.Cell(1, 3).Range.Text = "人数";
            table.Cell(1, 4).Range.Text = "满分值";
            table.Cell(1, 5).Range.Text = "最大值";
            table.Cell(1, 6).Range.Text = "最小值";
            table.Cell(1, 7).Range.Text = "平均值";
            table.Cell(1, 8).Range.Text = "标准差";
            table.Cell(1, 9).Range.Text = "差异系数";
            table.Cell(1, 10).Range.Text = "得分率";
            table.Cell(1, 11).Range.Text = "难度变化";

            for (int i = 0; i < summary.Rows.Count; i++)
            {
                table.Cell(i + 2, 1).Range.Text = subname[((string)summary.Rows[i]["sub"]).Trim()];
                table.Cell(i + 2, 2).Range.Text = ((string)summary.Rows[i]["year"]).Trim();
                table.Cell(i + 2, 3).Range.Text = ((int)summary.Rows[i]["total_num"]).ToString();
                table.Cell(i + 2, 4).Range.Text = string.Format("{0:F0}", ((decimal)summary.Rows[i]["fullmark"]));
                table.Cell(i + 2, 5).Range.Text = string.Format("{0:F1}", ((decimal)summary.Rows[i]["max"]));
                table.Cell(i + 2, 6).Range.Text = string.Format("{0:F1}", ((decimal)summary.Rows[i]["min"]));
                table.Cell(i + 2, 7).Range.Text = string.Format("{0:F1}", ((decimal)summary.Rows[i]["avg"]));
                table.Cell(i + 2, 8).Range.Text = string.Format("{0:F2}", ((decimal)summary.Rows[i]["stDev"]));
                table.Cell(i + 2, 9).Range.Text = string.Format("{0:F2}", ((decimal)summary.Rows[i]["Dfactor"]));
                table.Cell(i + 2, 10).Range.Text = string.Format("{0:F2}", ((decimal)summary.Rows[i]["difficulty"]));
                table.Cell(i + 2, 11).Range.Text = string.Format("{0:F2}", ((decimal)summary.Rows[i]["diff"]));
            }

            verticalCellMerge(table, 2, 1);
            verticalCellMerge(table, 2, 11);

            table.Select();
            oWord.Selection.set_Style(ref TableContent2);

            range = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            range.InsertParagraphAfter();

        }

        private void verticalCellMerge(Word.Table table, int startRowIndex, int columnIndex)
        {
            string previousText = table.Cell(startRowIndex++, columnIndex).Range.Text;    // 保存对比文字
            int previousRowIndex = startRowIndex - 1;    // 因刚已经+1了，所以再减回去
            for (int i = startRowIndex; i <= table.Rows.Count; ++i) // 遍历所有行的columnIndex列，发现相同的合并，从起始行的下一行开始对比
            {
                string currentText = table.Cell(i, columnIndex).Range.Text;
                if (previousText.Equals(currentText))
                {
                    table.Cell(previousRowIndex, columnIndex).Merge(table.Cell(i, columnIndex)); // 合并先前单元格和当前单元格
                    //table.Cell(previousRowIndex, columnIndex).Select();
                    //oWord.Selection.Text = currentText.TrimEnd('\r');
                    string text = currentText.Trim('\a').Trim('\r');
                    table.Cell(previousRowIndex, columnIndex).Range.Text = text;   // 因为合并后并没有将单元格内容去除，需要手动修改

                    table.Cell(previousRowIndex, columnIndex).Select();
                    oWord.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;    // 水平居中显示
                    table.Cell(previousRowIndex, columnIndex).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter; // 垂直居中
                }
                else
                {
                    previousText = currentText; // 将对比文字替换为当前的内容
                    previousRowIndex = i;   // 检索到不同的内容，将当前行下标置为先前行下标，用于合并
                }
            }
        }
        public void insertChart(string title, DataTable dt, string x_axis, string y_axis, object type, decimal fullmark)
        {
            if (dt.Columns.Count > 2)
            {
                ZedGraph.createMultipleCuve(dt, x_axis, y_axis, Convert.ToDouble(dt.Compute("Min([" + dt.Columns[0].ColumnName + "])", "")), Convert.ToDouble(dt.Compute("Max([" + dt.Columns[0].ColumnName + "])", "")), fullmark);
            }
            else if (y_axis.Equals("人数百分比") || y_axis.Equals("比率"))
            {
                ZedGraph.createMultipleCuve(dt, "分数", y_axis, Convert.ToDouble(dt.Compute("Min([" + dt.Columns[0].ColumnName + "])", "")), Convert.ToDouble(dt.Compute("Max([" + dt.Columns[0].ColumnName + "])", "")), fullmark);
            }

            else
            {
                double[][] data = new double[dt.Rows.Count][];
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    data[i] = new double[2];
                    data[i][0] = Convert.ToDouble((decimal)dt.Rows[i][0]);
                    data[i][1] = Convert.ToDouble((decimal)dt.Rows[i][1]);

                }

                ZedGraph.createDiffCuve(_config, data, Convert.ToDouble(dt.Compute("Min([" + dt.Columns[0].ColumnName + "])", "")), Convert.ToDouble(dt.Compute("Max([" + dt.Columns[0].ColumnName + "])", "")));
            }
            Word.Range dist_rng = oDoc.Bookmarks.get_Item(oEndOfDoc).Range;
            dist_rng.Paste();
            Utils.mutex_clipboard.ReleaseMutex();
            //Excel.Application eapp = new Excel.Application();
            //eapp.Visible = false;
            //Excel.Workbooks wk = eapp.Workbooks;
            //Excel._Workbook _wk = wk.Add(oMissing);
            //Excel.Sheets shs = _wk.Sheets;

            //Word.InlineShape dist_shape;
            //Word.Range dist_rng = oDoc.Bookmarks.get_Item(oEndOfDoc).Range;

            ////Excel.Workbook dist_book = (Excel.Workbook)dist_shape.OLEFormat.Object;
            //Excel.Worksheet dist_Sheet = shs.get_Item(1);

            //dist_Sheet.Cells.Clear();


            //object[,] data = new object[dt.Rows.Count + 2, dt.Columns.Count + 1];
            //for (int i = 0; i < dt.Columns.Count; i++)
            //    data[0, i] = dt.Columns[i].ColumnName;
            //int in_row = 1;
            //foreach (DataRow dr1 in dt.Rows)
            //{
            //    int col = 0;
            //    foreach (var item in dr1.ItemArray)
            //    {
            //        data[in_row, col] = item;
            //        col++;
            //    }
            //    in_row++;
            //}
            //Excel.Range rng = dist_Sheet.Range[dist_Sheet.Cells[1][1], dist_Sheet.Cells[dt.Columns.Count][dt.Rows.Count + 1]];
            //rng.Value2 = data;
            //Excel.Chart chart_dist = _wk.Charts.Add(oMissing, dist_Sheet, oMissing, oMissing);

            //Excel.Range dist_chart_rng = (Excel.Range)dist_Sheet.Cells[1, 1];

            //chart_dist.ChartWizard(dist_chart_rng.CurrentRegion, type, Type.Missing, Excel.XlRowCol.xlColumns, 1, 1, false, "", x_axis, y_axis, "");
            //if (dt.Columns.Count > 2)
            //    _wk.ActiveChart.HasLegend = true;
            //FormatExcel(_wk, x_axis, y_axis);
            //dist_shape = dist_rng.InlineShapes.AddOLEObject(ref oClassType, _wk.Name,
            //ref oMissing, ref oMissing, ref oMissing,
            //ref oMissing, ref oMissing, ref oMissing);
            //dist_shape.Range.InsertCaption(oWord.CaptionLabels["图"], title, oMissing, Word.WdCaptionPosition.wdCaptionPositionAbove, oMissing);
            //dist_rng.MoveEnd(Word.WdUnits.wdParagraph, 1);
            //dist_rng.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            ////dist_Sheet.UsedRange.CopyPicture();
            //dist_shape.Width = 375;
            //dist_shape.Height = 220;
            ////dist_rng.PasteExcelTable(true, true, false);
            ////dist_shape.ConvertToShape();
            dist_rng = oDoc.Bookmarks.get_Item(oEndOfDoc).Range;
            dist_rng.MoveStart(Word.WdUnits.wdParagraph, 1);
            dist_rng.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            dist_rng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            dist_rng.InsertParagraphAfter();
            //ReleaseExcel(_wk, eapp);
        }
    }
}
