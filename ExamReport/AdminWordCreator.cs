using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExamReport
{
    class AdminWordCreator
    {
        private object CaptionTitle = "CaptionTitle";
        private object ExamTitle0 = "ExamTitle0";
        private object ExamTitle1 = "ExamTitle1";
        private object ExamTitle2 = "ExamTitle2";
        private object ExamTitle3 = "ExamTitle3";
        private object ExamBodyText = "ExamBodyText";
        private object TableContent = "TableContent";
        private object TableContent2 = "TableContent2";

        private Configuration _config;

        object oMissing = System.Reflection.Missing.Value;
        object oEndOfDoc = "\\endofdoc"; /* \endofdoc is a predefined bookmark */
        Word._Application oWord;
        Word._Document oDoc;
        object oParagrahbreak = Microsoft.Office.Interop.Word.WdBreakType.wdLineBreak;
        Object oTrue = true;
        Object oFalse = false;

        object oClassType = "Excel.Chart.8";

        public AdminWordCreator(Configuration config)
        {
            _config = config;
        }
        public void creating_word(Admin_WordData w_data, Admin_WordData l_data)
        {
            object filepath = @_config.CurrentDirectory + @"\template2.dotx";
            //object filepath = @"D:\项目\给王卅的编程资料\中考\c.dotx";
            //Start Word and create a new document.

            oWord = new Word.Application();

            oWord.Visible = _config.isVisible;
            oDoc = oWord.Documents.Add(ref filepath, ref oMissing,
            ref oMissing, ref oMissing);
            Utils.WriteFrontPage(_config, oDoc);

            insertText(ExamTitle0, "北京市整体");
            insertText(ExamTitle1, "总体");
            insertTotalTable_final("    试卷总分分析表", "文科", w_data.total, "理科", l_data.total);
            insertChart("    文科总分分数分布图", w_data);
            insertFreqTable_single("    文科总分频数分布表", w_data.total_freq);
            insertChart("    理科总分分数分布图", l_data);
            insertFreqTable_single("    理科总分频数分布表", l_data.total_freq);
            insertGKLineTable("", w_data.total_level, l_data.total_level);
            insertText(ExamTitle1, "文科");
            insertSubDiffTable("    文科各学科得分率表", w_data.sub_diff);
            insertHistGraph("    文科各学科得分率图", w_data.sub_diff);
            insertSubDiffTable("    理科各学科得分率表", l_data.sub_diff);
            insertHistGraph("    理科各学科得分率图", l_data.sub_diff);

            insertText(ExamTitle0, "城区、郊区");
            insertText(ExamTitle1, "文科");

            insertTotalTable_final("    城区、郊区总分分析表", "城区", w_data.urban, "郊区", w_data.country);
            insertUrbCntTable("    城区、郊区文科各学科得分率分析表", w_data.urban_sub, w_data.country_sub);
            insertMultiHistGraph("    城区、郊区文科各学科得分率分析图", w_data.urban_sub, w_data.country_sub);

            insertText(ExamTitle1, "理科");

            insertTotalTable_final("    城区、郊区总分分析表", "城区", l_data.urban, "郊区", l_data.country);
            insertUrbCntTable("    城区、郊区文科各学科得分率分析表", l_data.urban_sub, l_data.country_sub);
            insertMultiHistGraph("    城区、郊区文科各学科得分率分析图", l_data.urban_sub, l_data.country_sub);

            insertText(ExamTitle0, "区县分析");
            insertText(ExamTitle1, "文科");
            insertText(ExamTitle2, "总分");

            insertQXtable("    各区文科总分分析表", w_data.districts.Rows[0], w_data.districts.Rows[1]);
            insertText(ExamTitle2, "语文学科"); 
            insertQXtable("    语文学科得分率表", w_data.districts.Rows[2], w_data.districts.Rows[3]);
            insertQXchart("    语文学科得分率图", w_data.districts.Rows[3]);
            insertText(ExamTitle2, "数学（文）学科"); 
            insertQXtable("    数学（文）学科得分率表", w_data.districts.Rows[4], w_data.districts.Rows[5]);
            insertQXchart("    数学（文）学科得分率图", w_data.districts.Rows[5]);
            insertText(ExamTitle2, "英语"); 
            insertQXtable("    英语学科得分率表", w_data.districts.Rows[6], w_data.districts.Rows[7]);
            insertQXchart("    英语学科得分率图", w_data.districts.Rows[7]);
            insertText(ExamTitle2, "历史"); 
            insertQXtable("    历史学科得分率表", w_data.districts.Rows[8], w_data.districts.Rows[9]);
            insertQXchart("    历史学科得分率图", w_data.districts.Rows[9]);
            insertText(ExamTitle2, "地理"); 
            insertQXtable("    地理学科得分率表", w_data.districts.Rows[10], w_data.districts.Rows[11]);
            insertQXchart("    地理学科得分率图", w_data.districts.Rows[11]);
            insertText(ExamTitle2, "政治"); 
            insertQXtable("    政治学科得分率表", w_data.districts.Rows[12], w_data.districts.Rows[13]);
            insertQXchart("    政治学科得分率图", w_data.districts.Rows[13]);

            insertText(ExamTitle1, "理科");
            insertText(ExamTitle2, "总分"); 
            insertQXtable("    各区理科总分分析表", l_data.districts.Rows[0], l_data.districts.Rows[1]);
            insertText(ExamTitle2, "语文"); 
            insertQXtable("    语文学科得分率表", l_data.districts.Rows[2], l_data.districts.Rows[3]);
            insertQXchart("    语文学科得分率图", w_data.districts.Rows[3]);
            insertText(ExamTitle2, "数学（理）"); 
            insertQXtable("    数学（理）学科得分率表", l_data.districts.Rows[4], l_data.districts.Rows[5]);
            insertQXchart("    数学（理）学科得分率图", w_data.districts.Rows[5]);
            insertText(ExamTitle2, "英语"); 
            insertQXtable("    英语学科得分率表", l_data.districts.Rows[6], l_data.districts.Rows[7]);
            insertQXchart("    英语学科得分率图", w_data.districts.Rows[7]);
            insertText(ExamTitle2, "物理"); 
            insertQXtable("    物理学科得分率表", l_data.districts.Rows[8], l_data.districts.Rows[9]);
            insertQXchart("    物理学科得分率图", w_data.districts.Rows[9]);
            insertText(ExamTitle2, "化学"); 
            insertQXtable("    化学学科得分率表", l_data.districts.Rows[10], l_data.districts.Rows[11]);
            insertQXchart("    化学学科得分率图", w_data.districts.Rows[11]);
            insertText(ExamTitle2, "生物"); 
            insertQXtable("    生物学科得分率表", l_data.districts.Rows[12], l_data.districts.Rows[13]);
            insertQXchart("    生物学科得分率图", w_data.districts.Rows[13]);

            foreach (Word.TableOfContents table in oDoc.TablesOfContents)
                table.Update();
            Utils.Save(_config, oDoc, oWord);
        }
        public void insertQXchart(string title, DataRow diff)
        {
            ZedGraph.createQXBar(diff);

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

        public void insertQXtable(string title, DataRow avg, DataRow diff)
        {
            Word.Table table;
            Word.Range range = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            table = oDoc.Tables.Add(range, 6, 11, ref oMissing, oTrue);

            table.Range.InsertCaption(oWord.CaptionLabels["表"], title, oMissing, Word.WdCaptionPosition.wdCaptionPositionAbove, oMissing);
            range.MoveEnd(Word.WdUnits.wdParagraph, 1);
            range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

            table.Cell(1, 1).Range.Text = "区县";
            table.Cell(1, 2).Range.Text = "全市";
            table.Cell(1, 3).Range.Text = "城区";
            table.Cell(1, 4).Range.Text = "郊区";
            table.Cell(1, 5).Range.Text = "东城";
            table.Cell(1, 6).Range.Text = "西城";
            table.Cell(1, 7).Range.Text = "朝阳";
            table.Cell(1, 8).Range.Text = "丰台";
            table.Cell(1, 9).Range.Text = "石景山";
            table.Cell(1, 10).Range.Text = "海淀";
            table.Cell(1, 11).Range.Text = "门头沟";

            table.Cell(4, 1).Range.Text = "区县";
            table.Cell(4, 2).Range.Text = "燕山";
            table.Cell(4, 3).Range.Text = "房山";
            table.Cell(4, 4).Range.Text = "通州";
            table.Cell(4, 5).Range.Text = "顺义";
            table.Cell(4, 6).Range.Text = "昌平";
            table.Cell(4, 7).Range.Text = "大兴";
            table.Cell(4, 8).Range.Text = "怀柔";
            table.Cell(4, 9).Range.Text = "平谷";
            table.Cell(4, 10).Range.Text = "密云";
            table.Cell(4, 11).Range.Text = "延庆";

            table.Cell(2, 1).Range.Text = "平均分";
            table.Cell(3, 1).Range.Text = "得分率";

            table.Cell(5, 1).Range.Text = "平均分";
            table.Cell(6, 1).Range.Text = "得分率";

            for (int i = 0; i < 10; i++)
            {
                table.Cell(2, i + 2).Range.Text = string.Format("{0:F1}", avg[i]);
                table.Cell(3, i + 2).Range.Text = string.Format("{0:F2}", diff[i]);

                table.Cell(5, i + 2).Range.Text = string.Format("{0:F1}", avg[i + 10]);
                table.Cell(6, i + 2).Range.Text = string.Format("{0:F2}", diff[i + 10]);
            }
            table.Rows[1].HeadingFormat = -1;


            table.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            table.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;

            table.Select();
            oWord.Selection.set_Style(ref TableContent2);
            range = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            range.InsertParagraphAfter();

        }
        public void insertUrbCntTable(string title, DataTable urban, DataTable country)
        {
            Word.Table table;
            Word.Range range = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            table = oDoc.Tables.Add(range, 3, urban.Rows.Count + 1, ref oMissing, oTrue);

            table.Range.InsertCaption(oWord.CaptionLabels["表"], title, oMissing, Word.WdCaptionPosition.wdCaptionPositionAbove, oMissing);
            range.MoveEnd(Word.WdUnits.wdParagraph, 1);
            range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

            table.Cell(1, 1).Range.Text = "分类";
            table.Cell(2, 1).Range.Text = "城区";
            table.Cell(3, 1).Range.Text = "郊区";

            for (int col = 2; col <= urban.Rows.Count + 1; col++)
            {
                table.Cell(1, col).Range.Text = urban.Rows[col - 2]["sub"].ToString().Trim();
                table.Cell(2, col).Range.Text = string.Format("{0:F2}", urban.Rows[col - 2]["diff"]);
                table.Cell(3, col).Range.Text = string.Format("{0:F2}", country.Rows[col - 2]["diff"]);
            }

            table.Rows[1].HeadingFormat = -1;


            table.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            table.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;

            table.Select();
            oWord.Selection.set_Style(ref TableContent2);
            range = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            range.InsertParagraphAfter();
        }
        public void insertMultiHistGraph(string title, DataTable urban, DataTable country)
        {
            Dictionary<string, DataTable> data_list = new Dictionary<string, DataTable>();
            data_list.Add("城区", urban);
            data_list.Add("郊区", country);

            //ZedGraph.createSubDiffBar(data_list);
            DotNetCharting.CreateMutipleColumn(data_list);

            Word.Range dist_rng = oDoc.Bookmarks.get_Item(oEndOfDoc).Range;
            dist_rng.Paste();
            Utils.mutex_clipboard.ReleaseMutex();
            dist_rng.InsertCaption(oWord.CaptionLabels["图"], title, oMissing, Word.WdCaptionPosition.wdCaptionPositionBelow, oMissing);
            dist_rng.MoveEnd(Word.WdUnits.wdParagraph, 1);
            dist_rng.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            dist_rng = oDoc.Bookmarks.get_Item(oEndOfDoc).Range;
            dist_rng.MoveStart(Word.WdUnits.wdParagraph, 1);
            dist_rng.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            dist_rng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            dist_rng.InsertParagraphAfter();
        }
        public void insertHistGraph(string title, DataTable data)
        {
            List<DataTable> data_list = new List<DataTable>();
            data_list.Add(data);

            DotNetCharting.CreateColumn(data);
            //ZedGraph.createSubDiffBar(data_list);

            Word.Range dist_rng = oDoc.Bookmarks.get_Item(oEndOfDoc).Range;
            dist_rng.Paste();
            Utils.mutex_clipboard.ReleaseMutex();
            dist_rng.InsertCaption(oWord.CaptionLabels["图"], title, oMissing, Word.WdCaptionPosition.wdCaptionPositionBelow, oMissing);
            dist_rng.MoveEnd(Word.WdUnits.wdParagraph, 1);
            dist_rng.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            dist_rng = oDoc.Bookmarks.get_Item(oEndOfDoc).Range;
            dist_rng.MoveStart(Word.WdUnits.wdParagraph, 1);
            dist_rng.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            dist_rng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            dist_rng.InsertParagraphAfter();
        }
        public void insertSubDiffTable(string title, DataTable data)
        {
            Word.Table table;
            Word.Range range = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            table = oDoc.Tables.Add(range, 2, data.Rows.Count + 1, ref oMissing, oTrue);

            table.Range.InsertCaption(oWord.CaptionLabels["表"], title, oMissing, Word.WdCaptionPosition.wdCaptionPositionAbove, oMissing);
            range.MoveEnd(Word.WdUnits.wdParagraph, 1);
            range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

            table.Cell(1, 1).Range.Text = "学科";
            table.Cell(2, 1).Range.Text = "得分率";

            for (int col = 2; col <= data.Rows.Count + 1; col++)
            {
                table.Cell(1, col).Range.Text = data.Rows[col - 2]["sub"].ToString().Trim();
                table.Cell(2, col).Range.Text = string.Format("{0:F2}", data.Rows[col - 2]["diff"]);
            }

            table.Rows[1].HeadingFormat = -1;


            table.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            table.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;

            table.Select();
            oWord.Selection.set_Style(ref TableContent2);
            range = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            range.InsertParagraphAfter();

        }
        public void insertGKLineTable(string title, DataTable w_data, DataTable l_data)
        {
            Word.Table table;
            Word.Range range = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            table = oDoc.Tables.Add(range, 6, 7, ref oMissing, oTrue);

            table.Range.InsertCaption(oWord.CaptionLabels["表"], title, oMissing, Word.WdCaptionPosition.wdCaptionPositionAbove, oMissing);
            range.MoveEnd(Word.WdUnits.wdParagraph, 1);
            range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

            //table.Cell(1, 1).Range.Text = "类别";
            //table.Cell(1, 2).Range.Text = "文科";
            //table.Cell(1, 3).Range.Text = "文科";
            //table.Cell(1, 4).Range.Text = "文科";
            //table.Cell(1, 5).Range.Text = "理科";
            //table.Cell(1, 6).Range.Text = "理科";
            //table.Cell(1, 7).Range.Text = "理科";

            //table.Cell(2, 1).Range.Text = "类别";
            table.Cell(2, 2).Range.Text = "分数线";
            table.Cell(2, 3).Range.Text = "人数";
            table.Cell(2, 4).Range.Text = "比率";
            table.Cell(2, 5).Range.Text = "分数线";
            table.Cell(2, 6).Range.Text = "人数";
            table.Cell(2, 7).Range.Text = "比率";

            for (int i = 0; i < w_data.Rows.Count; i++)
            {
                table.Cell(i + 3, 1).Range.Text = w_data.Rows[i]["text"].ToString().Trim();
                table.Cell(i + 3, 2).Range.Text = w_data.Rows[i]["level"].ToString().Trim();
                table.Cell(i + 3, 3).Range.Text = w_data.Rows[i]["frequency"].ToString();
                table.Cell(i + 3, 4).Range.Text = string.Format("{0:F1}", w_data.Rows[i]["rate"]);

                table.Cell(i + 3, 5).Range.Text = l_data.Rows[i]["level"].ToString().Trim();
                table.Cell(i + 3, 6).Range.Text = l_data.Rows[i]["frequency"].ToString().Trim();
                table.Cell(i + 3, 7).Range.Text = string.Format("{0:F1}",l_data.Rows[i]["rate"]);
            }

            table.Rows[1].HeadingFormat = -1;


            table.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            table.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;

            table.Select();
            oWord.Selection.set_Style(ref TableContent2);

            table.Cell(1, 1).Merge(table.Cell(2, 1));
            table.Cell(1, 1).Range.Text = "类别";
            table.Cell(1, 2).Merge(table.Cell(1, 3));
            table.Cell(1, 2).Merge(table.Cell(1, 3));
            table.Cell(1, 2).Range.Text = "文科";
            table.Cell(1, 3).Merge(table.Cell(1, 4));
            table.Cell(1, 3).Merge(table.Cell(1, 4));
            table.Cell(1, 3).Range.Text = "理科";


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
        void insertFreqTable_single(string title, DataTable dt)
        {
            Word.Table table;
            Word.Range range = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            table = oDoc.Tables.Add(range, dt.Rows.Count + 1, 5, ref oMissing, oTrue);

            table.Range.InsertCaption(oWord.CaptionLabels["表"], title, oMissing, Word.WdCaptionPosition.wdCaptionPositionAbove, oMissing);
            range.MoveEnd(Word.WdUnits.wdParagraph, 1);
            range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;


            table.Cell(1, 1).Range.Text = "分值";
            table.Cell(1, 2).Range.Text = "人数";
            table.Cell(1, 3).Range.Text = "比率(%)";
            table.Cell(1, 4).Range.Text = "累计人数";
            table.Cell(1, 5).Range.Text = "累计比率(%)";


            for (int i = 0; i < dt.Rows.Count; i++)
            {
                table.Cell(i + 2, 1).Range.Text = dt.Rows[i]["totalmark"].ToString().Trim();
                table.Cell(i + 2, 2).Range.Text = dt.Rows[i]["frequency"].ToString();
                table.Cell(i + 2, 3).Range.Text = string.Format("{0:F2}", dt.Rows[i]["rate"]);
                table.Cell(i + 2, 4).Range.Text = dt.Rows[i]["accumulateFreq"].ToString();
                table.Cell(i + 2, 5).Range.Text = string.Format("{0:F2}", dt.Rows[i]["accumulateRate"]);
            }

            table.Rows[1].HeadingFormat = -1;


            table.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            table.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;

            table.Select();
            oWord.Selection.set_Style(ref TableContent2);

            range = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            range.InsertParagraphAfter();
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
        void insertTotalTable_final(string title, string w_title, Admin_WordData.basic_stat w_data, string l_title, Admin_WordData.basic_stat l_data)
        {
            Word.Table table;
            Word.Range range = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;

            table = oDoc.Tables.Add(range, 3, 9, ref oMissing, oTrue);
            table.Range.InsertCaption(oWord.CaptionLabels["表"], title, oMissing, Word.WdCaptionPosition.wdCaptionPositionAbove, oMissing);
            //range.MoveEnd(Word.WdUnits.wdParagraph, 1);
            range.Select();
            range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;


            table.Rows[1].HeadingFormat = -1;


            table.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            table.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;

            table.Cell(1, 1).Range.Text = "分类";
            table.Cell(1, 2).Range.Text = "人数";
            table.Cell(1, 3).Range.Text = "满分值";
            table.Cell(1, 4).Range.Text = "最大值";
            table.Cell(1, 5).Range.Text = "最小值";
            table.Cell(1, 6).Range.Text = "平均值";
            table.Cell(1, 7).Range.Text = "标准差";
            table.Cell(1, 8).Range.Text = "差异系数";
            table.Cell(1, 9).Range.Text = "得分率";

            table.Cell(2, 1).Range.Text = w_title;
            table.Cell(2, 2).Range.Text = w_data.totalnum.ToString();
            table.Cell(2, 3).Range.Text = string.Format("{0:F1}", w_data.fullmark);
            table.Cell(2, 4).Range.Text = string.Format("{0:F1}", w_data.max);
            table.Cell(2, 5).Range.Text = string.Format("{0:F1}", w_data.min);
            table.Cell(2, 6).Range.Text = string.Format("{0:F1}", w_data.avg);
            table.Cell(2, 7).Range.Text = string.Format("{0:F2}", w_data.stDev);
            table.Cell(2, 8).Range.Text = string.Format("{0:F2}", w_data.Dfactor);
            table.Cell(2, 9).Range.Text = string.Format("{0:F2}", w_data.difficulty);

            table.Cell(3, 1).Range.Text = l_title;
            table.Cell(3, 2).Range.Text = l_data.totalnum.ToString();
            table.Cell(3, 3).Range.Text = string.Format("{0:F1}", l_data.fullmark);
            table.Cell(3, 4).Range.Text = string.Format("{0:F1}", l_data.max);
            table.Cell(3, 5).Range.Text = string.Format("{0:F1}", l_data.min);
            table.Cell(3, 6).Range.Text = string.Format("{0:F1}", l_data.avg);
            table.Cell(3, 7).Range.Text = string.Format("{0:F2}", l_data.stDev);
            table.Cell(3, 8).Range.Text = string.Format("{0:F2}", l_data.Dfactor);
            table.Cell(3, 9).Range.Text = string.Format("{0:F2}", l_data.difficulty);

            table.Select();
            oWord.Selection.set_Style(ref TableContent2);

            range = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            range.InsertParagraphAfter();
        }

        public void insertChart(string title, Admin_WordData sdata)
        {
            DataTable dt = sdata.total_dist;
            double[][] data = new double[dt.Rows.Count][];
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                data[i] = new double[2];
                data[i][0] = Convert.ToDouble((int)dt.Rows[i][0]);
                data[i][1] = Convert.ToDouble((int)dt.Rows[i][1]);

            }
            double[] cuvedata = new double[2];
            cuvedata[0] = Convert.ToDouble(sdata.total.avg);
            cuvedata[1] = Convert.ToDouble(sdata.total.stDev);
            ZedGraph.createCuveAndBar(_config, cuvedata, data, Convert.ToDouble(sdata.total.max));
            


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

    }
}
