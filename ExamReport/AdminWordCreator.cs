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
            insertTotalTable_final("    试卷总分分析表", w_data, l_data);
            insertChart("文科总分分数分布图", w_data.total_dist, "总分", "人数", 750);
            insertFreqTable_single("文科总分频数分布表", w_data.total_freq);
            insertChart("理科总分分数分布图", l_data.total_dist, "总分", "人数", 750);
            insertFreqTable_single("理科总分频数分布表", l_data.total_freq);
            insertGKLineTable("", w_data.total_level, l_data.total_level);
            insertText(ExamTitle1, "文科");
            insertSubDiffTable("文科各学科得分率表", w_data.sub_diff);
            



        }
        public void insertHistGraph(string title, DataTable data)
        {

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
                table.Cell(2, col).Range.Text = string.Format("{0:F1}", data.Rows[col - 2]["diff"]);
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
            table.Cell(1, 2).Merge(table.Cell(1, 4));
            table.Cell(1, 2).Range.Text = "文科";
            table.Cell(1, 5).Merge(table.Cell(1, 6));
            table.Cell(1, 5).Merge(table.Cell(1, 7));
            table.Cell(1, 5).Range.Text = "理科";


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
        void insertTotalTable_final(string title, Admin_WordData w_data, Admin_WordData l_data)
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

            table.Cell(2, 1).Range.Text = "文科";
            table.Cell(2, 2).Range.Text = w_data.total.totalnum.ToString();
            table.Cell(2, 3).Range.Text = string.Format("{0:F1}", w_data.total.fullmark);
            table.Cell(2, 4).Range.Text = string.Format("{0:F1}", w_data.total.max);
            table.Cell(2, 5).Range.Text = string.Format("{0:F1}", w_data.total.min);
            table.Cell(2, 6).Range.Text = string.Format("{0:F1}", w_data.total.avg);
            table.Cell(2, 7).Range.Text = string.Format("{0:F2}", w_data.total.stDev);
            table.Cell(2, 8).Range.Text = string.Format("{0:F2}", w_data.total.Dfactor);
            table.Cell(2, 9).Range.Text = string.Format("{0:F2}", w_data.total.difficulty);

            table.Cell(3, 1).Range.Text = "理科";
            table.Cell(3, 2).Range.Text = l_data.total.totalnum.ToString();
            table.Cell(3, 3).Range.Text = string.Format("{0:F1}", w_data.total.fullmark);
            table.Cell(3, 4).Range.Text = string.Format("{0:F1}", w_data.total.max);
            table.Cell(3, 5).Range.Text = string.Format("{0:F1}", w_data.total.min);
            table.Cell(3, 6).Range.Text = string.Format("{0:F1}", w_data.total.avg);
            table.Cell(3, 7).Range.Text = string.Format("{0:F2}", w_data.total.stDev);
            table.Cell(3, 8).Range.Text = string.Format("{0:F2}", w_data.total.Dfactor);
            table.Cell(3, 9).Range.Text = string.Format("{0:F2}", w_data.total.difficulty);

            table.Select();
            oWord.Selection.set_Style(ref TableContent2);

            range = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            range.InsertParagraphAfter();
        }

        public void insertChart(string title, DataTable dt, string x_axis, string y_axis, double fullmark)
        {
            double[][] data = new double[dt.Rows.Count][];
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                data[i] = new double[2];
                data[i][0] = Convert.ToDouble((decimal)dt.Rows[i][0]);
                data[i][1] = Convert.ToDouble((decimal)dt.Rows[i][1]);

            }
            ZedGraph.createCuve(_config, x_axis, y_axis, data, 0, fullmark, Convert.ToDouble(dt.Compute("Max([" + dt.Columns[1].ColumnName + "])", "")));
            
            
            Word.Range dist_rng = oDoc.Bookmarks.get_Item(oEndOfDoc).Range;
            dist_rng.Paste();
            Utils.mutex_clipboard.ReleaseMutex();
           
            dist_rng = oDoc.Bookmarks.get_Item(oEndOfDoc).Range;
            dist_rng.MoveStart(Word.WdUnits.wdParagraph, 1);
            dist_rng.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            dist_rng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            dist_rng.InsertParagraphAfter();
        }

    }
}
