﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using Graph = Microsoft.Office.Interop.Graph;

namespace ExamReport
{
    class GKCompWordCreator
    {
        public bool is_wk;
        public string year;

        object oEndOfDoc = "\\endofdoc"; /* \endofdoc is a predefined bookmark */
        Word._Application oWord;
        Word._Document oDoc;
        object oMissing = System.Reflection.Missing.Value;
        Object oTrue = true;
        private object TableContent2 = "TableContent2";

        decimal[,] result = new decimal[4, 6];
        DataTable qx;

        public void pre_process(DataTable datatable, DataTable cq_table, DataTable jq_table)
        {
            
            string[] subs = {"yw", "sx", "yy", "hx_zz", "sw_ls", "wl_dl" };
            Dictionary<string, int> sub_fullmark_w = new Dictionary<string, int> { 
                {"yw",150}, {"sx", 150}, {"yy", 150}, {"hx_zz", 100}, {"sw_ls", 100}, {"wl_dl", 100}};
            Dictionary<string, int> sub_fullmark_l = new Dictionary<string, int> { 
                {"yw",150}, {"sx", 150}, {"yy", 150}, {"hx_zz", 100}, {"sw_ls", 80}, {"wl_dl", 120}};
            Dictionary<string, int> sub_fullmark;
            if (is_wk)
                sub_fullmark = sub_fullmark_w;
            else
                sub_fullmark = sub_fullmark_l;

            int i = 0;
            foreach(string key in sub_fullmark.Keys)
            {
                result[0, i] = datatable.AsEnumerable().Select(c => c.Field<decimal>(key)).Average() / sub_fullmark[key];
                result[1, i] = cq_table.AsEnumerable().Select(c => c.Field<decimal>(subs[i])).Average() / sub_fullmark[key];
                result[2, i] = jq_table.AsEnumerable().Select(c => c.Field<decimal>(subs[i])).Average() / sub_fullmark[key];
                result[3, i] = result[1, i] - result[2, i];

                i++;
            }

            qx = datatable.AsEnumerable().GroupBy(c => c.Field<string>("qxdm")).Select(c => new
            {
                qxdm = c.Key.ToString().Trim(),
                count = c.Count(),
                yw = c.Average(p => p.Field<decimal>("yw")) / sub_fullmark["yw"],
                sx = c.Average(p => p.Field<decimal>("sx")) / sub_fullmark["sx"],
                yy = c.Average(p => p.Field<decimal>("yy")) / sub_fullmark["yy"],
                hx_zz = c.Average(p => p.Field<decimal>("hx_zz")) / sub_fullmark["hx_zz"],
                sw_ls = c.Average(p => p.Field<decimal>("sw_ls")) / sub_fullmark["sw_ls"],
                wl_dl = c.Average(p => p.Field<decimal>("wl_dl")) / sub_fullmark["wl_dl"]

            }).ToDataTable();



        }

        public void creating_word()
        {
            
            object oMissing = System.Reflection.Missing.Value;
            object filepath = @Utils.CurrentDirectory + @"\compare_doc.docx";
            //Start Word and create a new document.
            oWord = new Word.Application();

            oWord.Visible = true;
            oDoc = oWord.Documents.Add(ref oMissing, ref oMissing,
            ref oMissing, ref oMissing);

            oDoc.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;

            insertCJTable();
        }

        public void insertCJTable()
        {
            //int count = ((PartitionData)totaldata[totaldata.Count - 1]).groups_analysis.Rows.Count;
            Word.Table table;
            Word.Range range = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            int col = 11;
            table = oDoc.Tables.Add(range, 7 + qx.Rows.Count, 13, ref oMissing, oTrue);
            //table.Range.InsertCaption(oWord.CaptionLabels["表"], title, oMissing, Word.WdCaptionPosition.wdCaptionPositionAbove, oMissing);
            range.MoveEnd(Word.WdUnits.wdParagraph, 1);
            range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            table.Rows[1].HeadingFormat = -1;


            table.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            table.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;

            //table.Cell(1, 1).Range.Text = "题组";

            for(int i = 2; i < 14; i++)
                table.Cell(1, i).Range.Text = year;

            table.Cell(2, 2).Range.Text = "语文";
            table.Cell(2, 3).Range.Text = "语文";
            table.Cell(2, 4).Range.Text = "数学";
            table.Cell(2, 5).Range.Text = "数学";
            table.Cell(2, 6).Range.Text = "英语";
            table.Cell(2, 7).Range.Text = "英语";
            if (is_wk)
            {
                table.Cell(2, 8).Range.Text = "政治";
                table.Cell(2, 9).Range.Text = "政治";
                table.Cell(2, 10).Range.Text = "历史";
                table.Cell(2, 11).Range.Text = "历史";
                table.Cell(2, 12).Range.Text = "地理";
                table.Cell(2, 13).Range.Text = "地理";
            }
            else
            {
                table.Cell(2, 8).Range.Text = "化学";
                table.Cell(2, 9).Range.Text = "化学";
                table.Cell(2, 10).Range.Text = "生物";
                table.Cell(2, 11).Range.Text = "生物";
                table.Cell(2, 12).Range.Text = "物理";
                table.Cell(2, 13).Range.Text = "物理";
            }

            table.Cell(3, 1).Range.Text = "全市";
            table.Cell(4, 1).Range.Text = "城区";
            table.Cell(5, 1).Range.Text = "郊区";
            table.Cell(6, 1).Range.Text = "城郊差异";

            for (int i = 0; i < 4; i++)
            {
                for (int j = 0; j < 6; j++)
                {
                    table.Cell(i + 3, 2 * j + 2).Range.Text = string.Format("{0:F2}", result[i, j]);
                    table.Cell(i + 3, 2 * j + 3).Range.Text = string.Format("{0:F2}", result[i, j]);
                }
            }

            for (int i = 0; i < 6; i++)
            {
                table.Cell(7, 2 + i * 2).Range.Text = "得分率";
                table.Cell(7, 3 + i * 2).Range.Text = "与全市对比";
            }

            for (int i = 0; i < qx.Rows.Count; i++)
            {
                for (int j = 0; j < 6; j++)
                {
                    table.Cell(8 + i, 2 + j * 2).Range.Text = string.Format("{0:F2}", qx.Rows[i][j+2]);
                    table.Cell(8 + i, 3 + j * 2).Range.Text = string.Format("{0:F2}", (decimal)qx.Rows[i][j+2] - result[0, j]);
                }
            }
            
            //table.Select();
            //oWord.Selection.set_Style(ref TableContent2);

            for(int i = 2; i <= 6; i++)
                horizonCellMerge(table, i, 2);

            range = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            range.InsertParagraphAfter();
        }

        private void horizonCellMerge(Word.Table table, int RowIndex, int startcolumnIndex)
        {
            string previousText = table.Cell(RowIndex, startcolumnIndex++).Range.Text;    // 保存对比文字
            int previouscolumnIndex = startcolumnIndex - 1;    // 因刚已经+1了，所以再减回去
            for (int i = startcolumnIndex; i <= table.Columns.Count; ) // 遍历所有行的columnIndex列，发现相同的合并，从起始行的下一行开始对比
            {
                string currentText = table.Cell(RowIndex, i).Range.Text;
                if (previousText.Equals(currentText))
                {
                    table.Cell(RowIndex, previouscolumnIndex).Merge(table.Cell(RowIndex, i)); // 合并先前单元格和当前单元格
                    //table.Cell(previousRowIndex, columnIndex).Select();
                    //oWord.Selection.Text = currentText.TrimEnd('\r');
                    string text = currentText.Trim('\a').Trim('\r');
                    table.Cell(RowIndex, previouscolumnIndex).Range.Text = text;   // 因为合并后并没有将单元格内容去除，需要手动修改

                    table.Cell(RowIndex, previouscolumnIndex).Select();
                    oWord.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;    // 水平居中显示
                    table.Cell(RowIndex, previouscolumnIndex).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter; // 垂直居中
                }
                else
                {
                    previousText = currentText; // 将对比文字替换为当前的内容
                    previouscolumnIndex = i;   // 检索到不同的内容，将当前行下标置为先前行下标，用于合并
                    i++;
                }
            }
        }
    }
}
