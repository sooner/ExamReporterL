using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
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
        public Partition_wordcreator.ChartCombine year1_comb;
        public Partition_wordcreator.ChartCombine year2_comb;
        public DataTable summary;

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
            int count = 10;

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

            for (int i = 0; i < totaldata.Count; i++)
            {
                table.Cell(i + 2, 1).Range.Text = ((PartitionData)totaldata[i]).name;
                table.Cell(i + 2, 2).Range.Text = ((PartitionData)totaldata[i]).total_num.ToString();
                table.Cell(i + 2, 3).Range.Text = FullmarkFormat((decimal)((PartitionData)totaldata[i]).fullmark);
                table.Cell(i + 2, 4).Range.Text = string.Format("{0:F1}", ((PartitionData)totaldata[i]).max);
                table.Cell(i + 2, 5).Range.Text = string.Format("{0:F1}", ((PartitionData)totaldata[i]).min);
                table.Cell(i + 2, 6).Range.Text = string.Format("{0:F2}", ((PartitionData)totaldata[i]).avg);
                table.Cell(i + 2, 7).Range.Text = string.Format("{0:F2}", ((PartitionData)totaldata[i]).stDev);
                table.Cell(i + 2, 8).Range.Text = string.Format("{0:F2}", ((PartitionData)totaldata[i]).Dfactor);
                table.Cell(i + 2, 9).Range.Text = string.Format("{0:F2}", ((PartitionData)totaldata[i]).difficulty);
                table.Cell(i + 2, 10).Range.Text = string.Format("{0:F2}", ((PartitionData)totaldata[i]).discriminant);
            }

            table.Select();
            oWord.Selection.set_Style(ref TableContent2);

            range = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            range.InsertParagraphAfter();

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
