using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExamReport
{
    public class SchoolWordCreator
    {
        private object ExamTitle0 = "ExamTitle0";
        private object CaptionTitle = "CaptionTitle";
        private object ExamTitle1 = "ExamTitle1";
        private object ExamTitle2 = "ExamTitle2";
        private object ExamTitle3 = "ExamTitle3";
        private object ExamBodyText = "ExamBodyText";
        private object TableContent = "TableContent";
        private object TableContent2 = "TableContent2";

        object oMissing = System.Reflection.Missing.Value;
        object oEndOfDoc = "\\endofdoc"; /* \endofdoc is a predefined bookmark */
        Word._Application oWord;
        Word._Document oDoc;
        WordData _sdata;
        List<PartitionData> _pdata;
        object oParagrahbreak = Microsoft.Office.Interop.Word.WdBreakType.wdLineBreak;
        object oPagebreak = Microsoft.Office.Interop.Word.WdBreakType.wdPageBreak;
        Object oTrue = true;
        Object oFalse = false;
        string _schoolname;
        object oClassType = "Excel.Chart.8";
        string _addr;

        public SchoolWordCreator(WordData sdata, List<PartitionData> pdata, string schoolname)
        {
            _sdata = sdata;
            _pdata = pdata;
            _schoolname = schoolname;
        }

        public void creating_word()
        {
            object filepath = @Utils.CurrentDirectory + @"\template.dotx";
            //Start Word and create a new document.

            oWord = new Word.Application();

            oWord.Visible = Utils.isVisible;
            oDoc = oWord.Documents.Add(ref filepath, ref oMissing,
            ref oMissing, ref oMissing);
            Utils.school_name = _schoolname;
            Utils.WriteFrontPage(oDoc);

            insertText(ExamTitle1, "总体分析");
            insertTotalTable("    总分分析表", _pdata);

            Partition_wordcreator.ChartCombine chartdata = new Partition_wordcreator.ChartCombine();
            foreach (PartitionData temp in _pdata)
            {
                chartdata.Add(temp.total_dist, temp.name);
            }


        }

        public void insertTotalTable(string title, List<PartitionData> totaldata)
        {
            Word.Table table;
            Word.Range range = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            
            table = oDoc.Tables.Add(range, totaldata.Count + 1, 10, ref oMissing, oTrue);
            table.Range.InsertCaption(oWord.CaptionLabels["表"], title, oMissing, Word.WdCaptionPosition.wdCaptionPositionAbove, oMissing);
            range.MoveEnd(Word.WdUnits.wdParagraph, 1);
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
            table.Cell(1, 10).Range.Text = "鉴别指数";

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
                table.Cell(i + 2, 10).Range.Text = string.Format("{0:F2}", ((WSLG_partitiondata)totaldata[i]).discriminant);
            }

            table.Select();
            oWord.Selection.set_Style(ref TableContent2);

            range = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            range.InsertParagraphAfter();

        }
        string FullmarkFormat(decimal remark)
        {
            return Math.Ceiling(Convert.ToDouble(remark)) == Convert.ToDouble(remark) ? Convert.ToInt32(remark).ToString() : string.Format("{0:F1}", remark);
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
    }
}
