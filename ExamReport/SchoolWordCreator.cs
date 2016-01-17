using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
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

        public Configuration _config;
        object oMissing = System.Reflection.Missing.Value;
        object oEndOfDoc = "\\endofdoc"; /* \endofdoc is a predefined bookmark */
        Word._Application oWord;
        Word._Document oDoc;
        WordData _sdata;
        List<WSLG_partitiondata> _pdata;
        object oParagrahbreak = Microsoft.Office.Interop.Word.WdBreakType.wdLineBreak;
        object oPagebreak = Microsoft.Office.Interop.Word.WdBreakType.wdPageBreak;
        Object oTrue = true;
        Object oFalse = false;
        string _schoolname;
        object oClassType = "Excel.Chart.8";

        Dictionary<string, List<string>> _groups_group;
        DataTable _groups;

        WordData zh_data;
        List<WSLG_partitiondata> zh_pdata;
        DataTable groups;
        Dictionary<string, List<string>> groups_group;


        public SchoolWordCreator(WordData sdata, List<WSLG_partitiondata> pdata, DataTable groups, string schoolname, Dictionary<string, List<string>> groups_group)
        {
            _sdata = sdata;
            _pdata = pdata;
            _schoolname = schoolname;
            _groups_group = groups_group;
            _groups = groups;
        }

        public void SetUpZHparam(WordData zh_data_, List<WSLG_partitiondata> zh_pdata_, DataTable groups_, Dictionary<string, List<string>> groups_group_)
        {
            zh_data = zh_data_;
            zh_pdata = zh_pdata_;
            groups = groups_;
            groups_group = groups_group_;
        }

        public void creating_ZH_word()
        {
            object filepath = @_config.CurrentDirectory + @"\template2.dotx";
            //Start Word and create a new document.

            oWord = new Word.Application();

            oWord.Visible = _config.isVisible;
            oDoc = oWord.Documents.Add(ref filepath, ref oMissing,
            ref oMissing, ref oMissing);
            Utils.WriteFrontPage(_config, oDoc, _schoolname);

            insertText(ExamTitle0, " 整体统计分析");
            insertText(ExamTitle1, "总体分析");
            insertTotalTable("    总分分析表", zh_pdata);
            ChartCombine chartdata = new ChartCombine();
            foreach (PartitionData temp in zh_pdata)
            {
                chartdata.Add(temp.total_dist, temp.name);
            }
            insertChart("    总分分布曲线图", chartdata.target, "分数", "比率（%）", Excel.XlChartType.xlXYScatterLines, ((PartitionData)zh_pdata[0]).fullmark);
            insertGroupTopic("    题组整体分析表", zh_pdata, 3);
            insertFreqTable("    总分频数分布表", zh_pdata);
            insertText(ExamTitle1, "题组分析");
            List<string> keys = new List<string>(groups_group.Keys);
            int group_count = 3;
            for (int i = 1; i < groups_group.Count; i++)
            {
                insertText(ExamTitle2, keys[i]);
                foreach (string group in groups_group[keys[i]])
                {
                    insertText(ExamTitle3, group);
                    insertTH(groups.Rows[group_count]["th"].ToString().Trim());
                    insertSingleGroupTotal("    " + group + "总分分析表", group_count, true, zh_pdata);
                    ChartCombine tempdata = new ChartCombine();
                    for (int j = 0; j < zh_pdata.Count; j++)
                    {
                        tempdata.Add(getGroupData(zh_pdata, j, group_count).group_dist, ((PartitionData)zh_pdata[j]).name);
                    }
                    insertChart("    " + group + "分数分布图", tempdata.target, "分数", "比率(%)", Excel.XlChartType.xlLineMarkers, (decimal)((PartitionData)zh_pdata[0]).groups_analysis.Rows[group_count]["fullmark"]);
                    insertGroupDiffChart("    " + group + "难度曲线图", ((WordData.group_data)zh_data.single_group_analysis[group_count]).group_difficulty);
                    insertGroupSingleAnalysis("    " + group + "分组分析表", ((WordData.group_data)zh_data.single_group_analysis[group_count]).group_detail);
                    oDoc.Characters.Last.InsertBreak(oPagebreak);
                    group_count++;
                }
            }
            
            insertText(ExamTitle0, " " + _config.subject.Substring(3) + "统计分析");

            creating_word_part2();
        }
        public void creating_word()
        {
            object filepath = @_config.CurrentDirectory + @"\template.dotx";
            //Start Word and create a new document.

            oWord = new Word.Application();

            oWord.Visible = _config.isVisible;
            oDoc = oWord.Documents.Add(ref filepath, ref oMissing,
            ref oMissing, ref oMissing);
            Utils.WriteFrontPage(_config, oDoc, _schoolname);

            creating_word_part2();
        }
        public void creating_word_part2()
        {
            
            insertText(ExamTitle1, "总体分析");
            insertTotalTable("    总分分析表", _pdata);

            ChartCombine chartdata = new ChartCombine();
            foreach (PartitionData temp in _pdata)
            {
                chartdata.Add(temp.total_dist, temp.name);
            }

            insertChart("    总分分布曲线图", chartdata.target, "分数", "比率（%）", Excel.XlChartType.xlLineMarkers, ((PartitionData)_pdata[0]).fullmark);
            insertTotalTopic("    题目整体分析表", _pdata);
            insertGroupTopic("    题组整体分析表", _pdata, ((PartitionData)_pdata[2]).groups_analysis.Rows.Count);

            insertFreqTable("    总分频数分布表", _pdata);
            oDoc.Characters.Last.InsertBreak(oPagebreak);
            insertText(ExamTitle1, "题组分析");
            List<string> keys = new List<string>(_groups_group.Keys);
            int group_count = 0;
            for (int i = 0; i < _groups_group.Count; i++)
            {
                insertText(ExamTitle2, keys[i]);
                foreach (string group in _groups_group[keys[i]])
                {
                    if (group.Equals("totalmark"))
                        continue;
                    insertText(ExamTitle3, group);
                    insertTH(_groups.Rows[group_count]["th"].ToString().Trim());
                    insertSingleGroupTotal("    " + group + "总分分析表", group_count, true, _pdata);
                    ChartCombine tempdata = new ChartCombine();
                    for (int j = 0; j < _pdata.Count; j++)
                    {
                        tempdata.Add(getGroupData(_pdata, j, group_count).group_dist, ((PartitionData)_pdata[j]).name);
                    }
                    insertChart("    " + group + "分数分布图", tempdata.target, "分数", "比率(%)", Excel.XlChartType.xlLineMarkers, (decimal)((PartitionData)_pdata[0]).groups_analysis.Rows[group_count]["fullmark"]);
                    insertGroupDiffChart("    " + group + "难度曲线图", ((WordData.group_data)_sdata.single_group_analysis[group_count]).group_difficulty);
                    insertGroupSingleAnalysis("    " + group + "分组分析表", ((WordData.group_data)_sdata.single_group_analysis[group_count]).group_detail);
                    oDoc.Characters.Last.InsertBreak(oPagebreak);
                    group_count++;
                }
            }
            
            insertText(ExamTitle1, "题目分析");
            for (int i = 0; i < ((PartitionData)_pdata[2]).total_analysis.Rows.Count; i++)
            {
                string numstr = "第" + getSingleName(2, i).Substring(1) + "题";
                insertText(ExamTitle3, numstr);
                insertSingleGroupTotal("    " + numstr + "分析表", i, false, _pdata);
                insertSingleChart("    " + numstr + "难度曲线图", ((WordData.single_data)_sdata.single_topic_analysis[i]).single_difficulty, "分数", "难度", Excel.XlChartType.xlXYScatterSmooth);
                insertMultipleChart("    " + numstr + "分组难度曲线图", ((WordData.single_data)_sdata.single_topic_analysis[i]).single_dist, "分组", "难度", Excel.XlChartType.xlLineMarkers);
                insertGroupTable("    " + numstr + "分组分析表", ((WordData.single_data)_sdata.single_topic_analysis[i]).single_detail, ((WordData.single_data)_sdata.single_topic_analysis[i]).stype);
                oDoc.Characters.Last.InsertBreak(oPagebreak);
            }

            foreach (Word.TableOfContents table in oDoc.TablesOfContents)
                table.Update();
            Utils.Save(_config, oDoc, oWord, _schoolname);
        }
        public void insertGroupTable(string title, DataTable dt, WordData.single_type type)
        {

            Word.Table table;
            Word.Range range = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;

            table = oDoc.Tables.Add(range, dt.Rows.Count + 1, dt.Columns.Count, ref oMissing, oTrue);
            table.Range.InsertCaption(oWord.CaptionLabels["表"], title, oMissing, Word.WdCaptionPosition.wdCaptionPositionAbove, oMissing);
            range.MoveEnd(Word.WdUnits.wdParagraph, 1);
            range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            table.Rows[1].HeadingFormat = -1;


            table.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            table.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;

            if (type == WordData.single_type.single || type == WordData.single_type.multiple)
                table.Cell(1, 1).Range.Text = "选项";
            else
                table.Cell(1, 1).Range.Text = "分值";
            for (int i = 1; i < dt.Columns.Count; i++)
            {
                if (dt.Columns[i].ColumnName.Trim().Equals("frequency"))
                    table.Cell(1, i + 1).Range.Text = "人数";
                else if (dt.Columns[i].ColumnName.Trim().Equals("rate"))
                    table.Cell(1, i + 1).Range.Text = "比率(%)";
                else if (dt.Columns[i].ColumnName.Trim().Equals("correlation"))
                {
                    if (type == WordData.single_type.sub)
                        table.Cell(1, i + 1).Range.Text = "平均值";
                    else
                        table.Cell(1, i + 1).Range.Text = "相关系数";
                }
                else
                    table.Cell(1, i + 1).Range.Text = dt.Columns[i].ColumnName;
            }
            int row = 2;
            foreach (DataRow dr in dt.Rows)
            {
                if (dr["mark"].ToString().Trim().Equals("未选") || dr["mark"].ToString().Trim().Equals("未选或多选"))
                {
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        if (dt.Columns[j].ColumnName.Trim().Equals("rate") ||
                                dt.Columns[j].ColumnName.Trim().Equals("correlation"))
                            table.Cell(dt.Rows.Count - 1, j + 1).Range.Text = string.Format("{0:F2}", dr[j]);
                        else if (dt.Columns[j].ColumnName.Trim().Equals("mark"))
                            table.Cell(dt.Rows.Count - 1, j + 1).Range.Text = dr[j].ToString().Trim();
                        else if (dt.Columns[j].ColumnName.Trim().Equals("frequency"))
                            table.Cell(dt.Rows.Count - 1, j + 1).Range.Text = dr[j].ToString();
                        else
                            table.Cell(dt.Rows.Count - 1, j + 1).Range.Text = Convert.ToInt32(dr[j]).ToString().Trim();
                    }
                }
                else if (dr["mark"].ToString().Trim().Equals("合计"))
                {
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        if (dt.Columns[j].ColumnName.Trim().Equals("rate"))

                            table.Cell(dt.Rows.Count, j + 1).Range.Text = string.Format("{0:F2}", dr[j]);
                        else if (dt.Columns[j].ColumnName.Trim().Equals("correlation"))
                            table.Cell(dt.Rows.Count, j + 1).Range.Text = "-";
                        else if (dt.Columns[j].ColumnName.Trim().Equals("mark"))
                            table.Cell(dt.Rows.Count, j + 1).Range.Text = dr[j].ToString().Trim();
                        else if (dt.Columns[j].ColumnName.Trim().Equals("frequency"))
                            table.Cell(dt.Rows.Count, j + 1).Range.Text = dr[j].ToString();
                        else
                            table.Cell(dt.Rows.Count, j + 1).Range.Text = Convert.ToInt32(dr[j]).ToString().Trim();
                    }
                }
                else if (dr["mark"].ToString().Trim().Equals("得分率"))
                {
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        if (dt.Columns[j].ColumnName.Trim().Equals("rate") ||
                            dt.Columns[j].ColumnName.Trim().Equals("correlation") ||
                            dt.Columns[j].ColumnName.Trim().Equals("frequency"))
                            table.Cell(dt.Rows.Count + 1, j + 1).Range.Text = "-";
                        else if (dt.Columns[j].ColumnName.Trim().Equals("mark"))
                            table.Cell(dt.Rows.Count + 1, j + 1).Range.Text = dr[j].ToString().Trim();

                        else
                            table.Cell(dt.Rows.Count + 1, j + 1).Range.Text = string.Format("{0:F2}", dr[j]);
                    }
                }
                else
                {
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        if (dt.Columns[j].ColumnName.Trim().Equals("rate") ||
                            dt.Columns[j].ColumnName.Trim().Equals("correlation"))
                            table.Cell(row, j + 1).Range.Text = string.Format("{0:F2}", dr[j]);
                        else if (dt.Columns[j].ColumnName.Trim().Equals("mark"))
                            table.Cell(row, j + 1).Range.Text = dr[j].ToString().Trim();
                        else if (dt.Columns[j].ColumnName.Trim().Equals("frequency"))
                            table.Cell(row, j + 1).Range.Text = dr[j].ToString();
                        else
                            table.Cell(row, j + 1).Range.Text = Convert.ToInt32(dr[j]).ToString().Trim();
                    }
                    row++;
                }

            }

            table.Select();
            oWord.Selection.set_Style(ref TableContent2);
            changeStyle(table);
            range = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            range.InsertParagraphAfter();
        }
        public void insertMultipleChart(string title, DataTable dt, string x_axis, string y_axis, object type)
        {
            ZedGraph.createMultipleChoiceCuve(_config, dt, x_axis, y_axis);
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
        public void insertSingleChart(string title, DataTable dt, string x_axis, string y_axis, object type)
        {
            double[][] data = new double[dt.Rows.Count][];
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                data[i] = new double[2];
                data[i][0] = Convert.ToDouble((decimal)dt.Rows[i][0]);
                data[i][1] = Convert.ToDouble((decimal)dt.Rows[i][1]);

            }

            ZedGraph.createDiffCuve(_config, data, Convert.ToDouble(dt.Compute("Min([" + dt.Columns[0].ColumnName + "])", "")), Convert.ToDouble(dt.Compute("Max([" + dt.Columns[0].ColumnName + "])", "")));
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
        public string getSingleName(int total, int row)
        {
            return ((PartitionData)_pdata[total]).total_analysis.Rows[row]["number"].ToString().Trim();
        }
        public void insertGroupSingleAnalysis(string title, DataTable single_group_dt)
        {
            Word.Table single_group_analysis;
            Word.Range single_group_rng = oDoc.Bookmarks.get_Item(oEndOfDoc).Range;

            //DataTable single_group_dt = ((WordData.group_data)_sdata.single_group_analysis[group_num]).group_detail;
            single_group_analysis = oDoc.Tables.Add(single_group_rng, single_group_dt.Rows.Count + 1, single_group_dt.Columns.Count, ref oMissing, oTrue);

            single_group_analysis.Rows[1].HeadingFormat = -1;
            //object single_group_title = group_dr["number"].ToString().Trim() + "分组分析表";
            single_group_analysis.Range.InsertCaption(oWord.CaptionLabels["表"], title, oMissing, Word.WdCaptionPosition.wdCaptionPositionAbove, oMissing);
            single_group_rng.MoveEnd(Word.WdUnits.wdParagraph, 1);
            single_group_rng.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            //single_group_analysis.Range.set_Style(TableContent2);
            //single_group_analysis.Range.Select();
            //oWord.Selection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            //oWord.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            //single_group_analysis.Range.ParagraphFormat.Space1();
            single_group_analysis.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            single_group_analysis.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;

            //single_group_analysis.Range.Paragraphs.OutlineLevel = Word.WdOutlineLevel.wdOutlineLevelBodyText;
            //single_group_analysis.Range.Font.Size = 10;
            //single_group_analysis.Range.Font.Name = "黑体";
            int i;
            int j = 0;
            for (i = 0; i < single_group_dt.Columns.Count; i++)
            {
                if (single_group_dt.Columns[i].ColumnName.Trim().Equals("mark"))
                    single_group_analysis.Cell(1, i + 1).Range.Text = "分值";
                else if (single_group_dt.Columns[i].ColumnName.Trim().Equals("rate"))
                    single_group_analysis.Cell(1, i + 1).Range.Text = "比率(%)";
                else if (single_group_dt.Columns[i].ColumnName.Trim().Equals("avg"))
                    single_group_analysis.Cell(1, i + 1).Range.Text = "平均值";
                else if (single_group_dt.Columns[i].ColumnName.Trim().Equals("frequency"))
                    single_group_analysis.Cell(1, i + 1).Range.Text = "人数";
                else
                    single_group_analysis.Cell(1, i + 1).Range.Text = single_group_dt.Columns[i].ColumnName;
            }
            for (i = 0; i < single_group_dt.Rows.Count; i++)
            {

                for (j = 0; j < single_group_dt.Columns.Count; j++)
                {
                    if (i == single_group_dt.Rows.Count - 1)
                    {
                        if (single_group_dt.Columns[j].ColumnName.Trim().Equals("rate") ||
                        single_group_dt.Columns[j].ColumnName.Trim().Equals("avg") || single_group_dt.Columns[j].ColumnName.Trim().Equals("frequency"))
                            single_group_analysis.Cell(i + 2, j + 1).Range.Text = "-";
                        else if (single_group_dt.Columns[j].ColumnName.Trim().Equals("mark"))
                            single_group_analysis.Cell(i + 2, j + 1).Range.Text = single_group_dt.Rows[i][j].ToString().Trim();

                        else
                            single_group_analysis.Cell(i + 2, j + 1).Range.Text = string.Format("{0:F2}", single_group_dt.Rows[i][j]);
                    }
                    else
                    {
                        if (single_group_dt.Columns[j].ColumnName.Trim().Equals("rate") ||
                            single_group_dt.Columns[j].ColumnName.Trim().Equals("avg"))
                            single_group_analysis.Cell(i + 2, j + 1).Range.Text = string.Format("{0:F2}", single_group_dt.Rows[i][j]);
                        else if (single_group_dt.Columns[j].ColumnName.Trim().Equals("mark"))
                            single_group_analysis.Cell(i + 2, j + 1).Range.Text = single_group_dt.Rows[i][j].ToString().Trim();
                        else if (single_group_dt.Columns[j].ColumnName.Trim().Equals("frequency"))
                            single_group_analysis.Cell(i + 2, j + 1).Range.Text = single_group_dt.Rows[i][j].ToString();
                        else
                            single_group_analysis.Cell(i + 2, j + 1).Range.Text = Convert.ToInt32(single_group_dt.Rows[i][j]).ToString().Trim();
                    }



                }
            }
            single_group_analysis.Select();
            oWord.Selection.set_Style(ref TableContent2);
            Word.Range single_table_range = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            single_table_range.InsertParagraphAfter();
        }
        public void insertGroupDiffChart(string title, DataTable difficulty_data)
        {
            DataTable dt = difficulty_data;
            double[][] data = new double[dt.Rows.Count][];
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                data[i] = new double[2];
                data[i][0] = Convert.ToDouble((decimal)dt.Rows[i][0]);
                data[i][1] = Convert.ToDouble((decimal)dt.Rows[i][1]);

            }

            ZedGraph.createDiffCuve(_config, data, Convert.ToDouble(dt.Compute("Min([" + dt.Columns[0].ColumnName + "])", "")), Convert.ToDouble(dt.Compute("Max([" + dt.Columns[0].ColumnName + "])", "")));
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
        public PartitionData.group_data getGroupData(List<WSLG_partitiondata> sdata, int total, int group)
        {
            return (PartitionData.group_data)((PartitionData)sdata[total]).single_group_analysis[group];
        }
        public void insertSingleGroupTotal(string title, int total, bool isGroup, List<WSLG_partitiondata> sdata)
        {
            Word.Table table;
            Word.Range range = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            
            table = oDoc.Tables.Add(range, sdata.Count + 1, 10, ref oMissing, oTrue);
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

            for (int i = 0; i < sdata.Count; i++)
            {
                DataTable data;
                PartitionData partition = (PartitionData)sdata[i];
                if (isGroup)
                    data = partition.groups_analysis;
                else
                    data = partition.total_analysis;
                table.Cell(i + 2, 1).Range.Text = partition.name;
                table.Cell(i + 2, 2).Range.Text = isGroup ? partition.total_num.ToString() : data.Rows[total]["total_num"].ToString();
                table.Cell(i + 2, 3).Range.Text = FullmarkFormat((decimal)data.Rows[total]["fullmark"]);
                table.Cell(i + 2, 4).Range.Text = string.Format("{0:F1}", data.Rows[total]["max"]);
                table.Cell(i + 2, 5).Range.Text = string.Format("{0:F1}", data.Rows[total]["min"]);
                table.Cell(i + 2, 6).Range.Text = string.Format("{0:F2}", data.Rows[total]["avg"]);
                table.Cell(i + 2, 7).Range.Text = string.Format("{0:F2}", data.Rows[total]["stDev"]);
                table.Cell(i + 2, 8).Range.Text = string.Format("{0:F2}", data.Rows[total]["dfactor"]);
                table.Cell(i + 2, 9).Range.Text = string.Format("{0:F2}", data.Rows[total]["difficulty"]);
                
                    if (isGroup)
                        table.Cell(i + 2, 10).Range.Text = string.Format("{0:F2}", ((WSLG_partitiondata)partition).group_discriminant[total]);
                    else
                        table.Cell(i + 2, 10).Range.Text = string.Format("{0:F2}", ((WSLG_partitiondata)partition).total_discriminant[total]);
                
            }
            table.Select();
            oWord.Selection.set_Style(ref TableContent2);
            range = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            range.InsertParagraphAfter();
        }

        public void insertFreqTable(string title, List<WSLG_partitiondata> sdata)
        {
            int count = ((PartitionData)sdata[2]).freq_analysis.Rows.Count;
            Word.Table table;
            Word.Range range = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;

            table = oDoc.Tables.Add(range, sdata.Count * count + 1, 6, ref oMissing, oTrue);
            table.Range.InsertCaption(oWord.CaptionLabels["表"], title, oMissing, Word.WdCaptionPosition.wdCaptionPositionAbove, oMissing);
            range.MoveEnd(Word.WdUnits.wdParagraph, 1);
            range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            table.Rows[1].HeadingFormat = -1;


            table.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            table.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;

            table.Cell(1, 1).Range.Text = "分值";
            table.Cell(1, 2).Range.Text = "分类";
            table.Cell(1, 3).Range.Text = "人数";
            table.Cell(1, 4).Range.Text = "比率(%)";
            table.Cell(1, 5).Range.Text = "累计人数";
            table.Cell(1, 6).Range.Text = "累计比率(%)";

            string[] lastFreq = new string[sdata.Count];
            string[] lastRate = new string[sdata.Count];
            for (int i = 0; i < sdata.Count; i++)
            {
                lastFreq[i] = "0";
                lastRate[i] = "0.00";
            }
            for (int i = 0; i < count; i++)
            {
                int tablerow = i * sdata.Count + 2;
                decimal primarykey = (decimal)((PartitionData)sdata[2]).freq_analysis.Rows[i]["totalmark"];
                int j = 0;
                foreach (PartitionData data in sdata)
                {
                    if (data.freq_analysis.Rows.Contains(primarykey))
                    {
                        table.Cell(tablerow, 1).Range.Text = string.Format("{0:F0}", data.freq_analysis.Rows.Find(primarykey)["totalmark"]) + "～";
                        table.Cell(tablerow, 2).Range.Text = data.name;
                        table.Cell(tablerow, 3).Range.Text = data.freq_analysis.Rows.Find(primarykey)["frequency"].ToString().Trim();
                        table.Cell(tablerow, 4).Range.Text = string.Format("{0:F2}", data.freq_analysis.Rows.Find(primarykey)["rate"]);
                        table.Cell(tablerow, 5).Range.Text = data.freq_analysis.Rows.Find(primarykey)["accumulateFreq"].ToString().Trim();
                        table.Cell(tablerow, 6).Range.Text = string.Format("{0:F2}", data.freq_analysis.Rows.Find(primarykey)["accumulateRate"]);

                        lastFreq[j] = data.freq_analysis.Rows.Find(primarykey)["accumulateFreq"].ToString().Trim();
                        lastRate[j] = string.Format("{0:F2}", data.freq_analysis.Rows.Find(primarykey)["accumulateRate"]);
                    }
                    else
                    {
                        table.Cell(tablerow, 1).Range.Text = string.Format("{0:F0}", primarykey) + "～";
                        table.Cell(tablerow, 2).Range.Text = data.name;
                        table.Cell(tablerow, 3).Range.Text = "0";
                        table.Cell(tablerow, 4).Range.Text = "0.00";
                        table.Cell(tablerow, 5).Range.Text = lastFreq[j];
                        table.Cell(tablerow, 6).Range.Text = lastRate[j];
                    }
                    tablerow++;
                    j++;
                }
            }

            table.Select();
            oWord.Selection.set_Style(ref TableContent2);

            verticalCellMerge(table, 2, 1);


            range = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            range.InsertParagraphAfter();
        }

        public void insertGroupTopic(string title, List<WSLG_partitiondata> totaldata, int count)
        {
            //int count = ((PartitionData)totaldata[totaldata.Count - 1]).groups_analysis.Rows.Count;
            Word.Table table;
            Word.Range range = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;

            table = oDoc.Tables.Add(range, totaldata.Count * count + 1, 11, ref oMissing, oTrue);
            table.Range.InsertCaption(oWord.CaptionLabels["表"], title, oMissing, Word.WdCaptionPosition.wdCaptionPositionAbove, oMissing);
            range.MoveEnd(Word.WdUnits.wdParagraph, 1);
            range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            table.Rows[1].HeadingFormat = -1;


            table.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            table.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;

            table.Cell(1, 1).Range.Text = "题组";
            table.Cell(1, 2).Range.Text = "分类";
            table.Cell(1, 3).Range.Text = "人数";
            table.Cell(1, 4).Range.Text = "满分值";
            table.Cell(1, 5).Range.Text = "最大值";
            table.Cell(1, 6).Range.Text = "最小值";
            table.Cell(1, 7).Range.Text = "平均值";
            table.Cell(1, 8).Range.Text = "标准差";
            table.Cell(1, 9).Range.Text = "差异系数";
            table.Cell(1, 10).Range.Text = "得分率";
            table.Cell(1, 11).Range.Text = "鉴别指数";

            for (int i = 0; i < count; i++)
            {
                int tablerow = i * totaldata.Count + 2;
                foreach (PartitionData data in totaldata)
                {
                    table.Cell(tablerow, 1).Range.Text = data._group_ans.Rows[i][0].ToString().Trim();
                    table.Cell(tablerow, 2).Range.Text = data.name;
                    table.Cell(tablerow, 3).Range.Text = data.total_num.ToString();
                    table.Cell(tablerow, 4).Range.Text = FullmarkFormat((decimal)data.groups_analysis.Rows[i]["fullmark"]);
                    table.Cell(tablerow, 5).Range.Text = string.Format("{0:F1}", data.groups_analysis.Rows[i]["max"]);
                    table.Cell(tablerow, 6).Range.Text = string.Format("{0:F1}", data.groups_analysis.Rows[i]["min"]);
                    table.Cell(tablerow, 7).Range.Text = string.Format("{0:F2}", data.groups_analysis.Rows[i]["avg"]);
                    table.Cell(tablerow, 8).Range.Text = string.Format("{0:F2}", data.groups_analysis.Rows[i]["stDev"]);
                    table.Cell(tablerow, 9).Range.Text = string.Format("{0:F2}", data.groups_analysis.Rows[i]["dfactor"]);
                    table.Cell(tablerow, 10).Range.Text = string.Format("{0:F2}", data.groups_analysis.Rows[i]["difficulty"]);
                    table.Cell(tablerow, 11).Range.Text = string.Format("{0:F2}", ((WSLG_partitiondata)data).group_discriminant[i]);
                    tablerow++;
                }
            }
            table.Select();
            oWord.Selection.set_Style(ref TableContent2);

            verticalCellMerge(table, 2, 1);


            range = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            range.InsertParagraphAfter();
        }
        public void insertTotalTopic(string title, List<WSLG_partitiondata> totaldata)
        {
            int count = ((PartitionData)totaldata[totaldata.Count - 1]).total_analysis.Rows.Count;
            Word.Table table;
            Word.Range range = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            
            table = oDoc.Tables.Add(range, totaldata.Count * count + 1, 11, ref oMissing, oTrue);
            table.Range.InsertCaption(oWord.CaptionLabels["表"], title, oMissing, Word.WdCaptionPosition.wdCaptionPositionAbove, oMissing);
            range.MoveEnd(Word.WdUnits.wdParagraph, 1);
            range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            table.Rows[1].HeadingFormat = -1;


            table.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            table.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;

            table.Cell(1, 1).Range.Text = "题目";
            table.Cell(1, 2).Range.Text = "分类";
            table.Cell(1, 3).Range.Text = "人数";
            table.Cell(1, 4).Range.Text = "满分值";
            table.Cell(1, 5).Range.Text = "最大值";
            table.Cell(1, 6).Range.Text = "最小值";
            table.Cell(1, 7).Range.Text = "平均值";
            table.Cell(1, 8).Range.Text = "标准差";
            table.Cell(1, 9).Range.Text = "差异系数";
            table.Cell(1, 10).Range.Text = "得分率";
            table.Cell(1, 11).Range.Text = "鉴别指数";

            for (int i = 0; i < count; i++)
            {
                int tablerow = i * totaldata.Count + 2;
                foreach (PartitionData data in totaldata)
                {
                    table.Cell(tablerow, 1).Range.Text = data.total_analysis.Rows[i]["number"].ToString().Trim().Substring(1);
                    table.Cell(tablerow, 2).Range.Text = data.name;
                    table.Cell(tablerow, 3).Range.Text = ((int)data.total_analysis.Rows[i]["total_num"]).ToString();
                    table.Cell(tablerow, 4).Range.Text = FullmarkFormat((decimal)data.total_analysis.Rows[i]["fullmark"]);
                    table.Cell(tablerow, 5).Range.Text = string.Format("{0:F1}", data.total_analysis.Rows[i]["max"]);
                    table.Cell(tablerow, 6).Range.Text = string.Format("{0:F1}", data.total_analysis.Rows[i]["min"]);
                    table.Cell(tablerow, 7).Range.Text = string.Format("{0:F2}", data.total_analysis.Rows[i]["avg"]);
                    table.Cell(tablerow, 8).Range.Text = string.Format("{0:F2}", data.total_analysis.Rows[i]["stDev"]);
                    table.Cell(tablerow, 9).Range.Text = string.Format("{0:F2}", data.total_analysis.Rows[i]["dfactor"]);
                    table.Cell(tablerow, 10).Range.Text = string.Format("{0:F2}", data.total_analysis.Rows[i]["difficulty"]);
                    table.Cell(tablerow, 11).Range.Text = string.Format("{0:F2}", ((WSLG_partitiondata)data).total_discriminant[i]);
                    tablerow++;
                }
            }
            table.Select();
            oWord.Selection.set_Style(ref TableContent2);

            verticalCellMerge(table, 2, 1);


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
            dist_rng.InsertCaption(oWord.CaptionLabels["图"], title, oMissing, Word.WdCaptionPosition.wdCaptionPositionAbove, oMissing);
            dist_rng.MoveEnd(Word.WdUnits.wdParagraph, 1);
            dist_rng.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            dist_rng = oDoc.Bookmarks.get_Item(oEndOfDoc).Range;
            dist_rng.MoveStart(Word.WdUnits.wdParagraph, 1);
            dist_rng.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            dist_rng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            dist_rng.InsertParagraphAfter();
            
        }

        public void insertTotalTable(string title, List<WSLG_partitiondata> totaldata)
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
        public class ChartCombine
        {
            public DataTable target;
            int num = 2;
            public void Add(DataTable dt, string name)
            {
                if (target == null)
                {
                    target = dt.Copy();
                    target.Columns[0].ColumnName = "mark";
                    target.Columns[1].ColumnName = name;

                    return;
                }
                target.Columns.Add(name, typeof(decimal));
                foreach (DataRow dr in target.Rows)
                    dr[name] = 0;
                target.PrimaryKey = new DataColumn[] { target.Columns["mark"] };
                foreach (DataRow dr in dt.Rows)
                {
                    if (target.Rows.Contains(dr["mark"]))
                    {
                        DataRow oldrow = target.Rows.Find(dr["mark"]);
                        oldrow[name] = dr["rate"];
                    }
                    else
                    {
                        DataRow newrow = target.NewRow();
                        newrow["mark"] = dr["mark"];
                        for (int i = 1; i < dt.Columns.Count; i++)
                            newrow[i] = 0;
                        newrow[name] = dr["rate"];
                    }
                }
                DataView dv = target.DefaultView;
                dv.Sort = "mark";
                target = dv.ToTable();
                num++;
            }
        }
        void insertTH(string TH)
        {
            Word.Range first = oDoc.Bookmarks.get_Item(oEndOfDoc).Range;
            first.set_Style(ExamBodyText);
            first.InsertAfter("本题组共包含");
            int count = 0;
            string[] th_string = TH.ToString().Trim().Split(new char[2] { ',', '，' });
            foreach (string temp in th_string)
            {
                if (System.Text.RegularExpressions.Regex.IsMatch(temp, "^\\d+~\\d+$"))
                //if(th.Contains('~'))
                {
                    string[] num = temp.Split('~');
                    int j;
                    int size = Convert.ToInt32(num[0]) < Convert.ToInt32(num[1]) ? Convert.ToInt32(num[1]) : Convert.ToInt32(num[0]);
                    int start = Convert.ToInt32(num[0]) > Convert.ToInt32(num[1]) ? Convert.ToInt32(num[1]) : Convert.ToInt32(num[0]);
                    //此处需判断size和start的边界问题
                    for (j = start; j < size + 1; j++)
                    {
                        count++;
                    }

                }
                else
                    count++;
            }
            first.InsertAfter(count.ToString());
            first.InsertAfter("道试题，题号是：");
            for (int i = 0; i < th_string.Length; i++)
            {
                if (System.Text.RegularExpressions.Regex.IsMatch(th_string[i], "^\\d+~\\d+$"))
                {
                    string[] num = th_string[i].Split('~');
                    first.InsertAfter(num[0]);
                    first.InsertAfter("～");
                    first.InsertAfter(num[1]);
                }
                else
                    first.InsertAfter(th_string[i]);
                if (i < th_string.Length - 1)
                    first.InsertAfter("、");
            }
            first.InsertParagraphAfter();

            //oDoc.Characters.Last.Select();
            //oWord.Selection.HomeKey(Word.WdUnits.wdLine, oMissing);
            //oWord.Selection.Delete(Word.WdUnits.wdCharacter, oMissing);
            //oWord.Selection.Range.set_Style(ExamBodyText);
        }

        void changeStyle(Word.Table table)
        {
            for (int i = 1; i <= table.Rows.Count; i++)
            {
                if (table.Cell(i, 1).Range.Text.Equals("未选或多选\r\a"))
                {
                    table.Cell(i, 1).Range.Select();
                    oWord.Selection.Range.Font.Size = 9f;
                }
            }
        }
    }
}
