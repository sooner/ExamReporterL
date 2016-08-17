using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Data;
using dotnetCHARTING.WinForms;
using Chart = dotnetCHARTING.WinForms.Chart;
using Series = dotnetCHARTING.WinForms.Series;
using SeriesCollection = dotnetCHARTING.WinForms.SeriesCollection;
using System.Windows.Forms;


namespace ExamReport
{
    public static class DotNetCharting
    {
        public static Color[] color = new Color[] { Color.FromArgb(0, 156, 255), Color.FromArgb(255, 99, 49), Color.FromArgb(49, 255, 49), Color.FromArgb(255, 255, 0), };

        public static void CreateColumn( DataTable dt)
        {
            
                Chart chart = new Chart();
                //清空图片
                chart.SeriesCollection.Clear();
                //标题框设置
                //标题的颜色
                chart.TitleBox.Label.Color = Color.Black;
                //标题字体设置
                chart.TitleBox.Label.Font = new Font("Microsoft Sans Serif", 9F, FontStyle.Bold, GraphicsUnit.Point, 134); ;

                //控制柱状图颜色
                chart.ShadingEffectMode = ShadingEffectMode.One;

                chart.TitleBox.Position = TitleBoxPosition.None;

                //图表背景颜色
                chart.ChartArea.Background.Color = Color.White;
                //1.图表类型
                chart.DefaultSeries.Type = SeriesType.Column ;// SeriesType.Column;
                //chart.DefaultSeries.Type = SeriesType.Cylinder;
                //2.图表类型
                //柱状图
                //chart.Type = ChartType.TreeMap;
                ////横向柱状图
                chart.Type = ChartType.Combo;// ChartType.ComboHorizontal
                ////横向柱状图
                //chart.Type =_chartType;// ChartType.Gantt;
                ////饼状图
                //chart.Type = ChartType.Pies;

                //y轴图表阴影颜色
                //chart.YAxis.AlternateGridBackground.Color = Color.FromArgb(255, 250, 250, 250);

                chart.LegendBox.HeaderLabel = new dotnetCHARTING.WinForms.Label("图表说明", new Font("Microsoft Sans Serif", 10F, FontStyle.Bold, GraphicsUnit.Point, 134));
                //chart.LegendBox.HeaderLabel.Font = new Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 134);
                chart.LegendBox.Label.Font = new Font("Microsoft Sans Serif", 9F, FontStyle.Bold, GraphicsUnit.Point, 134);
                chart.Palette = new Color[] { Color.FromArgb(0, 156, 255), Color.FromArgb(255, 99, 49), Color.FromArgb(49, 255, 49), Color.FromArgb(255, 255, 0), };

                chart.Title = "";
                //X轴柱状图值字体
                chart.XAxis.Label.Text = "";
                chart.XAxis.Label.Font = new Font("Microsoft Sans Serif", 9F, FontStyle.Bold, GraphicsUnit.Point, 134);
                //设置X轴刻度值说明字体
                chart.XAxis.DefaultTick.Label.Font = new Font("Microsoft Sans Serif", 9F, FontStyle.Bold, GraphicsUnit.Point, 134); ;

                chart.YAxis.StaticColumnWidth = 3; //每个单元格的宽度


                //Y轴柱状图值字体
                chart.YAxis.Label.Text = "";
                chart.YAxis.Label.Font = new Font("Microsoft Sans Serif", 9F, FontStyle.Bold, GraphicsUnit.Point, 134); ;
                //设置Y轴刻度值说明字体
                chart.YAxis.DefaultTick.Label.Font = new Font("Microsoft Sans Serif", 9F, FontStyle.Bold, GraphicsUnit.Point, 134); ;

                //Y轴箭头标示
                chart.XAxis.Name = "";
                if (chart.Type == ChartType.ComboHorizontal)
                {
                    chart.XAxis.TickLabelPadding = 10;
                    chart.XAxis.Line.StartCap = System.Drawing.Drawing2D.LineCap.Square;
                    chart.XAxis.Line.EndCap = System.Drawing.Drawing2D.LineCap.ArrowAnchor;
                    chart.XAxis.Line.Width = 5;//箭头宽度
                    chart.XAxis.Line.Color = Color.Gray;
                }
                else
                {
                    chart.YAxis.TickLabelPadding = 10;
                    chart.YAxis.Line.StartCap = System.Drawing.Drawing2D.LineCap.Square;
                    chart.YAxis.Line.EndCap = System.Drawing.Drawing2D.LineCap.ArrowAnchor;
                    chart.YAxis.Line.Width = 5;//箭头宽度
                    chart.YAxis.Line.Color = Color.Gray;
                    //显示值格式化(小数点显示几位)
                    chart.YAxis.NumberPercision = 2;
                }

                //图片存放路径
                //chart.TempDirectory = System.Environment.CurrentDirectory + "\\";
                //图表宽度
                chart.Width = 460;
                //图表高度
                chart.Height = 291;
                chart.Series.Name = "";
                //单一图形
                //chart.Series.Data = _dt;
                //chart.SeriesCollection.Add();

                //图例在标题行显示，但是没有合计信息
                //chart.TitleBox.Position = TitleBoxPosition.FullWithLegend;
                //chart.TitleBox.Label.Alignment = StringAlignment.Center;
                //chart.LegendBox.Position = LegendBoxPosition.None; //不显示图例,指不在右侧显示，对上面一行的属性设置并没有影响
                chart.LegendBox.Visible = false;
                chart.DefaultSeries.DefaultElement.ShowValue = true;
                chart.ShadingEffect = true;
                chart.Use3D = true;
                chart.Series.DefaultElement.ShowValue = true;
                chart.SeriesCollection.Add(GetArrayData(dt, "", color[0]));
                Bitmap sourceBitmap = new Bitmap(chart.Width, chart.Height);
                chart.DrawToBitmap(sourceBitmap, new Rectangle(0, 0, chart.Width, chart.Height));
                //Thread.CurrentThread.SetApartmentState(ApartmentState.STA);
                Utils.mutex_clipboard.WaitOne();
                Clipboard.Clear();
                Clipboard.SetImage(sourceBitmap);

            
        }

        private static SeriesCollection GetArrayData(DataTable dt, string name, Color my_color)
        {
            SeriesCollection sc = new SeriesCollection();
            try
            { 
                Series s = new Series();
                
                s.Name = name;
                s.Palette = new Color[] {my_color};
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    
                    var e = new Element();

                    // 每元素的名称
                    e.Name = dt.Rows[i][0].ToString();
                    //设置柱状图值的字体
                    e.SmartLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 6F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 134);
                    //e.SmartLabel.DynamicDisplay = true;
                    //e.SmartLabel.AutoWrap = true;
                    // 每元素的大小数值
                    e.YValue = Convert.ToDouble(dt.Rows[i][1].ToString());
                    //调整柱子颜色
                    //s.PaletteName = Palette.Three;
                    
                    //s.Palette = new Color[] { Color.FromArgb(16, 109, 156), Color.FromArgb(90, 146, 173), Color.FromArgb(0, 162, 222), Color.FromArgb(8, 186, 255), };
                    //s.Palette = new Color[] { Color.FromArgb(82, 89, 107), Color.FromArgb(189, 32, 16), Color.FromArgb(231, 186, 16), Color.FromArgb(99, 150, 41), Color.FromArgb(156, 85, 173), Color.FromArgb(206, 195, 198) };

                    s.Elements.Add(e);
                }
                sc.Add(s);
                return sc;

            }
            catch (Exception ex)
            {

                return sc;
            }
        }


        public static void CreateMutipleColumn(Dictionary<string, DataTable> dts)
        {

            Chart chart = new Chart();
            //清空图片
            chart.SeriesCollection.Clear();
            //标题框设置
            //标题的颜色
            chart.TitleBox.Label.Color = Color.Black;
            //标题字体设置
            chart.TitleBox.Label.Font = new Font("Microsoft Sans Serif", 9F, FontStyle.Bold, GraphicsUnit.Point, 134); ;

            //控制柱状图颜色
            chart.ShadingEffectMode = ShadingEffectMode.One;

            chart.TitleBox.Position = TitleBoxPosition.None;

            //图表背景颜色
            chart.ChartArea.Background.Color = Color.White;
            //1.图表类型
            chart.DefaultSeries.Type = SeriesType.Column;// SeriesType.Column;
            //chart.DefaultSeries.Type = SeriesType.Cylinder;
            //2.图表类型
            //柱状图
            //chart.Type = ChartType.TreeMap;
            ////横向柱状图
            chart.Type = ChartType.Combo;// ChartType.ComboHorizontal
            ////横向柱状图
            //chart.Type =_chartType;// ChartType.Gantt;
            ////饼状图
            //chart.Type = ChartType.Pies;

            //y轴图表阴影颜色
            //chart.YAxis.AlternateGridBackground.Color = Color.FromArgb(255, 250, 250, 250);

            chart.LegendBox.HeaderLabel = new dotnetCHARTING.WinForms.Label("图表说明", new Font("Microsoft Sans Serif", 10F, FontStyle.Bold, GraphicsUnit.Point, 134));
            //chart.LegendBox.HeaderLabel.Font = new Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 134);
            chart.LegendBox.Label.Font = new Font("Microsoft Sans Serif", 9F, FontStyle.Bold, GraphicsUnit.Point, 134);
            chart.Palette = new Color[] { Color.FromArgb(0, 156, 255), Color.FromArgb(255, 99, 49), Color.FromArgb(49, 255, 49), Color.FromArgb(255, 255, 0), };

            chart.Title = "";
            //X轴柱状图值字体
            chart.XAxis.Label.Text = "";
            chart.XAxis.Label.Font = new Font("Microsoft Sans Serif", 9F, FontStyle.Bold, GraphicsUnit.Point, 134);
            //设置X轴刻度值说明字体
            chart.XAxis.DefaultTick.Label.Font = new Font("Microsoft Sans Serif", 9F, FontStyle.Bold, GraphicsUnit.Point, 134); ;

            chart.YAxis.StaticColumnWidth = 3; //每个单元格的宽度


            //Y轴柱状图值字体
            chart.YAxis.Label.Text = "";
            chart.YAxis.Label.Font = new Font("Microsoft Sans Serif", 9F, FontStyle.Bold, GraphicsUnit.Point, 134); ;
            //设置Y轴刻度值说明字体
            chart.YAxis.DefaultTick.Label.Font = new Font("Microsoft Sans Serif", 9F, FontStyle.Bold, GraphicsUnit.Point, 134); ;

            //Y轴箭头标示
            chart.XAxis.Name = "";
            if (chart.Type == ChartType.ComboHorizontal)
            {
                chart.XAxis.TickLabelPadding = 10;
                chart.XAxis.Line.StartCap = System.Drawing.Drawing2D.LineCap.Square;
                chart.XAxis.Line.EndCap = System.Drawing.Drawing2D.LineCap.ArrowAnchor;
                chart.XAxis.Line.Width = 5;//箭头宽度
                chart.XAxis.Line.Color = Color.Gray;
            }
            else
            {
                chart.YAxis.TickLabelPadding = 10;
                chart.YAxis.Line.StartCap = System.Drawing.Drawing2D.LineCap.Square;
                chart.YAxis.Line.EndCap = System.Drawing.Drawing2D.LineCap.ArrowAnchor;
                chart.YAxis.Line.Width = 5;//箭头宽度
                chart.YAxis.Line.Color = Color.Gray;
                //显示值格式化(小数点显示几位)
                chart.YAxis.NumberPercision = 2;
            }

            //图片存放路径
            //chart.TempDirectory = System.Environment.CurrentDirectory + "\\";
            //图表宽度
            chart.Width = 460;
            //图表高度
            chart.Height = 291;
            chart.Series.Name = "";
            //单一图形
            //chart.Series.Data = _dt;
            //chart.SeriesCollection.Add();

            //图例在标题行显示，但是没有合计信息
            //chart.TitleBox.Position = TitleBoxPosition.FullWithLegend;
            //chart.TitleBox.Label.Alignment = StringAlignment.Center;
            //chart.LegendBox.Position = LegendBoxPosition.None; //不显示图例,指不在右侧显示，对上面一行的属性设置并没有影响
            chart.LegendBox.Visible = true;
            chart.DefaultSeries.DefaultElement.ShowValue = true;
            chart.ShadingEffect = true;
            chart.Use3D = true;
            chart.Series.DefaultElement.ShowValue = true;
            int count = 0;
            foreach (var kv in dts)
            {
                chart.SeriesCollection.Add(GetArrayData(kv.Value, kv.Key, color[count % 5]));
                count++;
            }
            
            Bitmap sourceBitmap = new Bitmap(chart.Width, chart.Height);
            chart.DrawToBitmap(sourceBitmap, new Rectangle(0, 0, chart.Width, chart.Height));
            //Thread.CurrentThread.SetApartmentState(ApartmentState.STA);
            Utils.mutex_clipboard.WaitOne();
            Clipboard.Clear();
            Clipboard.SetImage(sourceBitmap);


        }
 
    }
}
