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
        public static Color[] color = new Color[] {Color.LightSkyBlue, Color.LightSeaGreen, Color.FromArgb(0, 156, 255), Color.FromArgb(255, 255, 0), Color.FromArgb(0, 156, 255), Color.FromArgb(255, 99, 49), Color.FromArgb(49, 255, 49) };
        public static string CreateColumn_wh(DataTable dt, int weight, int height, string title, bool ishorizontal, int columnwidth, bool iszf)
        {
            Chart chart = new Chart();
            //清空图片
            chart.SeriesCollection.Clear();
            //标题框设置
            //标题的颜色
            chart.TitleBox.Label.Color = Color.Black;
            //标题字体设置
            chart.DefaultAxis.StaticColumnWidth = columnwidth;
            chart.TitleBox.Label.Font = new Font("Times New Roman", 7F, FontStyle.Bold, GraphicsUnit.Point, 134); ;

            //控制柱状图颜色
            chart.ShadingEffectMode = ShadingEffectMode.Three;

            chart.TitleBox.Position = TitleBoxPosition.Center;
            chart.TitleBox.Background = new Background(System.Drawing.Color.Transparent);
            chart.TitleBox.Line = new Line(Color.Transparent);
            chart.TitleBox.Shadow.Visible = false;

            //图表背景颜色
            chart.ChartArea.Background.Color = Color.White;
            chart.ChartArea.Line = new Line(Color.Gray, 1);
            //1.图表类型
            chart.DefaultSeries.Type = SeriesType.Column;// SeriesType.Column;
            //chart.DefaultSeries.Type = SeriesType.Cylinder;
            //2.图表类型
            //柱状图
            //chart.Type = ChartType.TreeMap;
            ////横向柱状图
            if (ishorizontal)
                chart.Type = ChartType.ComboHorizontal;
            else
                chart.Type = ChartType.Combo;// ChartType.ComboHorizontal
            ////横向柱状图
            //chart.Type =_chartType;// ChartType.Gantt;
            ////饼状图
            //chart.Type = ChartType.Pies;

            //y轴图表阴影颜色
            //chart.YAxis.AlternateGridBackground.Color = Color.FromArgb(255, 250, 250, 250);

            //chart.LegendBox.HeaderLabel = new dotnetCHARTING.WinForms.Label("图表说明", new Font("Microsoft Sans Serif", 10F, FontStyle.Bold, GraphicsUnit.Point, 134));
            ////chart.LegendBox.HeaderLabel.Font = new Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 134);
            //chart.LegendBox.Label.Font = new Font("Times New Roman", 7F, FontStyle.Bold, GraphicsUnit.Point, 134);
            //chart.Palette = new Color[] { Color.FromArgb(0, 156, 255), Color.FromArgb(255, 99, 49), Color.FromArgb(49, 255, 49), Color.FromArgb(255, 255, 0), };

            chart.Title = title;
            //X轴柱状图值字体
            chart.XAxis.Label.Text = "";
            chart.XAxis.Label.Font = new Font("Times New Roman", 7F, FontStyle.Bold, GraphicsUnit.Point, 134);
            //设置X轴刻度值说明字体
            chart.XAxis.DefaultTick.Label.Font = new Font("Times New Roman", 7F, FontStyle.Bold, GraphicsUnit.Point, 134); ;
            //if(ishorizontal) 
            //    chart.YAxis.StaticColumnWidth = columnwidth; //每个单元格的宽度


            //Y轴柱状图值字体
            chart.YAxis.Label.Text = "";
            chart.YAxis.Label.Font = new Font("Times New Roman", 7F, FontStyle.Bold, GraphicsUnit.Point, 134); ;
            //设置Y轴刻度值说明字体
            chart.YAxis.DefaultTick.Label.Font = new Font("Times New Roman", 7F, FontStyle.Bold, GraphicsUnit.Point, 134);

            if (ishorizontal)
            {
                chart.XAxis.Minimum = 0;
                chart.XAxis.MinimumInterval = 25;
                chart.XAxis.Maximum = 100;
                chart.XAxis.Percent = true;
            }
            else
            {
                chart.YAxis.Minimum = 0;
                chart.YAxis.MinimumInterval = 25;
                chart.YAxis.Maximum = 100;
                chart.YAxis.Percent = true;
            }
            

            //Y轴箭头标示
            chart.XAxis.Name = "";
            if (chart.Type == ChartType.ComboHorizontal)
            {
                chart.XAxis.TickLabelPadding = 10;
                //chart.XAxis.Line.StartCap = System.Drawing.Drawing2D.LineCap.Square;
                //chart.XAxis.Line.EndCap = System.Drawing.Drawing2D.LineCap.ArrowAnchor;
                //chart.XAxis.Line.Width = 5;//箭头宽度
                //chart.XAxis.Line.Color = Color.Gray;
                chart.XAxis.NumberPercision = 1;
            }
            else
            {
                chart.YAxis.TickLabelPadding = 10;
                //chart.YAxis.Line.StartCap = System.Drawing.Drawing2D.LineCap.Square;
                //chart.YAxis.Line.EndCap = System.Drawing.Drawing2D.LineCap.ArrowAnchor;
                //chart.YAxis.Line.Width = 5;//箭头宽度
                //chart.YAxis.Line.Color = Color.Gray;
                //显示值格式化(小数点显示几位)
                chart.YAxis.NumberPercision = 1;
            }

            //图片存放路径
            //chart.TempDirectory = System.Environment.CurrentDirectory + "\\";
            //图表宽度
            chart.Width = weight;
            //图表高度
            chart.Height = height;
            chart.Series.Name = "";
            //单一图形
            //chart.Series.Data = _dt;
            //chart.SeriesCollection.Add();

            //图例在标题行显示，但是没有合计信息
            //chart.TitleBox.Position = TitleBoxPosition.FullWithLegend;
            //chart.TitleBox.Label.Alignment = StringAlignment.Center;
            //chart.LegendBox.Position = LegendBoxPosition.None; //不显示图例,指不在右侧显示，对上面一行的属性设置并没有影响

            chart.LegendBox.Visible = false;
            //chart.LegendBox.HeaderLabel.Text = "";
            //chart.LegendBox.LabelStyle.Text = "";
            //chart.LegendBox.Template = "%Icon%Name";
            
            chart.DefaultSeries.DefaultElement.ShowValue = true;
            chart.ShadingEffect = true;
            chart.Use3D = false;
            chart.Series.DefaultElement.ShowValue = true;
            chart.DefaultElement.SmartLabel.Text = "%Value%";
            chart.DefaultElement.SmartLabel.Color = Color.DarkBlue;
            

            if (ishorizontal)
            {
                chart.DefaultElement.SmartLabel.LineAlignment = StringAlignment.Center;
                chart.DefaultElement.SmartLabel.Alignment = LabelAlignment.Automatic;
                //chart.DefaultElement.SmartLabel.Font = new Font("Times New Roman", 5F, FontStyle.Bold, GraphicsUnit.Point, 134);
                //chart.DefaultElement.SmartLabel.ForceVertical = true;
            }
            else
            {
                //chart.DefaultElement.SmartLabel.Alignment = LabelAlignment.Automatic;
                //chart.DefaultElement.SmartLabel.
                chart.DefaultElement.SmartLabel.Alignment = LabelAlignment.OutsideTop;
                //chart.DefaultElement.SmartLabel.ForceVertical = false;
                
            }

           
            chart.ImageFormat = ImageFormat.Emf;
            SeriesCollection sc = new SeriesCollection();
            chart.SeriesCollection.Add(GetArrayData(dt, sc, "百分位 Percentile", color[0], iszf));
            //Bitmap sourceBitmap = new Bitmap(chart.Width, chart.Height);
            //chart.DrawToBitmap(sourceBitmap, new Rectangle(0, 0, chart.Width, chart.Height));
            
            //Thread.CurrentThread.SetApartmentState(ApartmentState.STA);
            //Utils.mutex_clipboard.WaitOne();
            //Clipboard.Clear();
            //Clipboard.SetImage(sourceBitmap);

            return chart.FileManager.SaveImage();
        }
        public static void CreateColumn(DataTable dt)
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
                chart.ShadingEffectMode = ShadingEffectMode.Eight;

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
                chart.Palette = new Color[] {  Color.FromArgb(0, 156, 255), Color.FromArgb(255, 99, 49), Color.FromArgb(49, 255, 49), Color.FromArgb(255, 255, 0), };

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
                chart.YAxis.DefaultTick.Label.Font = new Font("Microsoft Sans Serif", 9F, FontStyle.Bold, GraphicsUnit.Point, 134);

                chart.YAxis.Minimum = 0.4;
                chart.YAxis.MinimumInterval = 0.05;

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
                SeriesCollection sc = new SeriesCollection();
                chart.SeriesCollection.Add(GetArrayData(dt, sc, "", color[0], false));
                Bitmap sourceBitmap = new Bitmap(chart.Width, chart.Height);
                chart.DrawToBitmap(sourceBitmap, new Rectangle(0, 0, chart.Width, chart.Height));
                //Thread.CurrentThread.SetApartmentState(ApartmentState.STA);
                Utils.mutex_clipboard.WaitOne();
                Clipboard.Clear();
                Clipboard.SetImage(sourceBitmap);

            
        }

        private static SeriesCollection GetArrayData(DataTable dt, SeriesCollection sc, string name, Color my_color, bool iszf)
        {
            
            try
            { 
                Series s = new Series();
                
                s.Name = name;
                s.Element.Color = my_color;
                
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    
                    var e = new Element();

                    // 每元素的名称
                    e.Name = dt.Rows[i][0].ToString();
                    //设置柱状图值的字体
                    e.SmartLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 134);
                    e.SmartLabel.DynamicPosition = true;
                    //e.SmartLabel.DynamicDisplay = true;
                    //e.SmartLabel.AutoWrap = true;
                    // 每元素的大小数值
            
                    e.YValue = Convert.ToDouble(dt.Rows[i][1].ToString());
                    e.Color = my_color;
                    if ((decimal)dt.Rows[i][1] == 0 && iszf)
                        e.SmartLabel.Text = "";        
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

            chart.YAxis.Minimum = 0.4;
            chart.YAxis.MinimumInterval = 0.05;
            //chart.YAxis.MinorInterval = 0.05;
            
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
            chart.Width = 590;
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
            chart.LegendBox.Position = LegendBoxPosition.BottomMiddle;
            chart.LegendBox.HeaderLabel.Text = "";
            chart.LegendBox.LabelStyle.Text = "";
            chart.LegendBox.Template = "%Icon%Name";
            chart.DefaultSeries.DefaultElement.ShowValue = true;
            chart.ShadingEffect = true;
            chart.Use3D = true;
            //chart.Series.DefaultElement.ShowValue = true;
            int count = 0;
            SeriesCollection sc = new SeriesCollection();
            foreach (var kv in dts)
            {
                GetArrayData(kv.Value, sc, kv.Key, color[count % 5], false);
                
                count++;
            }
            chart.SeriesCollection.Add(sc);
            Bitmap sourceBitmap = new Bitmap(chart.Width, chart.Height);
            chart.DrawToBitmap(sourceBitmap, new Rectangle(0, 0, chart.Width, chart.Height));
            //Thread.CurrentThread.SetApartmentState(ApartmentState.STA);
            Utils.mutex_clipboard.WaitOne();
            Clipboard.Clear();
            Clipboard.SetImage(sourceBitmap);


        }
 
    }
}
