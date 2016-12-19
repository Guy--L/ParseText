using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using MathNet.Numerics.Statistics;
using System.Windows.Forms.DataVisualization.Charting;
using System.Drawing;

namespace ParseText
{
    /// <summary>
    /// Separate file to avoid name collisions with Excel Interop
    /// </summary>
    partial class Program
    {
        private static List<string> colors = new List<string>()
        {
            "Blue", "Red", "Green", "Black", "Orange", "Pink", "Purple", "Brown", "Yellow"
        };

        static void ChartSeries(string name, Dictionary<string, List<Reading>> series)
        {
            if (series == null)
                return;

            var c = new Chart() { Size = new Size(1920, 1080) };
            c.Titles.Add("Normal vs Time for " + _currentsample + " " + name);
            c.Titles[0].Font = new Font("Arial", 14, FontStyle.Bold);

            var a = new ChartArea("Lather");
            a.AxisY.MajorGrid.LineColor = Color.LightGray;
            a.AxisY.LabelStyle.Font = new Font("Arial", 14);
            a.AxisY.Title = "Normal";
            a.AxisY.TitleFont = new Font("Arial", 14);
            a.AxisX.Title = "Time";
            a.AxisX.TitleFont = new Font("Arial", 14);
            a.AxisX.IsStartedFromZero = true;
            a.AxisY.IsStartedFromZero = true;
            a.AxisX.IsMarginVisible = false;
            a.AxisX.MajorGrid.LineColor = Color.LightGray;
            a.AxisX.LabelStyle.ForeColor = Color.Black;
            a.AxisX.LabelStyle.Font = new Font("Arial", 14);
            a.AxisX.IsLabelAutoFit = true;
            a.AxisX.Minimum = 0;
            c.ChartAreas.Add(a);

            c.Legends.Add(new Legend("Legend")
            {
                IsDockedInsideChartArea = true,
                DockedToChartArea = "Lather"
            });

            var n = 0;
            foreach (var line in series.Keys)
            {
                var t = new Series(line)
                {
                    ChartType = SeriesChartType.Line,
                    XValueType = ChartValueType.Double,
                    YValueType = ChartValueType.Double,
                    Color = Color.FromName(colors[n++]),
                    Legend = "Legend"
                };
                series[line].Select(r => r!=null?t.Points.AddXY(r.time, r.normal):0).ToList();
                c.Series.Add(t);
                t.ChartArea = "Lather";
            }
            string chartpath = _data;
            string outname = _currentsample + "_" + name.Split('(')[0];
            if (!form.notoutset) {
                chartpath = Path.Combine(form.outdir, _currentsample + " Graphs");
                Directory.CreateDirectory(chartpath);
            }
            var filename = Path.Combine(chartpath, outname);
            //form.WriteLine("saving chart " + filename);
            c.SaveImage(filename + ".png", ChartImageFormat.Png);
        }
    }
}
