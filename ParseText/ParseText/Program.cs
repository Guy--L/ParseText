using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Drawing;
using System.Windows.Forms;
using ClosedXML.Excel;
using System.Configuration;
using mn = MathNet.Numerics;
using MathNet.Numerics.Statistics;
using System.Windows.Forms.DataVisualization.Charting;
using Microsoft.SolverFoundation.Services;
using System.Diagnostics;

namespace ParseText
{
    class Program
    {
        /// <summary>
        /// Various types of tests in this Rheology lab
        /// </summary>
        enum TestType
        {
            Trim,
            Lather,
            Cohesion,
            Fract_Band,
            Oscillation,
            Error = -1
        };

        /// <summary>
        /// Number of rows in text files that correspond to the test types above
        /// </summary>
        private static List<int> rowmap = new List<int>() { 1, 144, 200, 418, 38, 0 };
        private static Dictionary<int, TestType> testmap = rowmap.Select((v, i) => new { value = v, index = i }).ToDictionary(v => v.value, k => (TestType)k.index);

        const int firstline = 9;
        private static string _data;
        private static string _infileprefix = @"Rheology Form Filled In ";
        private static string _outfilename = @"{0} Rheology Analysis v3 with SPTT Entry Macro (MACRO v4.1) {1}";
        private static string _outdirectory;
        private static string _currentsample;

        private static SolverContext context;
        private static Model model = null;
        private static Decision N0, TC;

        private static Dictionary<string, double[]> _t95man = new Dictionary<string, double[]>();
        private static List<double>[] _t95err = new List<double>[5]
        {
            new List<double>(),
            new List<double>(),
            new List<double>(),
            new List<double>(),
            new List<double>()
        };

        private static Form1 form;

        [STAThread]
        static void Main(string[] args)
        {
            _infileprefix = Properties.Settings.Default["infileprefix"].ToString();
            _outfilename = Properties.Settings.Default["outfilename"].ToString();
            _outdirectory = ConfigurationManager.AppSettings["outdirectory"];

            if (args.Length == 0) {
                args = new string[] { "." };
            }

            // look for request XLs in all directories on command line
            testmap.Add(143, TestType.Lather);
            testmap.Add(142, TestType.Lather);
            testmap.Add(417, TestType.Fract_Band);

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            form = new Form1();
            Application.Run(form);
        }

        /// <summary>
        /// Read each request XL in directory
        /// </summary>
        /// <param name="MyDir">Directory to iterate over</param>
        /// 
        public static void ControlXLInDir(string MyDir)
        {
            _data = MyDir;
            _t95man.Clear();

            var docs = Directory.GetFiles(MyDir, "*.xlsm");

            form.WriteLine("Test\tFilename");
            foreach (var d in docs)
            {
                ReadControlXL(d);
            }
        }

        /// <summary>
        /// After export using DataAnalysis, the letters near the end of the filename represent the can
        /// </summary>
        /// <param name="filename"></param>
        /// <returns>Letter code representing can</returns>
        private static string can(string filename)
        {
            var fix = filename.Split(' ');
            var tst = fix[fix.Length - 2];
            if (tst.Contains('('))
                return fix[fix.Length - 3];
            return tst.Split('-')[0];
        }

        /// <summary>
        /// Create a way to map from the can name to the list of files containing test results for that can
        /// </summary>
        /// <param name="dir"></param>
        /// <param name="sample"></param>
        /// <returns></returns>
        static ILookup<string, string> LoadSampleFiles(string dir, string sample)
        {
            var filemask = sample + "*.txt";
            var samplefiles = Directory.GetFiles(dir, filemask);

            if (samplefiles.Count() == 0)
            {
                filemask = filemask.Insert(_currentsample.Length, "-");
                samplefiles = Directory.GetFiles(dir, filemask);
            }
            //Console.WriteLine(samplefiles.Count() + " samples in " + dir + " " + filemask);
            return samplefiles.ToLookup(k => Program.can(k), v => v);
        }

        private static IXLWorksheet _outsh;
        private static string _can;

        static void ReadManualXL(string xlfile)
        {
            var maxl = new XLWorkbook(xlfile);
            var inxl = new XLWorkbook(xlfile);
            IXLWorksheet insh;
            if (!inxl.TryGetWorksheet("Summary Table", out insh))
                return;

            foreach (var row in insh.Rows())
            {
                if (row.RowNumber() < 4)
                    continue;

                var can = row.Cell(1).GetValue<string>();

                if (string.IsNullOrWhiteSpace(can))
                    continue;

                _t95man[can] = new double[] {
                      row.Cell(6).GetValue<double>()
                    , row.Cell(7).GetValue<double>()
                    , row.Cell(8).GetValue<double>()
                    , row.Cell(9).GetValue<double>()
                    , row.Cell(10).GetValue<double>()
                };
            }
        }

        private static string[] nLabels = new string[] { "n^2", "n0", "nInf", "tC", "t95" };
        private static int[] bins = new int[] { 200, 200, 200, 10000, 10000 };
        private static int[] show = new int[] { 20, 20, 20, 20, 20 }; 

        public static void ChartHistograms()
        {
            Dictionary<string, Tuple<Histogram, DescriptiveStatistics>> series = new Dictionary<string, Tuple<Histogram, DescriptiveStatistics>>();
            string[] dsc = new string[nLabels.Count()];
            int j = 0;
            foreach (var stat in nLabels)
            {
                var err = _t95err[j];
                var hist = new Histogram(err, bins[j]);
                var stats = new DescriptiveStatistics(err);
                form.WriteLine(stat + " % count above 5%: " + (err.Count(e => e > 0.05) * 100.0 / err.Count()).ToString("N2"));
                Debug.WriteLine(stat + " % count above 5%: " + (err.Count(e => e > 0.05) * 100.0 / err.Count()).ToString("N2"));
                dsc[j] = stat + ": mean " + stats.Mean.ToString("e3") + ", std " + stats.StandardDeviation.ToString("e3") + ", min " + stats.Minimum.ToString("e3") + ", max " + stats.Maximum.ToString("e3");
                series[stat+" (n = "+hist.DataCount+")"] = new Tuple<Histogram, DescriptiveStatistics>(hist, stats);
                j++;
            }

            foreach(var sc in dsc)
            {
                form.WriteLine(sc);
                Debug.WriteLine(sc);
            }

            ChartCounts(series);
        }

        static void ReadControlXL(string xlfile)
        {
            var inxl = new XLWorkbook(xlfile);
            IXLWorksheet insh;
            if (!inxl.TryGetWorksheet("Sample ID Sort", out insh))
                return;
            var request = Path.GetFileNameWithoutExtension(xlfile).Substring(_infileprefix.Length).Split(' ');
            var data = Path.Combine(_data, request[2]);

            if (!Directory.Exists(data))
            {
                form.WriteLine(data + " folder does not exist");
                return;
            }
            form.WriteLine(data + " folder exists");

            if (form.doCharts)
            {
                var manxl = Directory.GetFiles(data, "*.xlsm").FirstOrDefault(f => f.Contains("Manual"));
                if (manxl != null) ReadManualXL(manxl);
            }

            _currentsample = request[2];

            var outfilename = string.Format(_outfilename, string.Join(" ", request.Take(2)), request[2]) + ".xlsm";
            var outpath = form.notoutset ? data : form.outdir;
            var outfile = Path.Combine(outpath, outfilename);

            //form.WriteLine("writing to " + outfile);

            var outxl = new XLWorkbook("AnalysisTemplate.xlsm");
            _outsh = outxl.Worksheet("Summary Table");
            string samplename = "";
            ILookup<string, string> samples = null;

            foreach (var row in insh.Rows())
            {
                if (row.RowNumber() < 2)
                {
                    _outsh.FirstRow().Cell(1).Value = "Request " + request[1] + " - ";
                    _outsh.FirstRow().Cell(3).Value = request[2];
                    continue;
                }
                string sample = row.Cell(4).GetValue<string>();
                var outrowi = row.RowNumber() + 2;
                var outrow = _outsh.Row(outrowi);
                if (samplename != sample)
                {
                    samples = LoadSampleFiles(data, sample);
                    samplename = sample;
                    string average = "=AVERAGE(L" + outrowi + ":L" + (outrowi + 3) + ")";
                    outrow.Cell(14).SetFormulaA1(average);

                    average = "=AVERAGE(J" + outrowi + ":J" + (outrowi + 3) + ")";
                    outrow.Cell(11).SetFormulaA1(average);
                }

                var can = row.Cell(1).GetValue<string>();
                _can = can;
                outrow.Cell(1).SetValue<string>(can);
                outrow.Cell(2).SetValue<string>(row.Cell(2).GetValue<string>().Split(' ')[0]);
                outrow.Cell(3).SetValue<string>(row.Cell(3).GetValue<string>());
                outrow.Cell(4).SetValue<string>(row.Cell(4).GetValue<string>());
                outrow.Cell(5).SetValue<string>(row.Cell(5).GetValue<string>());

                foreach (var file in samples[can])
                {
                    ReadFile(file, outrow);
                }
            }

            outxl.SaveAs(outfile);
            //form.WriteLine("saved as " + outfilename);
        }

        private static int initialskip = 4;
        private static int initialtake = 5;
        private static int finalskip = 32;
        private static int finaltake = 3;
        private static double t95 = -Math.Log(.05);

        private static List<string> colors = new List<string>()
        {
            "Blue", "Red", "Green", "Black", "Orange", "Pink", "Purple", "Brown", "Yellow"
        };

        static void ChartCounts(Dictionary<string, Tuple<Histogram, DescriptiveStatistics>> series)
        {
            int j = 0;
            foreach (var line in series.Keys)
            {
                var c = new Chart() { Size = new Size(1920, 1080) };
                c.Titles.Add("Count vs Error for " + line);
                c.Titles[0].Font = new Font("Arial", 14, FontStyle.Bold);

                var a = new ChartArea("Lather");
                a.AxisY.MajorGrid.LineColor = Color.LightGray;
                a.AxisY.LabelStyle.Font = new Font("Arial", 14);
                a.AxisY.Title = "Count";
                a.AxisY.TitleFont = new Font("Arial", 14);
                a.AxisX.Title = "Error = |manual-fit|/manual";
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
                var t = new Series(line)
                {
                    ChartType = SeriesChartType.Column,
                    XValueType = ChartValueType.Double,
                    YValueType = ChartValueType.Double,
                    Color = Color.FromName(colors[n++]),
                    Legend = "Legend"
                };
                var hist = series[line].Item1;
                var stats = series[line].Item2;
                for (int i=0; i<show[j]; i++)
                {
                    t.Points.AddXY(hist[i].UpperBound, hist[i].Count);
                }
                TextAnnotation ta = new TextAnnotation();
                ta.Text = stats.Mean + " mean\n" + stats.StandardDeviation + " stdev\n";
                c.Annotations.Add(ta);
                c.Series.Add(t);
                t.ChartArea = "Lather";

                var chartpath = form.notoutset ? _data : form.outdir;
                var filename = Path.Combine(chartpath, line.Split('(')[0]);
                //form.WriteLine("saving chart " + filename);
                c.SaveImage(filename + ".png", ChartImageFormat.Png);
                j++;
            }
        }

        static void ChartSeries(string name, Dictionary<string, List<Reading>> series)
        {
            var c = new Chart() { Size = new Size(1920, 1080) };
            c.Titles.Add("Normal vs Time for " + name);
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

            c.Legends.Add(new Legend("Legend") {
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
                series[line].Select(r => t.Points.AddXY(r.time, r.normal)).ToList();
                c.Series.Add(t);
                t.ChartArea = "Lather";
            }
            var chartpath = form.notoutset ? _data : form.outdir;
            var filename = Path.Combine(chartpath, _currentsample + "_" + name.Split('(')[0]);
            //form.WriteLine("saving chart " + filename);
            c.SaveImage(filename + ".png", ChartImageFormat.Png);
        }

        //private static List<string> Issue3 = new List<string> {
        //    "G-0033f",
        //    "L-0058f"
        //};

        static void ReadFile(string file, IXLRow outrow)
        {
            var lines = File.ReadAllLines(file);
            var datalines = lines.Count() - firstline - 1;
            TestType testType = testmap[datalines];

            if (testType != TestType.Lather)
                return;

            var f = Path.GetFileNameWithoutExtension(file);
            form.WriteLine(testType.ToString() + "\t" + can(file));

            if (testType == TestType.Cohesion)
            {
                var pairs = lines.Skip(firstline).Take(rowmap[(int)TestType.Cohesion]).Select(s =>
                {
                    var a = s.Split('\t');
                    return new { time = double.Parse(a[0]), normal = double.Parse(a[2]) };
                }).ToList();
                var min = pairs.First(b => b.normal == pairs.Min(a => a.normal));
                outrow.Cell(12).SetValue<double>(min.normal);
                outrow.Cell(13).SetValue<double>(min.time);
            }
            if (testType == TestType.Lather)
            {
                var setup = lines.Skip(firstline).Take(rowmap[(int)TestType.Lather]).Select(s => new Reading(s));
                var data = setup.Where(d => d.rate >= 99.0 && d.rate <= 101.0).ToList();

                var max = data.Max(d => d.normal);

                var ninf = setup.Where(d => d.time > 20).Average(d => d.normal);
                var mid = (max + ninf) / 2;

                var pair = data.Select((v, i) => new { val = v.cutoff(mid), idx = i });
                max = pair.Max(d => d.val.normal);
                var maxidx = pair.First(d => d.val.normal == max).idx;

                var minip = pair.FirstOrDefault(d => d.idx > maxidx && d.val.normal <= mid);
                if (minip == null)
                {
                    form.WriteLine("no min index for " + _currentsample + " " + can(file));
                    return;
                }
                var mini = minip.idx;

                var maxip = pair.Skip(mini).FirstOrDefault(d => d.val.normal < ninf);
                if (maxip == null)
                {
                    form.WriteLine("no max index for " + _currentsample + " " + can(file));
                    return;
                }
                var maxi = maxip.idx - 1;

                while (maxi <= mini + 1)
                {
                    //form.Write(".");
                    maxi++;
                }

                //                form.WriteLine("max: " + max + ", inf: "+ ninf + ", avg: "+ mid);
                var n2fit = data.Skip(mini).Take(maxi - mini);

                var y2fit = n2fit.Select(d => Math.Log(Math.Abs(d.normal - ninf))).ToArray();
                var x2fit = n2fit.Select(d => d.time).ToArray();
                Tuple<double, double> p = mn.Fit.Line(x2fit, y2fit);   // item1 intercept, item2 slope
                var n0 = Math.Exp(p.Item1) + ninf;

                // Create the model
                if (model == null)
                {
                    context = SolverContext.GetContext();
                    model = context.CreateModel();
                    // Add a decisions
                    N0 = new Decision(Domain.Real, "n0");
                    TC = new Decision(Domain.Real, "tc");
                    model.AddDecisions(N0);
                    model.AddDecisions(TC);
                }
                var mind = data.Skip(mini).First();
                var maxd = data.Skip(maxi + 1).First();
                var note = false;
                var tc = -1 / (p.Item2 == 0 ? 1 : p.Item2);
                if (mind.time + 5 > maxd.time || mind.normal < maxd.normal || n0 > 10.0)
                {
                    n0 = 2.0;
                    note = true;
                }

                N0.SetInitialValue(n0);
                TC.SetInitialValue(tc);

                n2fit = data.Skip(mini).Where(t => t.normal > 0);
                var cost = new SumTermBuilder(n2fit.Count());

                n2fit.ForEach(d =>
                {
                    Term r = N0 + (ninf - N0) * (1 - Model.Exp(-d.time / TC));
                    r -= d.normal;
                    r *= r;
                    cost.Add(r);
                });

                model.AddGoal("Chi2", GoalKind.Minimize, cost.ToTerm());            // add goal

                //var directive = new CompactQuasiNewtonDirective();
                //var solver = context.Solve(directive);
                var solver = context.Solve();
                var report = solver.GetReport();
                if (note)
                {
                    var rpt = report.ToString();
                    var lns = rpt.Split('\n');
                    foreach (var ln in lns)
                        form.WriteLine(ln);
                }
                var chi2 = model.Goals.First().ToDouble();

                outrow.Cell(6).SetValue<double>(chi2);
                outrow.Cell(7).SetValue<double>(N0.GetDouble());
                outrow.Cell(8).SetValue<double>(ninf);
                outrow.Cell(9).SetValue<double>(TC.GetDouble());
                outrow.Cell(10).SetValue<double>(TC.GetDouble() * t95);

                var addedzero = (new List<Reading>() { new Reading(0, N0.GetDouble()) }).Concat(setup);

                //form.WriteLine("--> Chi2: " + chi2.ToString("F") + ", N0: " + N0.GetDouble().ToString("F") + ", TC: " + TC.GetDouble().ToString("F") + ", ninf: " + ninf.ToString("F"));
                model.RemoveGoal(model.Goals.First());                              // remove goal for next model run       

                if (form.doCharts)
                {
                    var xlv = _t95man[can(file)];
                    var xlf = addedzero.Select(d => new Reading(d.time, xlv[1] + (xlv[2] - xlv[1]) * (1 - Math.Exp(-d.time / xlv[3]))));
                    var fit = addedzero.Select(d => new Reading(d.time, N0.GetDouble() + (ninf - N0.GetDouble()) * (1 - Math.Exp(-d.time / TC.GetDouble()))));
                    var t95fit = new double[] { chi2, N0.GetDouble(), ninf, TC.GetDouble(), (TC.GetDouble() * t95) };
                    t95fit.Select((t, i) =>
                    {
                        var m = _t95man[can(file)][i];
                        var e = Math.Abs(m - t) / (m == 0 ? 1 : m);
                        if (e > 100000)
                        {
                            Debug.WriteLine("error: " + e.ToString("e5") + ", file: " + file + ", can: " + can(file) + ", i: " + i);
                        }
                        else if (!note)
                            _t95err[i].Add(e);
                        return 1;
                    }).ToList();

                    Dictionary<string, List<Reading>> Series = new Dictionary<string, List<Reading>>();
                    Series["readings"] = setup.ToList();
                    if (N0.GetDouble() < 1000.0) Series["fit"] = fit.ToList();
                    Series["xl"] = xlf.ToList();
                    Series["zero"] = new List<Reading>() { new Reading(data.Min(t => t.time), 0), new Reading(data.Max(t => t.time), 0) };
                    Series["midnormal"] = new List<Reading>() { new Reading(mind.time, 1), new Reading(mind.time, -1) };
                    Series["subplateau"] = new List<Reading>() { new Reading(maxd.time, 1), new Reading(maxd.time, -1) };
                    //Series["fitted"] = n2fit.ToList();
                    //var begin = n2fit.First();
                    //var nonzero = n2fit.First(t => t.normal > 0);
                    //Series["fitted"] = new List<Reading>() {
                    //    new Reading(begin.time, begin.normal),
                    //    new Reading(nonzero.time, nonzero.normal),
                    //    new Reading(n2fit.Last().time, nonzero.normal)
                    //};

                    var title = can(file) + " (chi2 = " + chi2.ToString("e3") + ")" + note;
                    ChartSeries(title, Series);
                }
            }
            if (testType == TestType.Oscillation)
            {
                var readings = lines.Skip(firstline).Select(s => new Reading(s));
                var initial = readings.Skip(initialskip).Take(initialtake);
                var final = readings.Skip(finalskip).Take(finaltake);

                var iline = new Line(initial);
                var fline = new Line(final);

                var intersectx = (iline.intercept - fline.intercept) / (fline.slope - iline.slope);
                //Console.WriteLine("intersectx: " + intersectx);

                var mid = readings.TakeWhile(s => s.strain < intersectx);
                var midtriple = readings.Skip(mid.Count() - 2).Take(3);
                var mline = new Line(midtriple);

                var ypstrain = (iline.intercept - mline.intercept) / (mline.slope - iline.slope);
                var ypt = readings.TakeWhile(s => s.strain < ypstrain);
                var ypi = readings.Skip(ypt.Count() - 1).Take(2);
                var yp = ypi.ToArray();

                var ypstress = (yp[1].stress - yp[0].stress) / (yp[1].strain - yp[0].strain) * (ypstrain - yp[0].strain) + yp[0].stress;
                var bpstrain = (fline.intercept - mline.intercept) / (mline.slope - fline.slope);
                var bpt = readings.TakeWhile(s => s.strain < bpstrain);
                var bp = readings.Skip(bpt.Count() - 1).Take(2).ToArray();
                var bpstress = (bp[1].stress - bp[0].stress) / (bp[1].strain - bp[0].strain) * (bpstrain - bp[0].strain) + bp[0].stress;

                var cross = readings.Skip(1).TakeWhile(s => s.prime >= s.dprime);
                var crx = readings.Skip(cross.Count()).Take(2);
                var cr = crx.ToArray();

                var pline = new Line(cr, c => c.prime);
                var dline = new Line(cr, c => c.dprime);

                var strd = cr[1].strain - cr[0].strain;
                var denom = ((cr[1].dprime - cr[0].dprime) / strd - (cr[1].prime - cr[0].prime) / strd);

                var gstrain = (pline.intercept - dline.intercept) / denom;
                var gstress = (cr[1].prime - cr[0].prime) / strd * (gstrain - cr[1].strain) + cr[1].prime;

                var gplateau = initial.Average(g => g.prime);
                var dplateau = initial.Average(g => g.dprime);

                outrow.Cell(15).SetValue(ypstress);
                outrow.Cell(16).SetValue(ypstrain);
                outrow.Cell(17).SetValue(bpstress);
                outrow.Cell(18).SetValue(bpstrain);

                outrow.Cell(19).SetValue(gstress);
                outrow.Cell(20).SetValue(gstrain);
                outrow.Cell(21).SetValue(gplateau);
                outrow.Cell(22).SetValue(dplateau);
            }
            if (testType == TestType.Fract_Band)
            {
                var orig = lines.Skip(firstline).Take(rowmap[(int)TestType.Fract_Band]).Select(s => new Reading(s)).ToList();
                var data = orig.Select(d => (d.rate >= 0.95 && d.rate <= 1.05)? d : new Reading(d));

                var max = data.Take(319).Max(s => s.shear);
                var sam = data.Take(319).TakeWhile(s => s.shear < max);

                var start = sam.Any()?sam.Count(): 319;
                var avg = data.Skip(start).Take(101).Average(s => s.shear);
                var category = avg < max;

                var at5 = data.Last(t => t.time <= 5.0).shear;
                 
                var span = data.Max(t => t.time);
                var at60 = span >= 60.0 ? data.First(t => t.time >= 60.0).shear : data.Last().shear;
                var delta = at60 - at5;

                outrow.Cell(23).SetValue<int>(category?1:0);
                outrow.Cell(24).SetValue<double>(at5);
                outrow.Cell(25).SetValue<double>(at60);
                outrow.Cell(26).SetValue<double>(delta);
                if (category) outrow.Cell(27).SetValue<double>(orig.Max(s => s.shear));
            }
        }
    }
}
