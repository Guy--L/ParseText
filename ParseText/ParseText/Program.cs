using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Drawing;
using ClosedXML.Excel;
using System.Configuration;
using mn = MathNet.Numerics;
//using Microsoft.SolverFoundation.Services;
using System.Windows.Forms.DataVisualization.Charting;

namespace ParseText
{
    class Line
    {
        private double[] x;
        private double[] y;
        
        private IEnumerable<Reading> _data;
        public double intercept;
        public double slope;

        public Line(IEnumerable<Reading> r) : this(r, c => c.stress)
        {}

        public Line(IEnumerable<Reading> r, Func<Reading, double> column)
        {
            _data = r;
            //_data.ForEach(d => Console.WriteLine(column(d) + ", " + d.strain));
            Fit(column);
            //Console.WriteLine("y=" + slope + "x+" + intercept);
        }

        public double Fit(Func<Reading, double> column)
        {
            //y = _data.Select(s => Math.Round(column(s), 3)).ToArray();
            //x = _data.Select(s => Math.Round(s.strain, 3)).ToArray();
            y = _data.Select(s => column(s)).ToArray();
            x = _data.Select(s => s.strain).ToArray();
            Tuple<double, double> z = mn.Fit.Line(x, y);

            intercept = z.Item1;
            slope = z.Item2;
            return 0.0;
        }
    }

    class Reading
    {
        public double stress;
        public double strain;
        public double prime;
        public double dprime;

        // overload reading members to provide two facades
        public double time
        {
            get { return stress; }
            set { stress = value; }
        }
        public double rate
        {
            get { return strain; }
            set { stress = value; }
        }
        public double normal
        {
            get { return prime; }
            set { prime = value; }
        }
        public double shear
        {
            get { return dprime; }
            set { dprime = value; }
        }
        public Reading(string val)
        {
            if (string.IsNullOrWhiteSpace(val))
                return;

            var a = val.Split('\t');
            stress = double.Parse(a[0]);
            strain = double.Parse(a[1]);
            prime = double.Parse(a[2]);
            dprime = double.Parse(a[3]);
        }
        public Reading(double t, double n)
        {
            time = t;
            normal = n;
        }
        public Reading(Reading toZero)
        {
            rate = 0.0;
            shear = 0.0;
            time = toZero.time;
            normal = toZero.normal;
        }
        public void print()
        {
            Console.WriteLine("(" + stress + ", " + strain + ", " + prime + ", " + dprime + ")");
        }
    }

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
        private static string _currentsample;

        private static Dictionary<string, double[]> _t95 = new Dictionary<string, double[]>();

        //private static SolverContext context;
        //private static Model model = null;
        //private static Decision N0, TC;

        static void Main(string[] args)
        {
            _data = ConfigurationManager.AppSettings["datadirectory"];
            _infileprefix = ConfigurationManager.AppSettings["infileprefix"];
            _outfilename = ConfigurationManager.AppSettings["outfilename"];

            if (args.Length == 0) {
                args = new string[] { "." };
            }

            // look for request XLs in all directories on command line
            testmap.Add(143, TestType.Lather);
            testmap.Add(142, TestType.Lather);

            foreach (var s in args)
            {
                Console.WriteLine(s);
                ControlXLInDir(s);
            }

            Console.WriteLine("done, hit key to close");
            Console.ReadKey();
        }

        /// <summary>
        /// Read each request XL in directory
        /// </summary>
        /// <param name="MyDir">Directory to iterate over</param>
        /// 
        public static void ControlXLInDir(string MyDir)
        {
            var docs = Directory.GetFiles(MyDir, "*.xlsm");

            Console.WriteLine("Test\tFilename");
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

                _t95[can] = new double[] { row.Cell(7).GetValue<double>(), row.Cell(8).GetValue<double>(), row.Cell(9).GetValue<double>() };
            }
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
                Console.WriteLine(data + " folder does not exist");
                return;
            }
            Console.WriteLine(data + " folder exists");
            _currentsample = request[2];

            var outfilename = string.Format(_outfilename, string.Join(" ", request.Take(2)), request[2]) + ".xlsm";
            
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
            
            outxl.SaveAs(Path.Combine(data, outfilename));
            Console.WriteLine("saved as " + outfilename);
        }

        private static int initialskip = 4;
        private static int initialtake = 5;
        private static int finalskip = 32;
        private static int finaltake = 3;
        private static double t95 = -Math.Log(.05);

        static void ChartFile(string name, List<Reading> data, List<Reading> fit, List<Reading> man)
        {
            var c = new Chart() { Size = new Size(1920, 1080) };
            c.Titles.Add("Normal vs Time for "+name);
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
            c.ChartAreas.Add(a);
            var t = new Series("Readings")
            {
                ChartType = SeriesChartType.FastLine,
                XValueType = ChartValueType.Double,
                YValueType = ChartValueType.Double,
                Color = Color.FromName("Blue")
            };
            var time = data.Select(x => x.time).ToList();
            var normal = data.Select(y => y.normal).ToList();
            t.Points.DataBindXY(time, normal);
            c.Series.Add(t);
            t.ChartArea = "Lather";

            var f = new Series("Fit")
            {
                ChartType = SeriesChartType.FastLine,
                XValueType = ChartValueType.Double,
                YValueType = ChartValueType.Double,
                Color = Color.FromName("Red")
            };
            time = fit.Select(x => x.time).ToList();
            normal = fit.Select(y => y.normal).ToList();
            f.Points.DataBindXY(time, normal);
            c.Series.Add(f);
            f.ChartArea = "Lather";

            var h = new Series("Manual")
            {
                ChartType = SeriesChartType.FastLine,
                XValueType = ChartValueType.Double,
                YValueType = ChartValueType.Double,
                Color = Color.FromName("Green")
            };
            time = man.Select(x => x.time).ToList();
            normal = man.Select(y => y.normal).ToList();
            h.Points.DataBindXY(time, normal);
            c.Series.Add(h);
            h.ChartArea = "Lather";

            Console.WriteLine("saving chart " + name);
            c.SaveImage(name + ".png", ChartImageFormat.Png);
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

            //if (testType != TestType.Lather)
            //    return;

            var f = Path.GetFileNameWithoutExtension(file);

            //if (!Issue3.Where(q => f.Contains(q)).Any())
            //    return;

            //Console.WriteLine(testType.ToString() + "\t" + Path.GetFileNameWithoutExtension(file));
            Console.WriteLine(testType.ToString() + "\t" + can(file));

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
                var data = lines.Skip(firstline).Take(rowmap[(int)TestType.Lather]).Select(s => new Reading(s)).Where(d => d.rate > 99.0 && d.rate < 101.0).ToList();

                var max = data.Max(d => d.normal);
                var ninf = data.Where(d => d.time > 20).Average(d => d.normal);
                var n2fit = data.Where(d => d.normal <= ((max + ninf) / 2.0) && d.time < 10);
                //var fit = data.Where(d => d.normal <= ((max + ninf) / 2.0) && d.time <= 10);

                //var y2f = data.Select(d => new Reading(d.time,  Math.Log(Math.Abs(d.normal - ninf))));
                var y2fit = n2fit.Select(d => Math.Log(Math.Abs(d.normal - ninf))).ToArray();
                var x2fit = n2fit.Select(d => d.time).ToArray();
                Tuple<double, double> p = mn.Fit.Line(x2fit, y2fit);   // item1 intercept, item2 slope

                var n0 = Math.Exp(p.Item1) + ninf;

                var fit = data.Select(d => new Reading(d.time, n0 + (ninf - n0) * (1 - Math.Exp(d.time * p.Item2))));
                //var NP = _t95[can(file)];
                //var man = data.Select(d => new Reading(d.time, NP[0] + (NP[1] - NP[0]) * (1 - Math.Exp(- d.time / NP[2]))));
                var n2 = n2fit.Sum(d => Math.Pow(d.normal - (n0 + (ninf - n0) * (1 - Math.Exp(d.time * p.Item2))), 2));
                //var g = mn.GoodnessOfFit.RSquared(x2fit.Select(x => p.Item1 + p.Item2 * x), y2fit); // == 1.0

                //ChartFile(can(file), data, fit.ToList(), man.ToList());

                // Create the model
                //if (model == null)
                //{
                //    context = SolverContext.GetContext();
                //    model = context.CreateModel();
                //    // Add a decisions
                //    N0 = new Decision(Domain.Real, "n0");
                //    TC = new Decision(Domain.Real, "tc");
                //    model.AddDecisions(N0);
                //    model.AddDecisions(TC);
                //}

                //Console.WriteLine("Log-linear regression estimates for "+ can(file));
                //Console.WriteLine("#: "+n2fit.Count()+"iN0: " + n0 + ", ninf: "+ninf+", p.Item1: "+p.Item1+", p.Item2: "+p.Item2);
                //Console.WriteLine("iTC: " + (-1 / p.Item2));
                //N0.SetInitialValue(n0);
                //TC.SetInitialValue(-1 / p.Item2);

                var tc = -1 / p.Item2;

                //var cost = new SumTermBuilder(n2fit.Count());

                //n2fit.ForEach(d =>
                //{
                //    Term r = N0 + (ninf - N0) * (1 - Model.Exp(-d.time / TC));
                //    r -= d.normal;
                //    r *= r;
                //    cost.Add(r);
                //});

                //model.AddGoal("Chi2", GoalKind.Minimize, cost.ToTerm());            // add goal

                ////var directive = new CompactQuasiNewtonDirective();
                ////var solver = context.Solve(directive);
                //var solver = context.Solve();
                //var report = solver.GetReport();
                //Console.Write(report);

                //var chi2 = model.Goals.First().ToDouble();

                //outrow.Cell(7).SetValue<double>(N0.GetDouble());
                //outrow.Cell(8).SetValue<double>(TC.GetDouble());
                //outrow.Cell(9).SetValue<double>(ninf);
                //outrow.Cell(10).SetValue<double>(TC.GetDouble() * t95);

                outrow.Cell(6).SetValue<double>(n2);
                outrow.Cell(7).SetValue<double>(n0);
                outrow.Cell(8).SetValue<double>(ninf);
                outrow.Cell(9).SetValue<double>(tc);
                outrow.Cell(10).SetValue<double>(tc * t95);

                //Console.WriteLine("--> Chi2: " + chi2 + ", N0: " + N0.GetDouble() + ", TC: " + TC.GetDouble());
                //model.RemoveGoal(model.Goals.First());                              // remove goal for next model run                
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
                //if (_can == "AG" || _can == "Z")
                //{
                //    Console.WriteLine("midtriple:");
                //    var rdg = midtriple.First();

                //    Console.Write(rdg.strain + " -- " + intersectx + " " + (rdg.strain < intersectx) + " < ");
                //    rdg.print();

                //    rdg = midtriple.ElementAt(1);
                //    Console.Write(rdg.strain + " -- " + intersectx + " " + (rdg.strain < intersectx) + " < ");
                //    rdg.print();

                //    rdg = midtriple.Last();
                //    Console.Write(rdg.strain + " -- " + intersectx + " " + (rdg.strain < intersectx) + " < ");
                //    rdg.print();

                //    Console.WriteLine("mid count: " + mid.Count());
                //    Console.WriteLine("ypt count: " + ypt.Count());
                //    //Console.WriteLine("denom: " + denom + ", strd: " + strd);
                //    Console.WriteLine("iline intercept: " + iline.intercept + ", slope: " + iline.slope);
                //    Console.WriteLine("mline intercept: " + mline.intercept + ", slope: " + mline.slope);

                //    //Console.WriteLine("gstrain: " + gstrain);
                //    //Console.WriteLine("gstress: " + gstress);
                //}
                //    Console.WriteLine(ypstress+" "+ypstrain);

                var bpstrain = (fline.intercept - mline.intercept) / (mline.slope - fline.slope);
                var bpt = readings.TakeWhile(s => s.strain < bpstrain);
                var bp = readings.Skip(bpt.Count() - 1).Take(2).ToArray();
                var bpstress = (bp[1].stress - bp[0].stress) / (bp[1].strain - bp[0].strain) * (bpstrain - bp[0].strain) + bp[0].stress;

                var cross = readings.Skip(1).TakeWhile(s => s.prime >= s.dprime);
                // old formula
                // var cr = readings.Skip(cross.Count() - 4).Take(2).ToArray();

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

                //Console.WriteLine("sam count: " + sam.Count());

                var start = sam.Any()?sam.Count(): 319;
                var avg = data.Skip(start).Take(101).Average(s => s.shear);
                var category = avg < max;
                //Console.WriteLine("avg max cat");
                //Console.WriteLine(avg + " " + max + " " + category);

                var at5 = data.Last(t => t.time <= 5.0).shear;
                 
                var span = data.Max(t => t.time);
                var at60 = span >= 60.0 ? data.First(t => t.time >= 60.0).shear : data.Last().shear;
                var delta = at60 - at5;

                outrow.Cell(23).SetValue<int>(category?1:0);
                outrow.Cell(24).SetValue<double>(at5);
                outrow.Cell(25).SetValue<double>(at60);
                outrow.Cell(26).SetValue<double>(delta);
                if (category) outrow.Cell(27).SetValue<double>(orig.Max(s => s.shear));

                //Console.WriteLine("cat at5 at60 delta peak");
                //Console.WriteLine(category + " " + at5 + " " + at60 + " " + delta + " " + span);
            }
        }
    }
}
