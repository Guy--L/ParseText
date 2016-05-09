using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using fm = System.Windows.Forms;
using ClosedXML.Excel;
using System.Configuration;
using MathNet.Numerics.Statistics;
using xl = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace ParseText
{
    partial class Program
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

        private static Dictionary<string, double[]> _t95man = new Dictionary<string, double[]>();
        public static List<double>[] _t95err = new List<double>[5]
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

            fm.Application.EnableVisualStyles();
            fm.Application.SetCompatibleTextRenderingDefault(false);

            form = new Form1();
            fm.Application.Run(form);
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

        static string outpath = "";
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

            _currentsample = request[2];

            var outfilename = string.Format(_outfilename, string.Join(" ", request.Take(2)), request[2]) + ".xlsm";
            outpath = form.notoutset ? data : form.outdir;
            var outfile = Path.Combine(outpath, outfilename);

            //form.WriteLine("writing to " + outfile);

            var outxl = new XLWorkbook("AnalysisTemplate.xlsm");
            _outsh = outxl.Worksheet("Summary Table");
            string samplename = "";
            ILookup<string, string> samples = null;

            int outrowi = 0;
            //form.WriteLine(string.Format("{0,6} {1,6} {2,6} {3,6} {4,6}", "N^2", "N0", "Ninf", "TC", "can"));

            foreach (var row in insh.Rows())
            {
                if (row.RowNumber() < 2)
                {
                    _outsh.FirstRow().Cell(1).Value = "Request " + request[1] + " - ";
                    _outsh.FirstRow().Cell(3).Value = request[2];
                    continue;
                }
                string sample = row.Cell(4).GetValue<string>();
                outrowi = row.RowNumber() + 2;
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
            _outsh.Range(outrowi + 1, 1, 64, 27).Clear(XLClearOptions.Formats);
            _outsh.Columns(3, 4).AdjustToContents();
            IXLColumn col1 = _outsh.Column(3);
            IXLColumn col2 = _outsh.Column(4);

            col1.Width *= 1.1;
            col2.Width *= 1.1;

            outxl.SaveAs(outfile);
            outxl.Dispose();

            //form.WriteLine("saved as " + outfilename);
        }

        private static int initialskip = 4;
        private static int initialtake = 5;
        private static int finalskip = 32;
        private static int finaltake = 3;
        private static double t95 = -Math.Log(.05);

        public static void Release()
        {
            releaseObject(rg);
            releaseObject(start);
            releaseObject(end);
            releaseObject(ws);
            if (wb != null) wb.Close(true, Type.Missing, Type.Missing);
            releaseObject(wb);
            releaseObject(wbs);
            if (excel != null)
            {
                excel.Application.Quit();
                excel.Quit();
            }
            releaseObject(excel);
        }

        private static xl.Application excel;
        private static xl.Workbooks wbs;
        private static xl.Workbook wb;
        private static xl.Worksheet ws;
        private static xl.Range start;
        private static xl.Range end;
        private static xl.Range rg;

        private static void releaseObject(object obj)
        {
            if (obj == null)
                return;

            try
            {
                Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        //private static List<string> Issue3 = new List<string> {
        //    "G-0033f",
        //    "L-0058f"
        //};

        static bool first = true;

        static double[] lastSolve = new double[]
        {
            double.NaN,
            double.NaN,
            double.NaN,
            double.NaN
        };

        static void ReadFile(string file, IXLRow outrow)
        {
            var lines = File.ReadAllLines(file);
            var datalines = lines.Count() - firstline - 1;
            TestType testType = testmap[datalines];

            var f = Path.GetFileNameWithoutExtension(file);
            //form.WriteLine(testType.ToString() + "\t" + can(file));

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
                int rows = rowmap[(int)TestType.Lather];
                object[,] arr = new object[rows, 4];

                List<Reading> setup = lines.Skip(firstline).Take(rowmap[(int)TestType.Lather]).Select((s, i) =>
                {
                    var a = s.Split('\t');
                    if (a.Length < 4)
                    {
                        Debug.WriteLine("no data at i: " + i);
                        return null;
                    }
                    arr[i, 0] = double.Parse(a[0]);
                    arr[i, 1] = double.Parse(a[1]);
                    arr[i, 2] = double.Parse(a[2]);
                    arr[i, 3] = double.Parse(a[3]);
                    return new Reading(s);
                }).ToList();

                bool rc = false;
                if (excel == null)
                {
                    excel = new xl.Application();
                    var template = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Solver.xlsm");
                    wbs = excel.Workbooks;
                    wb = wbs.Open(template);
                    ws = wb.Sheets["Automation"];
                    ws.Activate();
                    start = ws.Cells[4, 2];
                    end = ws.Cells[3 + rows, 5];
                    rg = ws.get_Range(start, end);
                    rc = excel.Run("CheckSolver");
                    form.WriteLine("Solver installed ok? " + rc);
                }

                rg.Value = arr;
                ws.Calculate();

                excel.Run("RunSolver");

                double chi2 = ws.Cells[3, 13].Value;
                double N0 = ws.Cells[4, 13].Value;
                double Ninf = ws.Cells[5, 13].Value;
                double TC = ws.Cells[6, 13].Value;

                var same = (lastSolve[0] == chi2 && lastSolve[1] == N0 && lastSolve[2] == Ninf && lastSolve[3] == TC);
                lastSolve[0] = chi2;
                lastSolve[1] = N0;
                lastSolve[2] = Ninf;
                lastSolve[3] = TC;

                var solveCode = ws.Cells[12, 13].Value;

                excel.Run("FinishSolver");
                //form.WriteLine(string.Format("{0,6:N2} {1,6:N2} {2,6:N2} {3,6:N2} {4,6}", chi2, N0, Ninf, TC, can(file)));

                if (solveCode == 9.0 && first)
                {
                    form.WriteLine("Solve result --> " + solveCode);
                    first = false;
                    var tstout = Path.Combine(outpath, "SolverErr.xlsm");
                    ws.SaveAs(tstout);
                }

                outrow.Cell(6).SetValue(chi2);
                outrow.Cell(7).SetValue(N0);
                outrow.Cell(8).SetValue(Ninf);
                outrow.Cell(9).SetValue(TC);
                outrow.Cell(10).SetValue(TC * t95);

                var addedzero = (new List<Reading>() { new Reading(0, N0) }).Concat(setup);

                //model.RemoveGoal(model.Goals.First());                              // remove goal for next model run       

                if (form.doCharts)
                {
                    //var xlv = _t95man[can(file)];
                    //var xlf = addedzero.Select(d => new Reading(d.time, xlv[1] + (xlv[2] - xlv[1]) * (1 - Math.Exp(-d.time / xlv[3]))));
                    Dictionary<string, List<Reading>> Series = new Dictionary<string, List<Reading>>();
                    Series["readings"] = setup.ToList();
                    if (N0 < 1000.0)
                    {
                        var fit = addedzero.Select(d => d==null?null: new Reading(d.time, N0 + (Ninf - N0) * (1 - Math.Exp(-d.time / TC))));
                        Series["fit"] = fit.ToList();
                    }

                    var title = can(file) + " (chi2 = " + chi2.ToString("e3") + ")";
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
