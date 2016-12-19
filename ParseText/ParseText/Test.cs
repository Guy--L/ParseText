using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using ClosedXML.Excel;
using xl = Microsoft.Office.Interop.Excel;

namespace ParseText
{
    class Test
    {
        const int firstline = 9;

        private enum TestType
        {
            Trim,
            Lather,
            Cohesion,
            Fract_Band,
            Oscillation,
            All,
            Error = -1
        };

        /// <summary>
        /// Number of rows in text files that correspond to the test types above
        /// </summary>
        private static List<int> rowmap = new List<int>() { 1, 144, 200, 418, 38, 0, 0 };
        private static Dictionary<int, TestType> testmap = rowmap
            .Select((v, i) => new { value = v, index = i })
            .ToDictionary(v => v.value, k => (TestType)k.index);

        private Dictionary<TestType, Action> analysis = new Dictionary<TestType, Action>();

        private TestType type { get; }
        private string[] lines { get; set; }
        private IXLRow outrow { get; set; }
        private int rows => rowmap[(int)type];

        public Dictionary<string, List<Reading>> Series;
        public string postitle;

        static Test()
        {
            testmap.Add(143, TestType.Lather);
            testmap.Add(142, TestType.Lather);
            testmap.Add(145, TestType.Lather);
            testmap.Add(417, TestType.Fract_Band);
        }

        public Test(string[] buffer, IXLRow xlrow)
        {
            lines = buffer;
            outrow = xlrow;
            var datalines = lines.Count() - firstline - 1;
            type = testmap.ContainsKey(datalines)?testmap[datalines]:TestType.All;

            analysis.Add(TestType.Lather, Lather);
            analysis.Add(TestType.Cohesion, Cohesion);
            analysis.Add(TestType.Fract_Band, Fract_Band);
            analysis.Add(TestType.Oscillation, Oscillation);

            postitle = "";
            Series = null;
        }

        public void Analyze()
        {
            analysis[type].Invoke();
        }

        private static xl.Application excel;
        private static xl.Workbooks wbs;
        private static xl.Workbook wb;
        private static xl.Worksheet ws;
        private static xl.Range start;
        private static xl.Range end;
        private static xl.Range rg;

        private static double t95 = -Math.Log(.05);

        static bool first = true;

        static double[] lastSolve = new double[]
        {
            double.NaN,
            double.NaN,
            double.NaN,
            double.NaN
        };

        public void Lather()
        {
            object[,] arr = new object[rows, 4];

            List<Reading> setup = lines.Skip(firstline).Take(rows).Select((s, i) =>
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
                Program.form.WriteLine("Solver installed ok? " + rc);
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
                Program.form.WriteLine("Solve result --> " + solveCode);
                first = false;
                var tstout = Path.Combine(Program.outpath, "SolverErr.xlsm");
                ws.SaveAs(tstout);
            }

            outrow.Cell(6).SetValue(chi2);
            outrow.Cell(7).SetValue(N0);
            outrow.Cell(8).SetValue(Ninf);
            outrow.Cell(9).SetValue(TC);
            outrow.Cell(10).SetValue(TC * t95);

            var addedzero = (new List<Reading>() { new Reading(0, N0) }).Concat(setup);

            //model.RemoveGoal(model.Goals.First());                              // remove goal for next model run       

            if (Program.form.doCharts)
            {
                //var xlv = _t95man[can(file)];
                //var xlf = addedzero.Select(d => new Reading(d.time, xlv[1] + (xlv[2] - xlv[1]) * (1 - Math.Exp(-d.time / xlv[3]))));
                Series = new Dictionary<string, List<Reading>>();
                Series["readings"] = setup.ToList();
                if (N0 < 1000.0)
                {
                    var fit = addedzero.Select(d => d == null ? null : new Reading(d.time, N0 + (Ninf - N0) * (1 - Math.Exp(-d.time / TC))));
                    Series["fit"] = fit.ToList();
                }
                postitle = " (chi2 = " + chi2.ToString("e3") + ")";
            }
        }

        public void Cohesion()
        {
            var pairs = lines.Skip(firstline).Take(rows).Select(s =>
            {
                var a = s.Split('\t');
                return new { time = double.Parse(a[0]), normal = double.Parse(a[2]) };
            }).ToList();
            var min = pairs.First(b => b.normal == pairs.Min(a => a.normal));
            outrow.Cell(12).SetValue<double>(min.normal);
            outrow.Cell(13).SetValue<double>(min.time);
        }

        public void Fract_Band()
        {
            var orig = lines.Skip(firstline).Take(rows).Select(s => new Reading(s)).ToList();
            var data = orig.Select(d => (d.rate >= 0.95 && d.rate <= 1.05) ? d : new Reading(d));

            var max = data.Take(319).Max(s => s.shear);
            var sam = data.Take(319).TakeWhile(s => s.shear < max);

            var start = sam.Any() ? sam.Count() : 319;
            var avg = data.Skip(start).Take(101).Average(s => s.shear);
            var category = avg < max;

            var at5 = data.Last(t => t.time <= 5.0).shear;

            var span = data.Max(t => t.time);
            var at60 = span >= 60.0 ? data.First(t => t.time >= 60.0).shear : data.Last().shear;
            var delta = at60 - at5;

            outrow.Cell(23).SetValue<int>(category ? 1 : 0);
            outrow.Cell(24).SetValue<double>(at5);
            outrow.Cell(25).SetValue<double>(at60);
            outrow.Cell(26).SetValue<double>(delta);
            if (category) outrow.Cell(27).SetValue<double>(orig.Max(s => s.shear));

        }

        private static int initialskip = 4;
        private static int initialtake = 5;
        private static int finalskip = 32;
        private static int finaltake = 3;

        public void Oscillation()
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

        public static void Release()
        {
            releaseObject(rg);
            releaseObject(start);
            releaseObject(end);
            releaseObject(ws);
            if (wb != null) wb.Close(false, Type.Missing, Type.Missing);
            releaseObject(wb);
            releaseObject(wbs);
            if (excel != null)
            {
                excel.Application.Quit();
                excel.Quit();
            }
            releaseObject(excel);

        }

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
                Program.form.WriteLine("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
