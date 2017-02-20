using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using xl = Microsoft.Office.Interop.Excel;

namespace ParseText
{
    class Lather : Test
    {
        static bool first = true;

        static double[] lastSolve = new double[]
        {
            double.NaN,
            double.NaN,
            double.NaN,
            double.NaN
        };

        private static xl.Application excel;
        private static xl.Workbooks wbs;
        private static xl.Workbook wb;
        private static xl.Worksheet ws;
        private static xl.Range start;
        private static xl.Range end;
        private static xl.Range rg;

        private static double t95 = -Math.Log(.05);

        public Lather()
        {
            targetrows = 144;
        }

        public override void Analyze()
        {
            object[,] arr = new object[targetrows, 4];

            List<Reading> setup = lines.Skip(firstline).Take(targetrows).Select((s, i) =>
            {
                var a = s.Split('\t');
                if (a.Length < 4)
                {
                    Program.form.WriteLine("file: "+Path.GetFileName(file));
                    Program.form.WriteLine("missing data at row " + i);
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
                end = ws.Cells[3 + targetrows, 5];
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
                Program.Series = new Dictionary<string, List<Reading>>();
                Program.Series["readings"] = setup.ToList();
                if (N0 < 1000.0)
                {
                    var fit = addedzero.Select(d => d == null ? null : new Reading(d.time, N0 + (Ninf - N0) * (1 - Math.Exp(-d.time / TC))));
                    Program.Series["fit"] = fit.ToList();
                }
                Program.postitle = " (chi2 = " + chi2.ToString("e3") + ")";
            }
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
