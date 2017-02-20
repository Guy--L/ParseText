using System;
using System.IO;
using System.Linq;
using fm = System.Windows.Forms;
using ClosedXML.Excel;
using System.Configuration;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;

namespace ParseText
{
    partial class Program
    {
        private static string _data;
        private static string _infileprefix = @"Rheology Form Filled In ";
        private static string _outfilename = @"{0} Rheology Analysis v3 with SPTT Entry Macro (MACRO v4.1) {1}";
        private static string _outdirectory;
        private static string _exportcmd;
        private static string _currentsample;

        public static Dictionary<string, List<Reading>> Series;
        public static string postitle;

        public static Form1 form;

        [STAThread]
        static void Main(string[] args)
        {
            _infileprefix = Properties.Settings.Default["infileprefix"].ToString();
            _outfilename = ConfigurationManager.AppSettings["outfilename"];
            _outdirectory = ConfigurationManager.AppSettings["outdirectory"];
            _exportcmd = ConfigurationManager.AppSettings["exportcmd"];

            if (args.Length == 0) {
                args = new string[] { "." };
            }

            // look for request XLs in all directories on command line

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

            string[] docs = null;

            try
            {
                docs = Directory.GetFiles(MyDir, "*.xlsm");
            }
            catch(Exception e)
            {
                form.WriteLine(e.Message);
                return;
            }

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
                filemask = sample + "*.tri";
                samplefiles = Directory.GetFiles(dir, filemask);
            }
            if (samplefiles.Count() == 0)
            {
                filemask = sample + "*.txt";
                filemask = filemask.Insert(_currentsample.Length, "-");
                samplefiles = Directory.GetFiles(dir, filemask);
            }
            //Console.WriteLine(samplefiles.Count() + " samples in " + dir + " " + filemask);
            return samplefiles.ToLookup(k => can(k).ToUpper(), v => v);
        }

        private static IXLWorksheet _outsh;
        private static string _can;

        public static string outpath = "";
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
                var blank = string.IsNullOrWhiteSpace(sample);

                if (outrowi == 0 && !blank)
                    outrowi = row.RowNumber() + 2;
                else {
                    if (blank) continue;
                    outrowi++;
                }

                var outrow = _outsh.Row(outrowi);
                if (samplename != sample && !blank)
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
                    var refile = ExportFile(file);
                    ReadFile(refile, outrow);
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

        static cmdShell cmd = null;

        static void test_onData(cmdShell sender, string e)
        {
            Debug.WriteLine("d "+e);
        }

        static void onChanged(object source, FileSystemEventArgs e)
        {
            // Specify what is done when a file is changed, created, or deleted.
            Debug.WriteLine("File: " + e.FullPath + " " + e.ChangeType);
            lock (sync)
            {
                Monitor.Pulse(sync);
                Debug.WriteLine("pulsed");
            }
        }

        static void onCreated(object source, FileSystemEventArgs e)
        {
            Debug.WriteLine("+ ");
            lock (sync)
            {
                Monitor.Pulse(sync);
            }
        }

        static FileSystemWatcher watch = null;
        static object sync = null;

        static string ExportFile(string file)
        {
            if (Path.GetExtension(file).ToLower() != ".tri")
                return file;

            var dir = Path.GetDirectoryName(file);

            if (sync == null) sync = new object();
            if (cmd == null)
            {
                cmd = new cmdShell();
                cmd.onData += test_onData;
            }

            if (watch == null)
            {
                watch = new FileSystemWatcher();
                watch.Changed += onChanged;
            }
            watch.Path = dir;
            watch.Filter = "*.txt";
            watch.EnableRaisingEvents = true;

            cmd.writewait(string.Format(_exportcmd, file), sync);
            
            return Path.ChangeExtension(file, "txt");
        }


        //private static List<string> Issue3 = new List<string> {
        //    "G-0033f",
        //    "L-0058f"
        //};

        static void ReadFile(string file, IXLRow outrow)
        {
            var lines = File.ReadAllLines(file);
            var count = lines.Count() - 10;

            var test = TestFactory.GetTest(lines);
            test.outrow = outrow;
            test.lines = lines;

            Debug.WriteLine($"file {file}");
            test.file = file;
            test.Analyze();

            ChartSeries(can(file) + postitle, Series);      // runs only if Series is non-null from Lather tests
            Series = null;
        }
    }
}
