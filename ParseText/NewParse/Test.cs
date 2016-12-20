using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

namespace NewParse
{
    class Test
    {
        private int firstline = 9;

        public enum TestType
        {
            Trim,
            Lather,
            Cohesion,
            Fract_Band,
            Oscillation,
            All,
            Error = -1
        };

        private static List<int[]> rangemap = new List<int[]>() {
            new int[] { 0, 3 },
            new int[] { 130, 160 },
            new int[] { 190, 210 },
            new int[] { 400, 430 },
            new int[] { 30, 50 }
        };

        /// <summary>
        /// Number of rows in text files that correspond to the test types above
        /// </summary>
        public static List<int> rowmap = new List<int>() { 1, 144, 200, 418, 38, 0, -1 };


        private readonly Dictionary<TestType, Action> analysis = new Dictionary<TestType, Action>();

        private TestType type { get; }
        private string[] lines { get; set; }
        private int rows => rowmap[(int)type];

        public static TestType typer(int c) => (TestType) rangemap
                                        .Select((v, i) => new { range = v, index = i })
                                        .Where(r => r.range[0] <= c && c <= r.range[1])
                                        .Select(r => r.index).SingleOrDefault();


        public Test(string[] buffer)
        {
            lines = buffer;
            var datalines = lines.Count() - firstline - 1;
            type = typer(datalines);
        }

        private static double t95 = -Math.Log(.05);

        static bool first = true;

        static double[] lastSolve = new double[]
        {
            double.NaN,
            double.NaN,
            double.NaN,
            double.NaN
        };

    }
}
