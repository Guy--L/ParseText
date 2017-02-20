using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;

namespace ParseText
{
    static class TestFactory
    {
        private static List<int[]> rangemap = new List<int[]>()
        {
            new int[] { -2, -1 },
            new int[] { 0, 7 },
            new int[] { 130, 160 },
            new int[] { 190, 210 },
            new int[] { 400, 430 },
            new int[] { 30, 50 }
        };

        private static List<Func<Test>> test = new List<Func<Test>>()
        {
            () => new Trim(),               // line count not in one of the ranges
            () => new Trim(),               
            () => new Lather(),
            () => new Cohesion(),
            () => new Fract_Band(),
            () => new Oscillation(),
        };

        public static Test GetTest(string[] inlines)
        {
            var count = inlines.Count();
            return inlines.Any(n => n == "[step]") ? new All() : GetTest(count);
        }

        public static Test GetTest(int count)
        {
            var type = rangemap.Select((v, i) => new { range = v, index = i })
                .Where(r => r.range[0] <= count && count <= r.range[1])
                .Select(r => r.index).SingleOrDefault();
            return test[type]();
        }
    }

    class Test
    {
        protected int targetrows;

        public int Take(int c) => targetrows > c ? c : targetrows;

        public int firstline = 9;
        public string[] lines { get; set; }
        public IXLRow outrow { get; set; }
        public string file { get; set; }

        public Test() {}

        public virtual void Analyze() { }
    }
}
