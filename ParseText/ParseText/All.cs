using System.Diagnostics;
using System.Linq;

namespace ParseText
{
    class All : Test
    {
        public All()
        {
            targetrows = 0;
        }

        public override void Analyze()
        {
            var last = 0;
            var count = lines.Select((s, i) => new { line = s, num = i })
                             .Where(s => s.line == "[step]")
                             .Select(s => { var cnt = s.num - last - 4; last = s.num; return cnt; }).ToList();

            count.Add(lines.Count() - last - 4);
            var pos = 4;
            foreach (var type in count)
            {
                var test = TestFactory.GetTest(type);
                test.outrow = outrow;
                test.firstline = 4;

                test.lines = lines.Skip(pos).Take(test.Take(type - 1)).ToArray();
                test.Analyze();

                pos += type + 4;
            }
        }
    }
}
