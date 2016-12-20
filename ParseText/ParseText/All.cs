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
                             .Select(s => { last = s.num; return s.num - last - 4; }).ToList();

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
