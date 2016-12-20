using System.Linq;

namespace ParseText
{
    class Cohesion : Test
    {
        public Cohesion()
        {
            targetrows = 200;
        }

        public override void Analyze()
        {
            var pairs = lines.Skip(firstline).Take(targetrows).Select(s =>
            {
                var a = s.Split('\t');
                return new { time = double.Parse(a[0]), normal = double.Parse(a[2]) };
            }).ToList();
            var min = pairs.First(b => b.normal == pairs.Min(a => a.normal));
            outrow.Cell(12).SetValue<double>(min.normal);
            outrow.Cell(13).SetValue<double>(min.time);
        }
    }
}
